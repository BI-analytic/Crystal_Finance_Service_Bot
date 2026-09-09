#     pip install pyTelegramBotAPI google-genai python-dotenv openpyxl


import os
import re
import sys
import json
import sqlite3
import hashlib
import tempfile
from datetime import datetime

import telebot
from google import genai
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from dotenv import load_dotenv

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
MODEL = os.getenv("MODEL", "gemini-3.5-flash")

# Кто может пользоваться ботом: id через запятую. Пустое значение = никто.
ALLOWED_USERS = {
    int(x) for x in os.getenv("ALLOWED_USERS", "").replace(" ", "").split(",") if x
}

# Показывать колонку с пояснением: да / нет
SHOW_COMMENT = os.getenv("SHOW_COMMENT", "да").strip().lower() in ("да", "yes", "true", "1")

# Сколько строк отправлять в модель за один запрос
BATCH_SIZE = int(os.getenv("BATCH_SIZE", "15"))

CATALOG_FILE = "catalog.xlsx"
PROMPT_FILE = "prompt.md"
CACHE_FILE = "cache.db"

if not BOT_TOKEN or BOT_TOKEN.startswith("вставь"):
    sys.exit("Не найден BOT_TOKEN — проверь .env")
if not GEMINI_API_KEY or GEMINI_API_KEY.startswith("вставь"):
    sys.exit("Не найден GEMINI_API_KEY — проверь .env")

bot = telebot.TeleBot(BOT_TOKEN)
client = genai.Client(api_key=GEMINI_API_KEY)


# ---------------------------------------------------------
#  2. Справочник и промпт
# ---------------------------------------------------------
def load_catalog(path=CATALOG_FILE):
    """Читает catalog.xlsx. Возвращает список активных статей."""
    if not os.path.exists(path):
        sys.exit(f"Не найден {path} — положи его рядом с bot.py")

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Справочник"]

    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = str(row[0]).strip() if row[0] else ""
        name = str(row[1]).strip() if row[1] else ""
        desc = str(row[2]).strip() if row[2] else ""
        use = str(row[3]).strip().lower() if len(row) > 3 and row[3] else "да"
        if name and use in ("да", "yes", "true", "1"):
            items.append({"code": code, "name": name, "desc": desc})
    wb.close()

    if not items:
        sys.exit("В catalog.xlsx нет активных статей — проверь колонку «Использовать»")
    return items


def catalog_to_text(items):
    """Превращает справочник в текст для промпта."""
    lines = []
    for it in items:
        desc = it["desc"] or "(описание не заполнено)"
        lines.append(f"[{it['code']}] {it['name']}\n    {desc}")
    return "\n".join(lines)


def build_prompt():
    """Собирает системный промпт: prompt.md + справочник."""
    if not os.path.exists(PROMPT_FILE):
        sys.exit(f"Не найден {PROMPT_FILE} — положи его рядом с bot.py")
    with open(PROMPT_FILE, encoding="utf-8") as f:
        template = f.read()

    items = load_catalog()
    text = catalog_to_text(items)
    version = hashlib.sha256(text.encode()).hexdigest()[:12]
    return template.replace("{catalog}", text), items, version


SYSTEM_PROMPT, CATALOG, CATALOG_VERSION = build_prompt()


# ---------------------------------------------------------
#  3. Кеш разборов (он же накопительная база исследований)
# ---------------------------------------------------------
def init_cache():
    con = sqlite3.connect(CACHE_FILE)
    con.execute("""
        CREATE TABLE IF NOT EXISTS analysis (
            key             TEXT PRIMARY KEY,
            item            TEXT,
            document        TEXT,
            content         TEXT,
            amount          TEXT,
            score           INTEGER,
            recommend_code  TEXT,
            recommend_name  TEXT,
            comment         TEXT,
            model           TEXT,
            catalog_version TEXT,
            created_at      TEXT
        )
    """)
    con.commit()
    return con


CACHE = init_cache()


def cache_key(row):
    """Один и тот же расход при том же справочнике даёт тот же ключ."""
    raw = "|".join([
        CATALOG_VERSION,
        str(row["item"]), str(row["document"]),
        str(row["content"]), str(row["amount"]),
    ])
    return hashlib.sha256(raw.encode()).hexdigest()


def cache_get(row):
    cur = CACHE.execute(
        "SELECT score, recommend_code, recommend_name, comment FROM analysis WHERE key = ?",
        (cache_key(row),),
    )
    r = cur.fetchone()
    if not r:
        return None
    return {"score": r[0], "recommend_code": r[1], "recommend_name": r[2], "comment": r[3]}


def cache_put(row, result):
    CACHE.execute(
        "INSERT OR REPLACE INTO analysis VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
        (
            cache_key(row), row["item"], row["document"], row["content"], str(row["amount"]),
            result["score"], result["recommend_code"], result["recommend_name"],
            result["comment"], MODEL, CATALOG_VERSION, datetime.now().isoformat(timespec="seconds"),
        ),
    )
    CACHE.commit()


# ---------------------------------------------------------
#  4. Чтение файла пользователя
# ---------------------------------------------------------
COLUMN_HINTS = {
    "item": ("стать",),
    "document": ("документ",),
    "content": ("содержан",),
    "amount": ("сумма",),
}


def find_columns(header):
    """Ищет нужные колонки по ключевым словам — заголовки могут отличаться."""
    found = {}
    for idx, title in enumerate(header):
        low = str(title or "").strip().lower()
        for field, hints in COLUMN_HINTS.items():
            if field not in found and any(h in low for h in hints):
                found[field] = idx
    return found


def read_input(path):
    """Возвращает (заголовки, строки). Ошибку бросает с понятным текстом."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    all_rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()

    header_idx, cols = None, None
    for i, r in enumerate(all_rows[:15]):
        c = find_columns(r)
        if len(c) == 4:
            header_idx, cols = i, c
            break

    if cols is None:
        raise ValueError(
            "Не нашёл нужные колонки. В файле должны быть заголовки со словами "
            "«статья», «документ», «содержание» и «сумма»."
        )

    header = [str(x or "").strip() for x in all_rows[header_idx]]
    rows = []
    for r in all_rows[header_idx + 1:]:
        item = str(r[cols["item"]] or "").strip() if cols["item"] < len(r) else ""
        if not item:
            continue
        rows.append({
            "item": item,
            "document": str(r[cols["document"]] or "").strip(),
            "content": str(r[cols["content"]] or "").strip(),
            "amount": r[cols["amount"]] if cols["amount"] < len(r) else None,
        })

    if not rows:
        raise ValueError("Файл прочитан, но строк с данными в нём нет.")
    return header, cols, rows


# ---------------------------------------------------------
#  5. Обращение к модели
# ---------------------------------------------------------
def parse_json(text):
    """Модель иногда оборачивает JSON в ```json ... ``` — снимаем обёртку."""
    text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    return json.loads(text)


def analyze_batch(batch):
    """Отправляет пачку строк в модель. Возвращает список результатов."""
    payload = [
        {
            "n": i + 1,
            "Статья": r["item"],
            "Документ": r["document"],
            "Содержание": r["content"],
            "Сумма": str(r["amount"]),
        }
        for i, r in enumerate(batch)
    ]

    user_text = (
        f"Проверь {len(batch)} строк расшифровки. "
        "Верни массив JSON, ровно один объект на строку.\n\n"
        + json.dumps(payload, ensure_ascii=False, indent=1)
    )

    response = client.models.generate_content(
        model=MODEL,
        contents=user_text,
        config={
            "system_instruction": SYSTEM_PROMPT,
            "response_mime_type": "application/json",
            "temperature": 0,
        },
    )
    data = parse_json(response.text)

    # раскладываем ответы по номерам строк
    by_n = {int(d.get("n", 0)): d for d in data if isinstance(d, dict)}
    results = []
    for i, row in enumerate(batch, start=1):
        d = by_n.get(i, {})
        results.append(normalize(d, row))
    return results


def normalize(d, row):
    """Приводит ответ модели к нужному виду и применяет правило Score."""
    try:
        score = int(d.get("score", 5))
    except (TypeError, ValueError):
        score = 5
    score = max(0, min(10, score))

    rec_code = str(d.get("recommend_code") or "").strip()
    rec_name = str(d.get("recommend_name") or "").strip()
    comment = str(d.get("comment") or "").strip()

    if not rec_name:
        rec_code = ""
        comment = "Модель не вернула ответ по этой строке — проверьте вручную."
        score = 5

    # Правило из ТЗ: статья не совпала — Score не больше 5.
    # Проверяем кодом, потому что модель это правило иногда нарушает.
    if norm(rec_name) != norm(row["item"]):
        score = min(score, 5)

    return {
        "score": score,
        "recommend_code": rec_code,
        "recommend_name": rec_name,
        "comment": comment,
    }


def norm(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower()


# ---------------------------------------------------------
#  6. Сборка результата
# ---------------------------------------------------------
ARIAL = "Arial"
FILL_OK = PatternFill("solid", fgColor="C6EFCE")    # 6-10
FILL_WARN = PatternFill("solid", fgColor="FFEB9C")  # 3-5
FILL_BAD = PatternFill("solid", fgColor="FFC7CE")   # 0-2


def write_output(rows, results, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "P&L Cost Analysis"

    headers = [
        "Наименование статьи бюджета",
        "Документ (наименование, №, дата документа)",
        "Содержание документа",
        "Сумма",
        "Score",
        "Recommend",
    ]
    if SHOW_COMMENT:
        headers.append("Комментарий")

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(name=ARIAL, size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for i, (row, res) in enumerate(zip(rows, results), start=2):
        rec = f"{res['recommend_code']} {res['recommend_name']}".strip()
        values = [row["item"], row["document"], row["content"], row["amount"],
                  res["score"], rec]
        if SHOW_COMMENT:
            values.append(res["comment"])

        for c, v in enumerate(values, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.font = Font(name=ARIAL, size=10)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(c in (3, 7)))

        ws.cell(row=i, column=4).number_format = "#,##0.00"
        s = ws.cell(row=i, column=5)
        s.alignment = Alignment(horizontal="center", vertical="center")
        s.fill = FILL_OK if res["score"] >= 6 else FILL_WARN if res["score"] >= 3 else FILL_BAD

    widths = {"A": 34, "B": 46, "C": 60, "D": 14, "E": 8, "F": 34, "G": 55}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{chr(64 + len(headers))}{len(rows) + 1}"

    wb.save(path)


# ---------------------------------------------------------
#  7. Телеграм: доступ и команды
# ---------------------------------------------------------
def allowed(message):
    return message.from_user.id in ALLOWED_USERS


def deny(message):
    bot.reply_to(
        message,
        "Доступ к боту ограничен.\n"
        f"Ваш Telegram ID: {message.from_user.id}\n"
        "Передайте его администратору для подключения.",
    )


@bot.message_handler(commands=["start", "help"])
def cmd_start(message):
    if not allowed(message):
        return deny(message)
    bot.reply_to(
        message,
        "Финансовая служба Кристалл — раздел P&L_Cost_Analysis\n\n"
        "Пришлите Excel с расшифровкой затрат по одному ЦФО.\n"
        "Нужные колонки: статья бюджета, документ, содержание, сумма.\n\n"
        "В ответ придёт тот же файл с колонками Score и Recommend.\n\n"
        "/reload — перечитать справочник\n"
        "/stats — что накоплено в базе\n"
        "/whoami — узнать свой Telegram ID",
    )


@bot.message_handler(commands=["whoami"])
def cmd_whoami(message):
    bot.reply_to(message, f"Ваш Telegram ID: {message.from_user.id}")


@bot.message_handler(commands=["reload"])
def cmd_reload(message):
    if not allowed(message):
        return deny(message)
    global SYSTEM_PROMPT, CATALOG, CATALOG_VERSION
    try:
        SYSTEM_PROMPT, CATALOG, CATALOG_VERSION = build_prompt()
        bot.reply_to(message, f"Справочник обновлён: {len(CATALOG)} активных статей.")
    except Exception as e:
        bot.reply_to(message, f"Не удалось перечитать справочник: {e}")


@bot.message_handler(commands=["stats"])
def cmd_stats(message):
    if not allowed(message):
        return deny(message)
    total = CACHE.execute("SELECT COUNT(*) FROM analysis").fetchone()[0]
    low = CACHE.execute("SELECT COUNT(*) FROM analysis WHERE score <= 5").fetchone()[0]
    bot.reply_to(
        message,
        f"В базе разборов: {total} строк\n"
        f"Из них со спорной статьёй (Score ≤ 5): {low}\n"
        f"Активных статей в справочнике: {len(CATALOG)}",
    )


# ---------------------------------------------------------
#  8. Основной сценарий: пришёл файл
# ---------------------------------------------------------
@bot.message_handler(content_types=["document"])
def handle_document(message):
    if not allowed(message):
        return deny(message)

    name = message.document.file_name or ""
    if not name.lower().endswith((".xlsx", ".xlsm")):
        return bot.reply_to(message, "Нужен файл Excel (.xlsx). Пришлите расшифровку в этом формате.")

    status = bot.reply_to(message, "Файл получен, читаю...")
    tmpdir = tempfile.mkdtemp()
    src = os.path.join(tmpdir, "input.xlsx")

    try:
        info = bot.get_file(message.document.file_id)
        with open(src, "wb") as f:
            f.write(bot.download_file(info.file_path))

        _, _, rows = read_input(src)
    except Exception as e:
        return bot.edit_message_text(f"Не смог прочитать файл.\n{e}", status.chat.id, status.message_id)

    total = len(rows)
    bot.edit_message_text(f"Строк на проверку: {total}. Начинаю анализ...",
                          status.chat.id, status.message_id)

    results = [None] * total
    pending = []          # строки, которых нет в кеше
    from_cache = 0

    for i, row in enumerate(rows):
        hit = cache_get(row)
        if hit:
            results[i] = hit
            from_cache += 1
        else:
            pending.append(i)

    done = 0
    try:
        for start in range(0, len(pending), BATCH_SIZE):
            idxs = pending[start:start + BATCH_SIZE]
            batch = [rows[i] for i in idxs]
            batch_results = analyze_batch(batch)

            for i, res in zip(idxs, batch_results):
                results[i] = res
                cache_put(rows[i], res)

            done += len(idxs)
            bot.edit_message_text(
                f"Обработано {done + from_cache} из {total}...",
                status.chat.id, status.message_id,
            )
    except Exception as e:
        return bot.edit_message_text(
            f"Ошибка при обращении к модели после {done + from_cache} из {total} строк.\n{e}\n\n"
            "Разобранные строки сохранены — пришлите файл ещё раз, "
            "бот продолжит с того места.",
            status.chat.id, status.message_id,
        )

    out = os.path.join(tmpdir, f"P&L_analysis_{datetime.now():%Y-%m-%d_%H%M}.xlsx")
    write_output(rows, results, out)

    low = sum(1 for r in results if r["score"] <= 5)
    avg = sum(r["score"] for r in results) / total

    bot.edit_message_text(f"Готово. Обработано {total} строк.", status.chat.id, status.message_id)
    with open(out, "rb") as f:
        bot.send_document(
            message.chat.id, f,
            caption=(
                f"Строк: {total}\n"
                f"Спорных (Score ≤ 5): {low}\n"
                f"Средний Score: {avg:.1f}\n"
                f"Взято из кеша: {from_cache}"
            ),
        )


@bot.message_handler(func=lambda m: True, content_types=["text"])
def fallback(message):
    if not allowed(message):
        return deny(message)
    bot.reply_to(message, "Пришлите Excel-файл с расшифровкой затрат. Подсказка: /help")


# ---------------------------------------------------------
#  9. Запуск
# ---------------------------------------------------------
if __name__ == "__main__":
    print(f"Справочник: {len(CATALOG)} активных статей (версия {CATALOG_VERSION})")
    print(f"Доступ разрешён: {len(ALLOWED_USERS)} пользователям")
    print(f"Колонка с комментарием: {'включена' if SHOW_COMMENT else 'выключена'}")
    print("Бот запущен.")
    bot.infinity_polling()
