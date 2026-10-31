
import os
import csv
import json
import time
import uuid
import sqlite3
import urllib.request
import urllib.parse
import textwrap
import html
from datetime import datetime
from openpyxl import Workbook, load_workbook

# =====================================
# MINA OFFICE BOT ULTIMATE V1
# =====================================

TOKEN = "9193727272:AAH9EHwtfhETo704P1hmqbas-8QFBkRGdi0"
ADMIN_ID = 8138427580  # ضع رقم حسابك هنا

API = f"https://api.telegram.org/bot{TOKEN}/"
FILE_API = f"https://api.telegram.org/file/bot{TOKEN}/"

MAX_DOWNLOAD = 10 * 1024 * 1024
TEMP_DIR = "temp_files"
DB_FILE = "mina_office.db"

os.makedirs(TEMP_DIR, exist_ok=True)

db = sqlite3.connect(DB_FILE, check_same_thread=False)

db.execute("""
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY,
    points INTEGER DEFAULT 0,
    requests INTEGER DEFAULT 0,
    joined TEXT
)
""")

db.execute("""
CREATE TABLE IF NOT EXISTS history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    operation TEXT,
    created_at TEXT
)
""")

db.commit()

user_modes = {}


# =====================================
# DATABASE
# =====================================

def register_user(uid):
    db.execute("""
    INSERT OR IGNORE INTO users(user_id, joined)
    VALUES (?, ?)
    """, (uid, datetime.now().strftime("%Y-%m-%d %H:%M")))

    db.commit()


def record_operation(uid, operation):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    db.execute("""
    UPDATE users
    SET points = points + 1,
        requests = requests + 1
    WHERE user_id = ?
    """, (uid,))

    db.execute("""
    INSERT INTO history(user_id, operation, created_at)
    VALUES (?, ?, ?)
    """, (uid, operation, now))

    db.commit()


# =====================================
# TELEGRAM API
# =====================================

def api(method, data=None):
    body = urllib.parse.urlencode(data or {}).encode()

    request = urllib.request.Request(
        API + method,
        data=body
    )

    with urllib.request.urlopen(request, timeout=40) as response:
        result = json.loads(response.read().decode())

    if not result.get("ok"):
        raise Exception(result.get("description", "Telegram API error"))

    return result["result"]


def send(chat, text, keyboard=None):
    data = {
        "chat_id": chat,
        "text": str(text)
    }

    if keyboard:
        data["reply_markup"] = json.dumps(
            keyboard,
            ensure_ascii=False
        )

    api("sendMessage", data)


def send_document(chat, path, caption="تم تجهيز الملف"):
    boundary = "----MinaOffice" + uuid.uuid4().hex

    with open(path, "rb") as file:
        file_data = file.read()

    filename = os.path.basename(path)

    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="chat_id"\r\n\r\n'
        f"{chat}\r\n"
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="caption"\r\n\r\n'
        f"{caption}\r\n"
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="document"; filename="{filename}"\r\n'
        f"Content-Type: application/octet-stream\r\n\r\n"
    ).encode() + file_data + f"\r\n--{boundary}--\r\n".encode()

    request = urllib.request.Request(
        API + "sendDocument",
        data=body,
        headers={
            "Content-Type": f"multipart/form-data; boundary={boundary}"
        }
    )

    with urllib.request.urlopen(request, timeout=60) as response:
        result = json.loads(response.read().decode())

    if not result.get("ok"):
        raise Exception(result.get("description", "فشل إرسال الملف"))


def download_file(file_id):
    info = api("getFile", {"file_id": file_id})

    size = info.get("file_size", 0)

    if size > MAX_DOWNLOAD:
        raise Exception("الحد الأقصى للملف 10 ميجابايت.")

    url = FILE_API + info["file_path"]

    with urllib.request.urlopen(url, timeout=60) as response:
        content = response.read(MAX_DOWNLOAD + 1)

    if len(content) > MAX_DOWNLOAD:
        raise Exception("الملف أكبر من الحد المسموح.")

    path = os.path.join(TEMP_DIR, uuid.uuid4().hex)

    with open(path, "wb") as file:
        file.write(content)

    return path


# =====================================
# KEYBOARD
# =====================================

def main_menu():
    return {
        "keyboard": [
            [{"text": "📝 نص إلى TXT"}, {"text": "🌐 نص إلى HTML"}],
            [{"text": "📕 نص إلى PDF"}, {"text": "🔢 عداد الكلمات"}],
            [{"text": "📊 CSV إلى Excel"}, {"text": "📄 Excel إلى CSV"}],
            [{"text": "📖 Excel إلى TXT"}, {"text": "👁 معاينة CSV"}],
            [{"text": "👤 حسابي"}, {"text": "📜 سجل العمليات"}],
            [{"text": "🛠 لوحة الإدارة"}, {"text": "❌ إلغاء"}]
        ],
        "resize_keyboard": True
    }


# =====================================
# FILE UTILITIES
# =====================================

def cleanup(*paths):
    for path in paths:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except OSError:
                pass


def text_to_txt(chat, uid, text):
    path = os.path.join(TEMP_DIR, f"{uid}.txt")

    with open(path, "w", encoding="utf-8") as file:
        file.write(text)

    send_document(chat, path, "تم إنشاء ملف TXT")
    record_operation(uid, "نص إلى TXT")
    cleanup(path)


def text_to_html(chat, uid, text):
    path = os.path.join(TEMP_DIR, f"{uid}.html")

    paragraphs = []

    for line in html.escape(text).splitlines():
        if line.strip():
            paragraphs.append(f"<p>{line}</p>")

    content = "\n".join(paragraphs)

    document = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mina Office</title>
<style>
body {{
    font-family: Arial, sans-serif;
    background: #f1f3f5;
    padding: 20px;
    line-height: 2;
}}
.paper {{
    max-width: 850px;
    margin: auto;
    background: white;
    padding: 35px;
    border-radius: 12px;
}}
h1 {{
    color: #174ea6;
    border-bottom: 2px solid #ddd;
}}
p {{
    font-size: 18px;
    white-space: pre-wrap;
}}
footer {{
    margin-top: 40px;
    color: #888;
    text-align: center;
}}
@media print {{
    body {{ background: white; }}
    .paper {{ box-shadow: none; }}
}}
</style>
</head>
<body>
<div class="paper">
<h1>مستند جديد</h1>
{content}
<footer>Created by Mina Office Bot</footer>
</div>
</body>
</html>"""

    with open(path, "w", encoding="utf-8") as file:
        file.write(document)

    send_document(chat, path, "تم إنشاء مستند HTML")
    record_operation(uid, "نص إلى HTML")
    cleanup(path)


# =====================================
# SIMPLE PDF
# =====================================

def create_simple_pdf(text, output_path):
    # ملاحظة: هذا الإصدار لا يدعم عرض العربية داخل PDF.

    lines = []

    for paragraph in text.splitlines():
        wrapped = textwrap.wrap(paragraph, width=85)
        lines.extend(wrapped if wrapped else [""])

    if not lines:
        lines = ["Empty document"]

    pages = [
        lines[i:i + 45]
        for i in range(0, len(lines), 45)
    ]

    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        None,
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
    ]

    page_ids = []

    for page_lines in pages:
        page_id = len(objects) + 1
        content_id = page_id + 1

        page_ids.append(page_id)

        commands = ["BT", "/F1 11 Tf", "50 790 Td"]

        for line in page_lines:
            safe = line.encode("cp1252", errors="replace")

            escaped = (
                safe.decode("latin-1")
                .replace("\\", "\\\\")
                .replace("(", "\\(")
                .replace(")", "\\)")
            )

            commands.append(f"({escaped}) Tj")
            commands.append("0 -16 Td")

        commands.append("ET")

        stream = "\n".join(commands).encode("latin-1")

        page_object = (
            f"<< /Type /Page /Parent 2 0 R "
            f"/MediaBox [0 0 595 842] "
            f"/Resources << /Font << /F1 3 0 R >> >> "
            f"/Contents {content_id} 0 R >>"
        ).encode()

        content_object = (
            f"<< /Length {len(stream)} >>\nstream\n".encode()
            + stream
            + b"\nendstream"
        )

        objects.append(page_object)
        objects.append(content_object)

    kids = " ".join(f"{pid} 0 R" for pid in page_ids)

    objects[1] = (
        f"<< /Type /Pages /Kids [{kids}] "
        f"/Count {len(page_ids)} >>"
    ).encode()

    pdf = bytearray(b"%PDF-1.4\n")
    offsets = [0]

    for number, obj in enumerate(objects, 1):
        offsets.append(len(pdf))
        pdf.extend(f"{number} 0 obj\n".encode())
        pdf.extend(obj)
        pdf.extend(b"\nendobj\n")

    xref = len(pdf)

    pdf.extend(
        f"xref\n0 {len(objects) + 1}\n".encode()
    )
    pdf.extend(b"0000000000 65535 f \n")

    for offset in offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode())

    pdf.extend(
        f"trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\n"
        f"startxref\n{xref}\n%%EOF\n".encode()
    )

    with open(output_path, "wb") as file:
        file.write(pdf)


def text_to_pdf(chat, uid, text):
    path = os.path.join(TEMP_DIR, f"{uid}.pdf")

    create_simple_pdf(text, path)

    send_document(chat, path, "تم إنشاء PDF")
    record_operation(uid, "نص إلى PDF")
    cleanup(path)


# =====================================
# CSV AND EXCEL
# =====================================

def csv_to_excel(chat, uid, source):
    target = os.path.join(TEMP_DIR, f"{uid}.xlsx")

    book = Workbook()
    sheet = book.active
    sheet.title = "Data"

    with open(source, "r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)
        file.seek(0)

        try:
            dialect = csv.Sniffer().sniff(sample)
        except csv.Error:
            dialect = csv.excel

        reader = csv.reader(file, dialect)

        for row in reader:
            sheet.append(row)

    book.save(target)

    send_document(chat, target, "تم تحويل CSV إلى Excel")
    record_operation(uid, "CSV إلى Excel")
    cleanup(source, target)


def excel_to_csv(chat, uid, source):
    target = os.path.join(TEMP_DIR, f"{uid}.csv")

    book = load_workbook(
        source,
        read_only=True,
        data_only=True
    )

    sheet = book.active

    with open(target, "w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)

        for row in sheet.iter_rows(values_only=True):
            writer.writerow([
                "" if value is None else value
                for value in row
            ])

    book.close()

    send_document(chat, target, "تم تحويل أول ورقة إلى CSV")
    record_operation(uid, "Excel إلى CSV")
    cleanup(source, target)


def excel_to_txt(chat, uid, source):
    target = os.path.join(TEMP_DIR, f"{uid}.txt")

    book = load_workbook(
        source,
        read_only=True,
        data_only=True
    )

    with open(target, "w", encoding="utf-8") as file:
        for sheet in book.worksheets:
            file.write(f"\n--- {sheet.title} ---\n")

            for index, row in enumerate(sheet.iter_rows(values_only=True)):
                if index >= 5000:
                    file.write("\nتم بلوغ حد 5000 صف.\n")
                    break

                file.write(
                    " | ".join(
                        str(value) if value is not None else ""
                        for value in row
                    ) + "\n"
                )

    book.close()

    send_document(chat, target, "تم استخراج محتوى Excel")
    record_operation(uid, "Excel إلى TXT")
    cleanup(source, target)


def csv_preview(chat, uid, source):
    lines = []

    with open(source, "r", encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file)

        for index, row in enumerate(reader):
            if index >= 15:
                lines.append("... انتهت المعاينة عند 15 صفًا")
                break

            lines.append(" | ".join(row))

    send(chat, "👁 معاينة CSV:\n\n" + "\n".join(lines))

    record_operation(uid, "معاينة CSV")
    cleanup(source)


# =====================================
# USER ACCOUNT AND HISTORY
# =====================================

def account(chat, uid):
    row = db.execute("""
    SELECT points, requests, joined
    FROM users
    WHERE user_id=?
    """, (uid,)).fetchone()

    send(
        chat,
        "👤 حسابك في Mina Office\n\n"
        f"🆔 رقمك: {uid}\n"
        f"⭐ النقاط: {row[0]}\n"
        f"⚙️ العمليات: {row[1]}\n"
        f"📅 تاريخ الانضمام: {row[2]}",
        main_menu()
    )


def history(chat, uid):
    rows = db.execute("""
    SELECT operation, created_at
    FROM history
    WHERE user_id=?
    ORDER BY id DESC
    LIMIT 10
    """, (uid,)).fetchall()

    if not rows:
        send(chat, "ليس لديك عمليات سابقة.")
        return

    result = "📜 آخر 10 عمليات:\n\n"

    for index, row in enumerate(rows, 1):
        result += f"{index}. {row[0]}\n"
        result += f"   {row[1]}\n\n"

    send(chat, result)


# =====================================
# ADMIN PANEL
# =====================================

def admin_panel(chat, uid):
    if uid != ADMIN_ID or ADMIN_ID == 0:
        send(chat, "⛔ ليس لديك صلاحية الدخول.")
        return

    users = db.execute(
        "SELECT COUNT(*) FROM users"
    ).fetchone()[0]

    requests = db.execute(
        "SELECT COALESCE(SUM(requests), 0) FROM users"
    ).fetchone()[0]

    points = db.execute(
        "SELECT COALESCE(SUM(points), 0) FROM users"
    ).fetchone()[0]

    files_count = 0
    storage = 0

    for root, dirs, files in os.walk(TEMP_DIR):
        for filename in files:
            path = os.path.join(root, filename)

            try:
                storage += os.path.getsize(path)
                files_count += 1
            except OSError:
                pass

    send(
        chat,
        "🛠 لوحة تحكم Mina Office\n\n"
        f"👥 المستخدمون: {users}\n"
        f"⚙️ العمليات: {requests}\n"
        f"⭐ النقاط: {points}\n"
        f"📁 الملفات المؤقتة: {files_count}\n"
        f"💾 المساحة: {storage / 1024 / 1024:.2f} MB\n\n"
        "🟢 البوت يعمل",
        main_menu()
    )


# =====================================
# DOCUMENT HANDLER
# =====================================

def handle_document(message):
    chat = message["chat"]["id"]
    uid = message["from"]["id"]

    mode = user_modes.get(uid)

    if not mode:
        send(chat, "اختار خدمة من القائمة الأول.", main_menu())
        return

    filename = message["document"].get("file_name", "")
    extension = os.path.splitext(filename)[1].lower()

    allowed = {
        "csv_excel": [".csv"],
        "excel_csv": [".xlsx"],
        "excel_txt": [".xlsx"],
        "csv_preview": [".csv"]
    }

    if extension not in allowed.get(mode, []):
        send(chat, "صيغة الملف غير مناسبة. راجع الصيغ المطلوبة.")
        return

    user_modes.pop(uid, None)
    source = None

    try:
        send(chat, "⏳ جاري معالجة الملف...")

        source = download_file(message["document"]["file_id"])

        if mode == "csv_excel":
            csv_to_excel(chat, uid, source)
        elif mode == "excel_csv":
            excel_to_csv(chat, uid, source)
        elif mode == "excel_txt":
            excel_to_txt(chat, uid, source)
        elif mode == "csv_preview":
            csv_preview(chat, uid, source)

    except Exception as error:
        cleanup(source)
        send(chat, f"حدث خطأ أثناء معالجة الملف:\n{error}")


# =====================================
# MESSAGE HANDLER
# =====================================

def handle(message):
    chat = message["chat"]["id"]
    uid = message["from"]["id"]
    text = message.get("text", "")

    register_user(uid)

    if "document" in message:
        handle_document(message)
        return

    if not text:
        return

    if text == "/start":
        user_modes.pop(uid, None)

        send(
            chat,
            "مرحبًا بك في Mina Office Bot Ultimate 📁\n\n"
            "مساعدك المكتبي لتحويل الملفات وإدارة النصوص.\n"
            "اختر الخدمة من القائمة.",
            main_menu()
        )

    elif text == "/myid":
        send(chat, f"رقم حسابك:\n{uid}")

    elif text == "/cancel" or text == "❌ إلغاء":
        user_modes.pop(uid, None)
        send(chat, "تم إلغاء العملية.", main_menu())

    elif text == "/admin" or text == "🛠 لوحة الإدارة":
        admin_panel(chat, uid)

    elif text == "👤 حسابي":
        account(chat, uid)

    elif text == "📜 سجل العمليات":
        history(chat, uid)

    elif text == "📝 نص إلى TXT":
        user_modes[uid] = "txt"
        send(chat, "أرسل النص الذي تريد تحويله إلى TXT.")

    elif text == "🌐 نص إلى HTML":
        user_modes[uid] = "html"
        send(chat, "أرسل النص الذي تريد تحويله إلى HTML.")

    elif text == "📕 نص إلى PDF":
        user_modes[uid] = "pdf"
        send(
            chat,
            "أرسل نصًا باللغة الإنجليزية لإنشاء PDF.\n"
            "ملاحظة: إصدار PDF الحالي لا يعرض العربية بشكل صحيح."
        )

    elif text == "🔢 عداد الكلمات":
        user_modes[uid] = "count"
        send(chat, "أرسل النص الذي تريد حساب كلماته.")

    elif text == "📊 CSV إلى Excel":
        user_modes[uid] = "csv_excel"
        send(chat, "أرسل ملف CSV.")

    elif text == "📄 Excel إلى CSV":
        user_modes[uid] = "excel_csv"
        send(chat, "أرسل ملف XLSX.")

    elif text == "📖 Excel إلى TXT":
        user_modes[uid] = "excel_txt"
        send(chat, "أرسل ملف XLSX.")

    elif text == "👁 معاينة CSV":
        user_modes[uid] = "csv_preview"
        send(chat, "أرسل ملف CSV لمعاينته.")

    elif user_modes.get(uid) == "txt":
        try:
            text_to_txt(chat, uid, text)
            user_modes.pop(uid, None)
        except Exception as error:
            send(chat, f"حدث خطأ: {error}")

    elif user_modes.get(uid) == "html":
        try:
            text_to_html(chat, uid, text)
            user_modes.pop(uid, None)
        except Exception as error:
            send(chat, f"حدث خطأ: {error}")

    elif user_modes.get(uid) == "pdf":
        try:
            text_to_pdf(chat, uid, text)
            user_modes.pop(uid, None)
        except Exception as error:
            send(chat, f"حدث خطأ: {error}")

    elif user_modes.get(uid) == "count":
        words = len(text.split())
        chars = len(text)

        record_operation(uid, "عداد الكلمات")
        user_modes.pop(uid, None)

        send(
            chat,
            "📊 تحليل النص\n\n"
            f"الكلمات: {words}\n"
            f"الحروف: {chars}\n"
            f"المسافات: {text.count(' ')}",
            main_menu()
        )

    else:
        send(chat, "استخدم القائمة لاختيار خدمة.", main_menu())


# =====================================
# MAIN
# =====================================

def main():
    print("=" * 35)
    print("MINA OFFICE BOT ULTIMATE V1")
    print("Starting...")
    print("=" * 35)

    # التأكد من عدم وجود Webhook يعطل polling
    webhook = api("getWebhookInfo")

    if webhook.get("url"):
        print("يوجد Webhook نشط.")
        print("أوقفه قبل تشغيل polling.")
        return

    offset = None

    while True:
        try:
            data = {"timeout": 25}

            if offset is not None:
                data["offset"] = offset

            updates = api("getUpdates", data)

            for update in updates:
                offset = update["update_id"] + 1

                if "message" in update:
                    try:
                        handle(update["message"])
                    except Exception as error:
                        print("Message error:", error)

        except KeyboardInterrupt:
            print("Bot stopped.")
            break

        except Exception as error:
            print("Connection error:", error)
            time.sleep(5)


if __name__ == "__main__":
    if TOKEN == "PUT_YOUR_BOT_TOKEN_HERE":
        print("ضع توكن البوت في TOKEN أولاً.")
    else:
        main()
