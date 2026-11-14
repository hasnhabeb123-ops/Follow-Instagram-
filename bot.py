
import os
import re
import json
import uuid
import sqlite3
import logging
from pathlib import Path
from datetime import datetime

from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from telegram import Update
from telegram.constants import ChatAction
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)


# ==========================================================
# 1. CONFIGURATION
# ==========================================================

BOT_TOKEN = os.getenv("8789485311:AAFHPlGDBu5Kr0sJe0y8JkHFDE0t9wOQXLg")

DATABASE = "bot_database.db"
OUTPUT_DIR = Path("generated_files")

INITIAL_BALANCE = 10
DOCUMENT_COST = 2

# ضع أرقام حسابات المشرفين هنا
ADMIN_IDS = []

# عدد الصفوف الفارغة تحت أسماء الأعمدة
EMPTY_ROWS = 5

MAX_COLUMNS = 30
MAX_FIELD_LENGTH = 80

OUTPUT_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)

logger = logging.getLogger(__name__)


# ==========================================================
# 2. DATABASE
# ==========================================================

def get_connection():
    conn = sqlite3.connect(DATABASE, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def init_database():
    with get_connection() as conn:

        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                first_name TEXT,
                last_name TEXT,
                username TEXT,
                balance INTEGER NOT NULL DEFAULT 10,
                joined_at TEXT,
                welcomed INTEGER NOT NULL DEFAULT 0
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS user_context (
                user_id INTEGER PRIMARY KEY,
                last_fields TEXT,
                last_request TEXT,
                last_file TEXT,
                updated_at TEXT,
                FOREIGN KEY(user_id) REFERENCES users(user_id)
            )
        """)

        conn.commit()


def register_user(user):
    with get_connection() as conn:

        existing = conn.execute(
            "SELECT * FROM users WHERE user_id = ?",
            (user.id,)
        ).fetchone()

        if existing is None:

            conn.execute("""
                INSERT INTO users (
                    user_id,
                    first_name,
                    last_name,
                    username,
                    balance,
                    joined_at,
                    welcomed
                )
                VALUES (?, ?, ?, ?, ?, ?, 0)
            """, (
                user.id,
                user.first_name or "",
                user.last_name or "",
                user.username or "",
                INITIAL_BALANCE,
                datetime.now().isoformat()
            ))

        else:

            conn.execute("""
                UPDATE users
                SET first_name = ?,
                    last_name = ?,
                    username = ?
                WHERE user_id = ?
            """, (
                user.first_name or "",
                user.last_name or "",
                user.username or "",
                user.id
            ))

        conn.commit()


def get_user(user_id):

    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM users WHERE user_id = ?",
            (user_id,)
        ).fetchone()


def get_balance(user_id):

    user = get_user(user_id)

    if user is None:
        return 0

    return user["balance"]


def set_welcomed(user_id):

    with get_connection() as conn:
        conn.execute(
            "UPDATE users SET welcomed = 1 WHERE user_id = ?",
            (user_id,)
        )
        conn.commit()


def save_context(user_id, fields, request_text, file_path):

    with get_connection() as conn:

        conn.execute("""
            INSERT INTO user_context (
                user_id,
                last_fields,
                last_request,
                last_file,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?)

            ON CONFLICT(user_id)
            DO UPDATE SET
                last_fields = excluded.last_fields,
                last_request = excluded.last_request,
                last_file = excluded.last_file,
                updated_at = excluded.updated_at
        """, (
            user_id,
            json.dumps(fields, ensure_ascii=False),
            request_text,
            str(file_path),
            datetime.now().isoformat()
        ))

        conn.commit()


def get_context(user_id):

    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM user_context WHERE user_id = ?",
            (user_id,)
        ).fetchone()

    if row is None:
        return None

    return {
        "fields": json.loads(row["last_fields"] or "[]"),
        "request": row["last_request"],
        "file": row["last_file"],
    }


def charge_user(user_id, amount):

    with get_connection() as conn:

        cursor = conn.execute("""
            UPDATE users
            SET balance = balance - ?
            WHERE user_id = ?
            AND balance >= ?
        """, (amount, user_id, amount))

        conn.commit()

        return cursor.rowcount == 1


def refund_user(user_id, amount):

    with get_connection() as conn:

        conn.execute("""
            UPDATE users
            SET balance = balance + ?
            WHERE user_id = ?
        """, (amount, user_id))

        conn.commit()


def add_coins(user_id, amount):

    if amount <= 0:
        return False

    with get_connection() as conn:

        cursor = conn.execute("""
            UPDATE users
            SET balance = balance + ?
            WHERE user_id = ?
        """, (amount, user_id))

        conn.commit()

        return cursor.rowcount == 1


# ==========================================================
# 3. USER NAME
# ==========================================================

def get_display_name(user):

    first = (user.first_name or "").strip()
    last = (user.last_name or "").strip()

    full_name = f"{first} {last}".strip()

    if full_name:
        return full_name

    if user.username:
        return f"@{user.username}"

    return "صديقي"


# ==========================================================
# 4. TEXT NORMALIZATION
# ==========================================================

def normalize_text(text):

    text = text.strip()

    replacements = {
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ى": "ي",
        "ـ": "",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    return text


def clean_field(field):

    field = field.strip()

    field = re.sub(
        r"^(و|ثم)\s*",
        "",
        field
    )

    field = re.sub(
        r"^[\s:：\-]+|[\s:：\-]+$",
        "",
        field
    )

    field = re.sub(
        r"\s+",
        " ",
        field
    )

    return field[:MAX_FIELD_LENGTH]


# ==========================================================
# 5. DYNAMIC FIELD EXTRACTION
# ==========================================================

def extract_fields(text):

    normalized = normalize_text(text)

    # يجب أن يحتوي الطلب على إشارة واضحة لإنشاء جدول
    table_pattern = (
        r"(جدول|جداول|اعمل\s*لي\s*جدول|"
        r"اعملي\s*جدول|انشئ\s*جدول|"
        r"إنشاء\s*جدول|صمم\s*جدول)"
    )

    if not re.search(table_pattern, normalized):
        return []

    # استخراج الجزء الذي يحتوي على الحقول
    patterns = [
        r"(?:فيه|في|بأعمدة|بعناوين|الاعمدة|الأعمدة)\s+(.+)",
        r"(?:يتضمن|يحتوي على)\s+(.+)",
    ]

    fields_text = None

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:
            fields_text = match.group(1)
            break

    if not fields_text:
        return []

    # حذف عبارات نهاية الطلب
    fields_text = re.split(
        r"\s+(?:من فضلك|لو سمحت|رجاء|بصيغة|بملف)\b",
        fields_text,
        maxsplit=1
    )[0]

    # تقسيم الحقول بحسب علامات الفصل
    parts = re.split(
        r"[,،;؛\n]+|\s+و(?=\s*(?:ال|اسم|رقم|تاريخ|عنوان|مكان|نوع|حالة|ملاحظات|هاتف|جنسية|وظيفة))",
        fields_text
    )

    fields = []

    for part in parts:

        field = clean_field(part)

        if not field:
            continue

        # منع تكرار الحقول
        if field not in fields:
            fields.append(field)

    if not fields or len(fields) > MAX_COLUMNS:
        return []

    return fields


# ==========================================================
# 6. EDITING INTENT DETECTION
# ==========================================================

def parse_edit(text, current_fields):

    normalized = normalize_text(text)

    # تغيير اسم عمود
    rename_match = re.search(
        r"(?:غير|بدل|عدّل|عدل)\s+(?:اسم\s+)?(?:خانة|عمود)\s+(.+?)"
        r"\s+(?:الى|إلى|باسم|ليصبح)\s+(.+)",
        normalized
    )

    if rename_match:

        old = clean_field(rename_match.group(1))
        new = clean_field(rename_match.group(2))

        for index, field in enumerate(current_fields):

            if normalize_text(field) == old:

                updated = current_fields.copy()
                updated[index] = new

                if new in updated[:index] or new in updated[index + 1:]:
                    return None, "هذا العمود موجود مسبقاً."

                return updated, None

        return None, f"لم أجد العمود: {old}"

    # حذف عمود
    delete_match = re.search(
        r"(?:احذف|حذف|شيل|إزالة|ازالة|الغاء|إلغاء)"
        r"\s+(?:خانة|عمود)\s+(.+)",
        normalized
    )

    if delete_match:

        target = clean_field(delete_match.group(1))

        updated = [
            field for field in current_fields
            if normalize_text(field) != target
        ]

        if len(updated) == len(current_fields):
            return None, f"لم أجد العمود: {target}"

        if not updated:
            return None, "لا يمكن حذف جميع الأعمدة."

        return updated, None

    # إضافة عمود
    add_match = re.search(
        r"(?:اضف|أضف|ضيف|إضافة|زيادة|ناقص)"
        r"\s+(?:خانة|عمود)?\s*(.+)",
        normalized
    )

    if add_match:

        new_field = clean_field(add_match.group(1))

        if not new_field:
            return None, "حدد اسم العمود الذي تريد إضافته."

        if any(
            normalize_text(f) == new_field
            for f in current_fields
        ):
            return None, "هذا العمود موجود مسبقاً."

        if len(current_fields) >= MAX_COLUMNS:
            return None, "وصل الجدول للحد الأقصى من الأعمدة."

        return current_fields + [new_field], None

    return None, None


# ==========================================================
# 7. WORD RTL HELPERS
# ==========================================================

def set_element(element, tag, value=True):

    node = element.find(qn(tag))

    if node is None:
        node = OxmlElement(tag)
        element.append(node)

    node.set(qn("w:val"), "1" if value else "0")

    return node


def set_paragraph_rtl(paragraph):

    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p_pr = paragraph._p.get_or_add_pPr()

    set_element(p_pr, "w:bidi", True)

    for run in paragraph.runs:

        r_pr = run._r.get_or_add_rPr()

        set_element(r_pr, "w:rtl", True)

        run.font.name = "Arial"
        run.font.size = Pt(11)


def set_cell_rtl(cell):

    tc_pr = cell._tc.get_or_add_tcPr()

    # اتجاه محتوى الخلية
    set_element(tc_pr, "w:textDirection", False)

    for paragraph in cell.paragraphs:
        set_paragraph_rtl(paragraph)


def set_cell_shading(cell, fill="D9EAF7"):

    tc_pr = cell._tc.get_or_add_tcPr()

    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)

    tc_pr.append(shading)


def set_table_rtl(table):

    tbl_pr = table._tbl.tblPr

    # ترتيب الأعمدة من اليمين لليسار
    set_element(tbl_pr, "w:bidiVisual", True)

    table.alignment = WD_TABLE_ALIGNMENT.CENTER


# ==========================================================
# 8. WORD DOCUMENT GENERATOR
# ==========================================================

def create_word_document(fields, user_id):

    if not fields:
        raise ValueError("لا يمكن إنشاء جدول بدون أعمدة.")

    document = Document()

    section = document.sections[0]

    section.right_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    # اتجاه الصفحة
    sect_pr = section._sectPr

    set_element(sect_pr, "w:bidi", True)

    # إنشاء الجدول: صف عناوين + صفوف فارغة
    table = document.add_table(
        rows=1 + EMPTY_ROWS,
        cols=len(fields)
    )

    table.style = "Table Grid"

    set_table_rtl(table)

    # العناوين
    for index, field in enumerate(fields):

        cell = table.rows[0].cells[index]

        cell.text = field

        set_cell_shading(cell)

        for paragraph in cell.paragraphs:

            set_paragraph_rtl(paragraph)

            for run in paragraph.runs:

                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)

        set_cell_rtl(cell)

    # تنسيق الصفوف الفارغة
    for row in table.rows[1:]:

        for cell in row.cells:

            cell.text = ""

            set_cell_rtl(cell)

    # تنسيق الخط الافتراضي
    for paragraph in document.paragraphs:
        set_paragraph_rtl(paragraph)

    filename = (
        f"user_{user_id}_"
        f"{uuid.uuid4().hex[:10]}.docx"
    )

    file_path = OUTPUT_DIR / filename

    document.save(str(file_path))

    return file_path


# ==========================================================
# 9. SEND DOCUMENT
# ==========================================================

async def send_document(
    update,
    context,
    fields,
    request_text,
    is_new=True
):

    user = update.effective_user
    user_id = user.id

    current_balance = get_balance(user_id)

    if current_balance < DOCUMENT_COST:

        await update.message.reply_text(
            f"رصيدك الحالي {current_balance} عملة.\n"
            f"تحتاج {DOCUMENT_COST} عملة لإنشاء الملف."
        )

        return

    file_path = None

    try:

        # إنشاء الملف أولاً
        file_path = create_word_document(
            fields,
            user_id
        )

        # الخصم بعد نجاح إنشاء الملف
        if not charge_user(user_id, DOCUMENT_COST):

            await update.message.reply_text(
                "رصيدك غير كافٍ لإتمام العملية."
            )

            return

        try:

            await update.message.reply_document(
                document=open(file_path, "rb"),
                filename="جدول_المستخدم.docx",
                caption=(
                    "تم إنشاء الجدول بنجاح."
                    if is_new
                    else "تم تحديث الجدول وإرسال النسخة الجديدة."
                )
            )

        except Exception:

            # استرداد الرصيد عند فشل الإرسال
            refund_user(user_id, DOCUMENT_COST)

            raise

        # حفظ سياق المستخدم بعد نجاح الإرسال
        save_context(
            user_id,
            fields,
            request_text,
            file_path
        )

        await update.message.reply_text(
            f"رصيدك المتبقي: {get_balance(user_id)} عملة."
        )

    except Exception as error:

        logger.exception(
            "Document generation/send error: %s",
            error
        )

        await update.message.reply_text(
            "حدث خطأ أثناء تجهيز الملف أو إرساله. "
            "حاول مرة ثانية."
        )

    finally:

        if file_path and file_path.exists():

            try:
                file_path.unlink()
            except OSError:
                logger.warning("Could not delete temporary file.")


# ==========================================================
# 10. START COMMAND
# ==========================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user = update.effective_user

    register_user(user)

    db_user = get_user(user.id)

    if not db_user["welcomed"]:

        name = get_display_name(user)

        await update.message.reply_text(
            f"أهلاً وسهلاً {name}!\n\n"
            "أنا بوت إنشاء جداول Word.\n"
            "أرسل طلبك بشكل طبيعي، مثلاً:\n\n"
            "عملي جدول فيه الاسم الثلاثي، "
            "اسم الأم، رقم الهاتف، الرقم الوطني.\n\n"
            "ويمكنك بعد إنشاء الجدول طلب إضافة "
            "أو حذف أو تغيير أسماء الأعمدة.\n\n"
            "استخدم /balance لمعرفة رصيدك."
        )

        set_welcomed(user.id)

    else:

        await update.message.reply_text(
            f"أهلاً بعودتك {get_display_name(user)}!\n"
            "كيف فيني ساعدك اليوم؟"
        )


# ==========================================================
# 11. BALANCE COMMAND
# ==========================================================

async def balance_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user = update.effective_user

    register_user(user)

    await update.message.reply_text(
        f"رصيدك الحالي: {get_balance(user.id)} عملة."
    )


# ==========================================================
# 12. HELP COMMAND
# ==========================================================

async def help_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    await update.message.reply_text(
        "الأوامر المتاحة:\n\n"
        "/start - بدء البوت\n"
        "/balance - معرفة الرصيد\n"
        "/help - المساعدة\n\n"
        "إنشاء جدول:\n"
        "عملي جدول فيه الاسم، الهاتف، العنوان.\n\n"
        "إضافة عمود:\n"
        "أضف خانة تاريخ الميلاد\n\n"
        "حذف عمود:\n"
        "احذف خانة العنوان\n\n"
        "تغيير اسم عمود:\n"
        "غير اسم خانة الهاتف إلى رقم الهاتف"
    )


# ==========================================================
# 13. ADMIN COMMAND
# ==========================================================

async def admin_add_coins(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    admin_id = update.effective_user.id

    if admin_id not in ADMIN_IDS:

        await update.message.reply_text(
            "هذا الأمر مخصص للمشرفين فقط."
        )

        return

    if len(context.args) != 2:

        await update.message.reply_text(
            "الاستخدام:\n/addcoins USER_ID AMOUNT"
        )

        return

    try:

        target_id = int(context.args[0])
        amount = int(context.args[1])

    except ValueError:

        await update.message.reply_text(
            "تأكد من إدخال أرقام صحيحة."
        )

        return

    if amount <= 0:

        await update.message.reply_text(
            "يجب أن يكون عدد العملات أكبر من صفر."
        )

        return

    if not add_coins(target_id, amount):

        await update.message.reply_text(
            "المستخدم غير مسجل في قاعدة البيانات."
        )

        return

    await update.message.reply_text(
        f"تمت إضافة {amount} عملة للمستخدم "
        f"{target_id}."
    )


# ==========================================================
# 14. MAIN MESSAGE HANDLER
# ==========================================================

async def handle_message(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    if not update.message or not update.message.text:
        return

    user = update.effective_user
    user_id = user.id

    register_user(user)

    text = update.message.text.strip()

    if not text:
        return

    await context.bot.send_chat_action(
        chat_id=update.effective_chat.id,
        action=ChatAction.TYPING
    )

    # أولاً: هل المستخدم يطلب إنشاء جدول جديد؟
    fields = extract_fields(text)

    if fields:

        await send_document(
            update,
            context,
            fields,
            text,
            is_new=True
        )

        return

    # ثانياً: هل يوجد جدول سابق للمستخدم؟
    saved_context = get_context(user_id)

    if saved_context is None:

        await update.message.reply_text(
            "أرسل طلب إنشاء جدول أولاً، "
            "حتى أقدر أتعامل مع التعديلات."
        )

        return

    current_fields = saved_context["fields"]

    # ثالثاً: محاولة فهم التعديل
    updated_fields, error = parse_edit(
        text,
        current_fields
    )

    if error:

        await update.message.reply_text(error)

        return

    if updated_fields is not None:

        await send_document(
            update,
            context,
            updated_fields,
            text,
            is_new=False
        )

        return

    # رسالة لا تحتوي طلب إنشاء أو تعديل واضح
    await update.message.reply_text(
        "ما قدرت أحدد المطلوب.\n\n"
        "إذا بدك تعدل الجدول، استخدم مثلاً:\n"
        "أضف خانة تاريخ الميلاد\n"
        "احذف خانة العنوان\n"
        "غير اسم خانة الهاتف إلى رقم الهاتف"
    )


# ==========================================================
# 15. ERROR HANDLER
# ==========================================================

async def error_handler(
    update: object,
    context: ContextTypes.DEFAULT_TYPE
):

    logger.error(
        "Unhandled exception:",
        exc_info=context.error
    )

    if isinstance(update, Update) and update.effective_message:

        try:

            await update.effective_message.reply_text(
                "صار خطأ غير متوقع. "
                "حاول مرة ثانية بعد قليل."
            )

        except Exception:

            logger.exception("Could not send error message.")


# ==========================================================
# 16. MAIN
# ==========================================================

def main():

    if (
        not BOT_TOKEN
        or BOT_TOKEN == "PUT_YOUR_BOT_TOKEN_HERE"
    ):

        raise ValueError(
            "ضع توكن البوت في متغير BOT_TOKEN "
            "أو داخل إعدادات التشغيل."
        )

    init_database()

    application = (
        Application.builder()
        .token(BOT_TOKEN)
        .build()
    )

    application.add_handler(
        CommandHandler("start", start)
    )

    application.add_handler(
        CommandHandler("balance", balance_command)
    )

    application.add_handler(
        CommandHandler("help", help_command)
    )

    application.add_handler(
        CommandHandler("addcoins", admin_add_coins)
    )

    application.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND,
            handle_message
        )
    )

    application.add_error_handler(error_handler)

    logger.info("Bot started successfully.")

    application.run_polling(
        allowed_updates=Update.ALL_TYPES
    )


if __name__ == "__main__":

    try:
        main()

    except KeyboardInterrupt:
        logger.info("Bot stopped.")

    except Exception:
        logger.exception("Fatal startup error.")
        raise
