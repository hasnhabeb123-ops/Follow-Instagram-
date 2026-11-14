import logging
import sqlite3
import os
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# توكن البوت ومعرف الأدمن الخاص بك
TOKEN = "8789485311:AAHlP1GDBu5Kr0sJe0y8JkHFDEot9wOQXlg"
ADMIN_ID = 5859090443

# إعدادات التسجيل لمتابعة الأخطاء
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# --- 1. إعداد قاعدة البيانات والذاكرة ---
def init_db():
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    # جدول المستخدمين ورصيدهم
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            balance INTEGER DEFAULT 5
        )
    ''')
    # جدول الذاكرة لتتبع آخر طلب لكل مستخدم
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS user_memory (
            user_id INTEGER PRIMARY KEY,
            last_request TEXT
        )
    ''')
    conn.commit()
    conn.close()

def get_or_create_user(user_id, username):
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    cursor.execute('SELECT balance FROM users WHERE user_id = ?', (user_id,))
    row = cursor.fetchone()
    
    if row is None:
        cursor.execute('INSERT INTO users (user_id, username, balance) VALUES (?, ?, ?)', (user_id, username, 5))
        conn.commit()
        balance = 5
    else:
        balance = row[0]
        
    conn.close()
    return balance

def update_balance(user_id, amount):
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    cursor.execute('UPDATE users SET balance = balance + ? WHERE user_id = ?', (amount, user_id))
    conn.commit()
    conn.close()

def save_user_memory(user_id, text):
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    cursor.execute('INSERT OR REPLACE INTO user_memory (user_id, last_request) VALUES (?, ?)', (user_id, text))
    conn.commit()
    conn.close()

def get_user_memory(user_id):
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    cursor.execute('SELECT last_request FROM user_memory WHERE user_id = ?', (user_id,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else None

# دالة لضبط اتجاه الفقرة إلى اليمين (RTL)
def set_rtl(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = 2  # محاذاة لليمين

# --- 2. أوامر البوت ---
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    balance = get_or_create_user(user.id, user.username)
    
    welcome_text = (
        f"أهلاً بك يا {user.first_name} في بوت توليد المستندات الذكي! 📄✨\n\n"
        f"💰 رصيدك الحالي: **{balance} عملة**\n\n"
        f"الأوامر المتاحة:\n"
        f"/balance - لعرض رصيدك\n"
        f"/word [النص] - لتوليد ملف Word منسق بالعربي\n"
        f"/edit [التعديل] - لتعديل آخر ملف طلبته\n"
        f"/help - لمعرفة طريقة الاستخدام"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")

async def balance_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    balance = get_or_create_user(user.id, user.username)
    await update.message.reply_text(f"💰 رصيدك الحالي هو: **{balance} عملة**", parse_mode="Markdown")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = (
        "📖 دليل الاستخدام الاحترافي:\n\n"
        "1️⃣ لتوليد ملف جديد:\n"
        "`/word تقرير مبيعات الشركة لعام 2026`\n\n"
        "2️⃣ للتعديل على آخر ملف:\n"
        "`/edit أضف قسم خاص بالخاتمة والتوصيات`\n\n"
        "ملاحظة: كل عملية توليد مستند جديدة تخصم عملة واحدة من رصيدك."
    )
    await update.message.reply_text(help_text, parse_mode="Markdown")

# --- 3. توليد ملفات الـ Word مع الذاكرة ---
async def create_word_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    balance = get_or_create_user(user.id, user.username)
    
    if balance < 1:
        await update.message.reply_text("❌ رصيدك غير كافي! تحتاج إلى عملة واحدة على الأقل لتوليد الملف.")
        return

    if not context.args:
        await update.message.reply_text("⚠️ يرجى كتابة محتوى أو عنوان للملف بعد الأمر.\nمثال: `/word تقرير مبيعات`", parse_mode="Markdown")
        return

    query_text = " ".join(context.args)
    
    # حفظ الطلب في ذاكرة المستخدم
    save_user_memory(user.id, query_text)
    
    # خصم عملة من رصيد المستخدم
    update_balance(user.id, -1)
    
    await update.message.reply_text("⏳ جاري تحليل الطلب وتصميم ملف الـ Word المنسق...")

    doc = Document()
    
    p1 = doc.add_heading('تقرير مستند صادر من البوت الذكي', 0)
    set_rtl(p1)
    
    p2 = doc.add_heading('موضوع الطلب الأساسي:', level=1)
    set_rtl(p2)
    
    p3 = doc.add_paragraph(query_text)
    set_rtl(p3)
    
    p4 = doc.add_heading('التفاصيل والتنسيق:', level=1)
    set_rtl(p4)
    
    p5 = doc.add_paragraph("تم إعداد هذا المستند بناءً على طلبك مع حفظ الذاكرة لتتمكن من التعديل عليه لاحقاً باستخدام الأمر /edit.")
    set_rtl(p5)
    
    file_path = f"document_{user.id}.docx"
    doc.save(file_path)

    with open(file_path, 'rb') as f:
        await update.message.reply_document(
            document=f,
            filename="Generated_Document.docx",
            caption="✅ تفضل ملف الـ Word الخاص بك جاهز وباتجاه صحيح!"
        )

    if os.path.exists(file_path):
        os.remove(file_path)

# --- 4. نظام التعديل التفاعلي (استرجاع الذاكرة) ---
async def edit_word_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    
    if not context.args:
        await update.message.reply_text("⚠️ يرجى كتابة التعديل المطلوب.\nمثال: `/edit أضف بند الميزانية`", parse_mode="Markdown")
        return

    edit_text = " ".join(context.args)
    last_request = get_user_memory(user.id)
    
    if not last_request:
        await update.message.reply_text("❌ لا يوجد طلب سابق مسجل في الذاكرة! يرجى إنشاء ملف جديد أولاً باستخدام `/word`.", parse_mode="Markdown")
        return

    await update.message.reply_text(f"🔄 جاري تعديل المستند بناءً على طلبك الأخير...\n(الطلب السابق: {last_request})")

    # دمج الطلب القديم مع التعديل الجديد لإنشاء ملف محدث
    doc = Document()
    
    p1 = doc.add_heading('مستند محدث من البوت الذكي', 0)
    set_rtl(p1)
    
    p2 = doc.add_heading('الموضوع الأصلي:', level=1)
    set_rtl(p2)
    
    p3 = doc.add_paragraph(last_request)
    set_rtl(p3)
    
    p4 = doc.add_heading('التعديلات المضافة حديثاً:', level=1)
    set_rtl(p4)
    
    p5 = doc.add_paragraph(edit_text)
    set_rtl(p5)
    
    file_path = f"updated_document_{user.id}.docx"
    doc.save(file_path)

    with open(file_path, 'rb') as f:
        await update.message.reply_document(
            document=f,
            filename="Updated_Document.docx",
            caption="✅ تفضل الملف المحدث بناءً على تعديلاتك الأخيرة!"
        )

    if os.path.exists(file_path):
        os.remove(file_path)

# --- 5. تشغيل البوت ---
def main():
    init_db()
    
    application = Application.builder().token(TOKEN).build()

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("balance", balance_command))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("word", create_word_command))
    application.add_handler(CommandHandler("edit", edit_word_command))

    print("Bot is starting with memory and editing support...")
    application.run_polling()

if __name__ == "__main__":
    main()
