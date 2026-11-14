import logging
import sqlite3
import os
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes
from docx import Document

# توكن البوت ومعرف الأدمن الخاص بك
TOKEN = "8789485311:AAFHPlGDBu5Kr0sJe0y8JkHFDE0t9wOQXLg"

ADMIN_ID = 5859090443

# إعدادات التسجيل لمتابعة الأخطاء
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# --- 1. إعداد قاعدة البيانات ---
def init_db():
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            balance INTEGER DEFAULT 5
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

# --- 2. أوامر البوت ---
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    balance = get_or_create_user(user.id, user.username)
    
    welcome_text = (
        f"أهلاً بك يا {user.first_name} في بوت توليد المستندات الذكي! 📄✨\n\n"
        f"💰 رصيدك الحالي: **{balance} عملة**\n\n"
        f"الأوامر المتاحة:\n"
        f"/balance - لعرض رصيدك\n"
        f"/word [النص] - لتوليد ملف Word منسق\n"
        f"/help - لمعرفة طريقة الاستخدام"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")

async def balance_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    balance = get_or_create_user(user.id, user.username)
    await update.message.reply_text(f"💰 رصيدك الحالي هو: **{balance} عملة**", parse_mode="Markdown")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = (
        "طريقة الاستخدام:\n"
        "اكتب الأمر /word متبوعاً بالعنوان أو المحتوى الذي تريده، وسيقوم البوت بإنشاء ملف Word احترافي وإرساله لك مقابل خصم عملة واحدة.\n\n"
        "مثال:\n"
        "/word تقرير عن هندسة الروبوتات"
    )
    await update.message.reply_text(help_text)

# --- 3. توليد ملفات الـ Word ---
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
    
    # خصم عملة من رصيد المستخدم
    update_balance(user.id, -1)
    
    await update.message.reply_text("⏳ جاري إنشاء وتنسيق ملف الـ Word، ثوانٍ معدودة...")

    # إنشاء ملف الـ Word باستخدام python-docx
    doc = Document()
    doc.add_heading('تقرير مستند صادر من البوت الذكي', 0)
    
    doc.add_heading('موضوع الطلب:', level=1)
    doc.add_paragraph(query_text)
    
    doc.add_heading('التفاصيل والمحتوى:', level=1)
    doc.add_paragraph("تم توليد هذا الملف تلقائياً بناءً على طلبك عبر بوت التيليجرام الاحترافي وبكل احترافية.")
    
    file_path = f"document_{user.id}.docx"
    doc.save(file_path)

    # إرسال الملف للمستخدم
    with open(file_path, 'rb') as f:
        await update.message.reply_document(
            document=f,
            filename="Generated_Document.docx",
            caption="✅ تفضل ملف الـ Word الخاص بك جاهز ومنسق!"
        )

    # تنظيف الملف من السيرفر
    if os.path.exists(file_path):
        os.remove(file_path)

# --- 4. تشغيل البوت ---
def main():
    init_db()
    
    application = Application.builder().token(TOKEN).build()

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("balance", balance_command))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("word", create_word_command))

    print("Bot is starting with document generation support...")
    application.run_polling()

if __name__ == "__main__":
    main()
