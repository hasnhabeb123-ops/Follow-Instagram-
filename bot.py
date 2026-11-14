import logging
import sqlite3
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes

# توكن البوت ومعرف الأدمن الخاص بك
TOKEN = "8789485311:AAHlP1GDBu5Kr0sJe0y8JkHFDEot9w"
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

# --- 2. أوامر البوت ---
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    balance = get_or_create_user(user.id, user.username)
    
    welcome_text = (
        f"أهلاً بك يا {user.first_name} في بوت توليد المستندات الذكي! 📄✨\n\n"
        f"هذا البوت يساعدك على إنشاء ملفات Word و Excel منسقة واحترافية خلال ثوانٍ.\n\n"
        f"💰 رصيدك الحالي: **{balance} عملة**\n\n"
        f"الأوامر المتاحة:\n"
        f"/balance - لعرض رصيدك\n"
        f"/help - لمعرفة طريقة الاستخدام"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")

async def balance_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    balance = get_or_create_user(user.id, user.username)
    await update.message.reply_text(f"💰 رصيدك الحالي هو: **{balance} عملة**", parse_mode="Markdown")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = (
        "طريقة الاستخدام بسيطة جداً:\n"
        "فقط اكتب طلبك بشكل واضح (مثلاً: ابحث لي عن كذا أو أنشئ جدول بيانات كذا)، وسيقوم البوت بتجهيز الملف لك مباشرة مقابل خصم عملة من رصيدك."
    )
    await update.message.reply_text(help_text)

# --- 3. تشغيل البوت ---
def main():
    init_db()
    
    # بناء تطبيق البوت
    application = Application.builder().token(TOKEN).build()

    # ربط الأوامر بالدوال
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("balance", balance_command))
    application.add_handler(CommandHandler("help", help_command))

    print("Bot is starting...")
    # بدء تشغيل البوت بنظام Polling
    application.run_polling()

if __name__ == "__main__":
    main()
    
