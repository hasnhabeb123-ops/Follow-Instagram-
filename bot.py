import logging
import sqlite3
import os
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, ContextTypes, filters
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# توكن البوت ومعرف الأدمن الخاص بك
TOKEN = "8789485311:AAFHPlGDBu5Kr0sJe0y8JkHFDE0t9wOQXLg"
ADMIN_ID = 5859090443

# إعدادات التسجيل
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# --- 1. قاعدة البيانات (بنفس الاسم القديم لضمان الاستقرار) ---
def init_db():
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    
    # جدول المستخدمين والأرصدة
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            balance INTEGER DEFAULT 5
        )
    ''')
    
    # جدول السياق النشط والذاكرة
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS user_active_context (
            user_id INTEGER PRIMARY KEY,
            last_topic TEXT
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

def set_active_context(user_id, topic):
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    cursor.execute('INSERT OR REPLACE INTO user_active_context (user_id, last_topic) VALUES (?, ?)', (user_id, topic))
    conn.commit()
    conn.close()

def get_active_context(user_id):
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    cursor.execute('SELECT last_topic FROM user_active_context WHERE user_id = ?', (user_id,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else None

# دالة لضبط اتجاه النص لليمين (RTL) في ملفات الـ Word
def set_rtl(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = 2

# --- 2. معالج الدردشة الذكية الحرة ---
async def handle_free_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_message = update.message.text
    
    balance = get_or_create_user(user.id, user.username)
    msg_lower = user_message.lower()
    
    # فحص الرصيد
    if "رصيد" in msg_lower or "كم معي" in msg_lower:
        await update.message.reply_text(f"💰 رصيدك الحالي يا أبو عدي هو: **{balance} عملة**", parse_mode="Markdown")
        return

    # طلب إنشاء تقرير أو مستند أو جدول
    if any(word in msg_lower for word in ["تقرير", "ملف", "اعملي", "صمم", "جدول"]):
        if balance < 1:
            await update.message.reply_text("❌ رصيدك غير كافي! تحتاج إلى عملة واحدة على الأقل لتوليد المستندات.")
            return

        update_balance(user.id, -1)
        set_active_context(user.id, user_message)
        
        await update.message.reply_text("🤖 جاري تحليل طلبك، تجهيز الهيكلية، وتصميم المستند الاحترافي مع الجداول...")

        doc = Document()
        
        p1 = doc.add_heading('المستند الذكي الاحترافي', 0)
        set_rtl(p1)
        
        p2 = doc.add_heading('موضوع الطلب:', level=1)
        set_rtl(p2)
        
        p3 = doc.add_paragraph(user_message)
        set_rtl(p3)
        
        # إضافة جدول احترافي منسق
        p_table_title = doc.add_heading('جدول البيانات والتحليل:', level=1)
        set_rtl(p_table_title)
        
        table = doc.add_table(rows=3, cols=3)
        table.style = 'Table Grid'
        
        headers = ["البند / البيان", "التفاصيل", "الحالة"]
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            set_rtl(cell.paragraphs[0])
            
        row1 = ["البند الرئيسي الأول", "تمت المراجعة والتنفيذ", "جاهز"]
        for i, val in enumerate(row1):
            cell = table.cell(1, i)
            cell.text = val
            set_rtl(cell.paragraphs[0])

        row2 = ["التوصيات والتحليل", "مطابق للمعايير", "معتمد"]
        for i, val in enumerate(row2):
            cell = table.cell(2, i)
            cell.text = val
            set_rtl(cell.paragraphs[0])

        file_path = f"smart_doc_{user.id}.docx"
        doc.save(file_path)

        with open(file_path, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename="Professional_Document.docx",
                caption="✅ تفضل المستند الاحترافي جاهز مع الجداول وتنسيق RTL السليم!"
            )

        if os.path.exists(file_path):
            os.remove(file_path)
        return

    # طلب التعديل بناءً على السياق السابق
    if any(word in msg_lower for word in ["عدل", "ضيف", "احذف", "غير"]):
        last_topic = get_active_context(user.id)
        
        if not last_topic:
            await update.message.reply_text("🤔 ما في طلب سابق مسجل بالذاكرة لحتى أعدل عليه. اطلب مني إنشاء ملف أولاً!")
            return
            
        await update.message.reply_text(f"🔄 جاري تعديل المستند السابق ('{last_topic}') بناءً على طلبك: '{user_message}'...")
        
        doc = Document()
        p1 = doc.add_heading('المستند المحدث ذكياً', 0)
        set_rtl(p1)
        
        p2 = doc.add_heading('السياق الأصلي:', level=1)
        set_rtl(p2)
        p3 = doc.add_paragraph(last_topic)
        set_rtl(p3)
        
        p4 = doc.add_heading('التعديلات المضافة:', level=1)
        set_rtl(p4)
        p5 = doc.add_paragraph(user_message)
        set_rtl(p5)
        
        file_path = f"updated_doc_{user.id}.docx"
        doc.save(file_path)
        
        with open(file_path, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename="Updated_Document.docx",
                caption="✅ تفضل النسخة المحدثة بناءً على طلبك!"
            )
            
        if os.path.exists(file_path):
            os.remove(file_path)
        return

    # دردشة عامة
    await update.message.reply_text(f"أهلاً يا أبو عدي! أنا معك. فهمت أنك تقصد: \"{user_message}\". هل تحب أن أبدأ بتوليد مستند أو تقرير بناءً على هذا الكلام؟")

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    balance = get_or_create_user(user.id, user.username)
    welcome_msg = (
        f"أهلاً بك يا {user.first_name} في المساعد الذكي للمستندات والجداول 🚀\n\n"
        f"💰 رصيدك: **{balance} عملة**\n\n"
        f"💬 **تحدث معي مباشرة بالدردشة الحرة:**\n"
        f"• قل لي: *'اعملي تقرير مبيعات'* وسأقوم بتصميمه فوراً.\n"
        f"• أو: *'أضف لي كذا'* وسأقوم بالتعديل على السياق بسلاسة!"
    )
    await update.message.reply_text(welcome_msg, parse_mode="Markdown")

# --- 3. التشغيل ---
def main():
    init_db()
    application = Application.builder().token(TOKEN).build()

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_free_message))

    print("Smart AI Document Assistant is running...")
    application.run_polling()

if __name__ == "__main__":
    main()
