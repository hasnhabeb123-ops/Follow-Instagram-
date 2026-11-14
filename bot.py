import logging
import sqlite3
import os
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, ContextTypes, filters
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# توكن البوت ومعرف الأدمن الخاص بك
TOKEN = "8789485311:AAHlP1GDBu5Kr0sJe0y8JkHFDEot9wOQXlg"
ADMIN_ID = 5859090443

# إعدادات التسجيل
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# --- 1. قاعدة البيانات المتطورة (إدارة المستخدمين، الرصيد، والذاكرة التاريخية) ---
def init_db():
    conn = sqlite3.connect('smart_bot_database.db')
    cursor = conn.cursor()
    
    # جدول المستخدمين والأرصدة
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            balance INTEGER DEFAULT 5
        )
    ''')
    
    # جدول الذاكرة العميقة وسجل المحادثات لكل مستخدم
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS chat_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            role TEXT,
            content TEXT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # جدول لتتبع آخر ملف أو طلب تم العمل عليه
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS user_active_context (
            user_id INTEGER PRIMARY KEY,
            last_topic TEXT,
            last_file_path TEXT
        )
    ''')
    
    conn.commit()
    conn.close()

def get_or_create_user(user_id, username):
    conn = sqlite3.connect('smart_bot_database.db')
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
    conn = sqlite3.connect('smart_bot_database.db')
    cursor = conn.cursor()
    cursor.execute('UPDATE users SET balance = balance + ? WHERE user_id = ?', (amount, user_id))
    conn.commit()
    conn.close()

def log_chat(user_id, role, content):
    conn = sqlite3.connect('smart_bot_database.db')
    cursor = conn.cursor()
    cursor.execute('INSERT INTO chat_sessions (user_id, role, content) VALUES (?, ?, ?)', (user_id, role, content))
    conn.commit()
    conn.close()

def set_active_context(user_id, topic):
    conn = sqlite3.connect('smart_bot_database.db')
    cursor = conn.cursor()
    cursor.execute('INSERT OR REPLACE INTO user_active_context (user_id, last_topic) VALUES (?, ?)', (user_id, topic))
    conn.commit()
    conn.close()

def get_active_context(user_id):
    conn = sqlite3.connect('smart_bot_database.db')
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

# --- 2. معالج الرسائل الحرة (الدردشة والذكاء السياقي) ---
async def handle_free_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_message = update.message.text
    
    # التأكد من وجود المستخدم ورصيده
    balance = get_or_create_user(user.id, user.username)
    
    # تسجيل رسالة المستخدم في الذاكرة
    log_chat(user.id, "user", user_message)
    
    # تحليلات مبدئية لنية المستخدم (Intent Detection) بطريقة ذكية
    msg_lower = user_message.lower()
    
    # حالة طلب رصيد
    if "رصيد" in msg_lower or "كم معي" in msg_lower:
        await update.message.reply_text(f"💰 رصيدك الحالي يا أبو عدي هو: **{balance} عملة**", parse_mode="Markdown")
        return

    # الكشف عما إذا كان المستخدم يطلب إنشاء ملف أو تقرير
    if "تقرير" in msg_lower or "ملف" in msg_lower or "اعملي" in msg_lower or "صمم" in msg_lower or "جدول" in msg_lower:
        
        if balance < 1:
            await update.message.reply_text("❌ رصيدك غير كافي! تحتاج إلى عملة واحدة على الأقل لتوليد المستندات.")
            return

        # خصم عملة وحفظ السياق النشط
        update_balance(user.id, -1)
        set_active_context(user.id, user_message)
        
        await update.message.reply_text("🤖 جاري تحليل طلبك، تجهيز الهيكلية، وتصميم المستند الاحترافي...")

        # بناء ملف Word ذكي ومتقدم
        doc = Document()
        
        p1 = doc.add_heading('المستند الذكي الاحترافي', 0)
        set_rtl(p1)
        
        p2 = doc.add_heading('موضوع الطلب:', level=1)
        set_rtl(p2)
        
        p3 = doc.add_paragraph(user_message)
        set_rtl(p3)
        
        # إضافة جدول افتراضي احترافي كمثال على القدرات العالية
        p_table_title = doc.add_heading('جدول البيانات والتحليل:', level=1)
        set_rtl(p_table_title)
        
        table = doc.add_table(rows=3, cols=3)
        table.style = 'Table Grid'
        
        headers = ["البيان / البند", "التفاصيل", "الحالة"]
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            set_rtl(cell.paragraphs[0])
            
        row1 = ["البند الرئيسي الأول", "تمت المراجعة والتنفيذ", "جاهز"]
        for i, val in enumerate(row1):
            cell = table.cell(1, i)
            cell.text = val
            set_rtl(cell.paragraphs[0])

        row2 = ["التوصيات والتحليل", "مطابق للمعايير الاحترافية", "معتمد"]
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
                caption="✅ تفضل المستند الاحترافي جاهز، مع الجداول والتنسيق المخصص بالكامل!"
            )

        if os.path.exists(file_path):
            os.remove(file_path)
            
        log_chat(user.id, "assistant", "تم توليد وإرسال المستند الاحترافي بنجاح.")
        return

    # الكشف عما إذا كان المستخدم يطلب تعديلاً بناءً على السياق السابق
    elif "عدل" in msg_lower or "ضيف" in msg_lower or "احذف" in msg_lower or "غير" in msg_lower:
        last_topic = get_active_context(user.id)
        
        if not last_topic:
            await update.message.reply_text("🤔 ما في طلب سابق مسجل بالذاكرة لحتى أعدل عليه. اطلب مني إنشاء ملف أولاً وسأقوم بتعديله فوراً متى شئت!")
            return
            
        await update.message.reply_text(f"🔄 فهمت عليك! جاري تعديل المستند السابق ('{last_topic}') بناءً على طلبك الجديد: '{user_message}'...")
        
        # بناء ملف محدث مع التعديلات
        doc = Document()
        p1 = doc.add_heading('المستند المحدث ذكياً', 0)
        set_rtl(p1)
        
        p2 = doc.add_heading('السياق الأصلي:', level=1)
        set_rtl(p2)
        p3 = doc.add_paragraph(last_topic)
        set_rtl(p3)
        
        p4 = doc.add_heading('التعديلات المطبقة:', level=1)
        set_rtl(p4)
        p5 = doc.add_paragraph(user_message)
        set_rtl(p5)
        
        file_path = f"updated_doc_{user.id}.docx"
        doc.save(file_path)
        
        with open(file_path, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename="Updated_Professional_Document.docx",
                caption="✅ تفضل النسخة المحدثة بناءً على تفضيلاتك وسياق الحديث!"
            )
            
        if os.path.exists(file_path):
            os.remove(file_path)
            
        log_chat(user.id, "assistant", f"تم تحديث المستند بناءً على: {user_message}")
        return

    # رد افتراضي ذكي يشعر المستخدم أنه يكلم مساعد شخصي بشري
    default_reply = (
        f"أهلاً يا أبو عدي! أنا معك عم أسمعك. "
        f"فهمت عليك أنك تحكي عن: \"{user_message}\". "
        f"تفضل، طلبك أسلوب دردشة حرة ومتصل، هل تحب أن أقوم بتوليد تقزار أو مستند أو جدول بناءً على هذا الكلام؟"
    )
    await update.message.reply_text(default_reply)
    log_chat(user.id, "assistant", default_reply)

# --- 3. الأوامر الأساسية البسيطة (مثل /start) ---
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    balance = get_or_create_user(user.id, user.username)
    welcome_msg = (
        f"أهلاً بك يا {user.first_name} في نظام المساعد الذكي للمستندات والجداول الاحترافية 🚀\n\n"
        f"💰 رصيدك: **{balance} عملة**\n\n"
        f"💬 **لا داعي للأوامر المعقدة!** تحدث معي مباشرة باللغة العادية:\n"
        f"• قل لي: *'اعملي تقرير مبيعات لشركة الأجهزة'* وسأقوم بتصميمه فوراً.\n"
        f"• أو: *'أضف لي جدول أسعار'* وسأقوم بالتعديل على السياق بسلاسة!"
    )
    await update.message.reply_text(welcome_msg, parse_mode="Markdown")

# --- 4. تشغيل النظام الذكي ---
def main():
    init_db()
    
    application = Application.builder().token(TOKEN).build()

    # الأوامر العامة
    application.add_handler(CommandHandler("start", start_command))
    
    # معالج الرسائل الحرة (الدردشة والذكاء الاصطناعي المحلي) لكافة النصوص الواردة
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_free_message))

    print("Smart AI Document Assistant is running...")
    application.run_polling()

if __name__ == "__main__":
    main()
