import logging
import sqlite3
import os
import re
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, ContextTypes, filters
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# التوكن الخاص ببوتك
TOKEN = "8789485311:AAFHPlGDBu5Kr0sJe0y8JkHFDE0t9wOQXLg"
ADMIN_ID = 5859090443

# إعدادات التسجيل
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# --- 1. قاعدة البيانات (SQLite) ---
def init_db():
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            balance INTEGER DEFAULT 10
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS user_active_context (
            user_id INTEGER PRIMARY KEY,
            last_prompt TEXT
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
        cursor.execute('INSERT INTO users (user_id, username, balance) VALUES (?, ?, ?)', (user_id, username, 10))
        conn.commit()
        balance = 10
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

def save_context(user_id, prompt):
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    cursor.execute('INSERT OR REPLACE INTO user_active_context (user_id, last_prompt) VALUES (?, ?)', (user_id, prompt))
    conn.commit()
    conn.close()

def get_context(user_id):
    conn = sqlite3.connect('bot_database.db')
    cursor = conn.cursor()
    cursor.execute('SELECT last_prompt FROM user_active_context WHERE user_id = ?', (user_id,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else None

# دالة ضبط اتجاه النص RTL لملفات الوورد
def set_rtl(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = 2

# دالة استخراج الأعمدة ديناميكياً من أي نص يكتبه المستخدم
def extract_dynamic_columns(text):
    # تقسيم النص بناءً على الأسطر أو الفواصل أو الواو العاطفة
    # تنظيف النص واستخراج الكلمات المفتاحية أو الأسطر كأعمدة للجدول
    lines = text.split('\n')
    columns = []
    
    for line in lines:
        line = line.strip()
        # تخطي الجمل الإنشائية الطويلة والتركيز على الخانات المطلوبة
        if any(w in line for w in ["عملي", "جدول", "فيه", "مقسم", "خانانات", "خانات", "ملف"]):
            # استخراج ما بعد الكلمات أو تفكيك السطر
            cleaned = re.sub(r'^(عملي|سوي|صمم|اعملي|جدول|فيه|مقسم|لـ|اربع|ست|خانات|حط)\s*', '', line).strip()
            if cleaned and len(cleaned) < 40:
                columns.append(cleaned)
        elif line and len(line) < 30:
            columns.append(line)
            
    # لو ما لقى أسطر مفصولة، يفصل بناءً على الفواصل أو الواو
    if len(columns) <= 1:
        # البحث عن الكلمات المفصولة بفواصل أو واو
        tokens = re.split(r'[,،\n]|\s+و\s+', text)
        for t in tokens:
            t = t.strip()
            if t and not any(w in t for w in ["جدول", "ملف", "خانات", "عملي", "فيه", "مقسم"]):
                if len(t) < 25:
                    columns.append(t)
                    
    # لو ما قدر يستخرج شي، ياخد كلمات افتتاحية أو الكلمات الموجودة
    if not columns or len(columns) < 2:
        columns = ["البيان / الحقل الأول", "البيان / الحقل الثاني", "البيان / الحقل الثالث"]
        
    return columns[:8] # حد أقصى 8 أعمدة لضمان جمالية الجدول

# --- 2. معالج المحادثة الحر بالكامل ---
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    text = update.message.text
    balance = get_or_create_user(user.id, user.username)
    text_lower = text.lower()

    # الاستعلام عن الرصيد
    if any(w in text_lower for w in ["رصيد", "كم معي", "رصيدي"]):
        await update.message.reply_text(f"💰 رصيدك الحالي يا {user.first_name} هو: **{balance} عملة**", parse_mode="Markdown")
        return

    # معالجة طلب إنشاء الجداول أو المستندات الحرة
    if any(w in text_lower for w in ["جدول", "ملف", "تقرير", "اعملي", "صمم", "سوي", "حط"]):
        if balance < 1:
            await update.message.reply_text("❌ رصيدك غير كافي! تحتاج إلى عملة واحدة لتوليد المستند.")
            return

        update_balance(user.id, -1)
        save_context(user.id, text)

        await update.message.reply_text(f"🤖 أهلاً بك يا {user.first_name}، جاري تحليل طلبك الحر وتصميم جدولك بالخانات المطلوبة...")

        # استخراج الأعمدة ديناميكياً بشكل كامل
        columns = extract_dynamic_columns(text)

        doc = Document()
        
        # العنوان والمحتوى
        p1 = doc.add_heading('مستند الذكاء الحر', 0)
        set_rtl(p1)
        
        p2 = doc.add_paragraph(f"الطلب الأساسي: {text}")
        set_rtl(p2)

        p_title = doc.add_heading('الجدول الديناميكي المطلوب:', level=1)
        set_rtl(p_title)

        # بناء الجدول بناءً على عدد الأعمدة المستخرجة ديناميكياً
        num_cols = len(columns)
        table = doc.add_table(rows=3, cols=num_cols)
        table.style = 'Table Grid'

        # تعبئة العناوين (الهيدر)
        for i, col_name in enumerate(columns):
            cell = table.cell(0, i)
            cell.text = col_name
            set_rtl(cell.paragraphs[0])

        # صف تجريبي أول
        for i in range(num_cols):
            cell = table.cell(1, i)
            cell.text = f"بيان {i+1}"
            set_rtl(cell.paragraphs[0])

        # صف تجريبي ثاني (فارغ للتعبئة)
        for i in range(num_cols):
            cell = table.cell(2, i)
            cell.text = "..."
            set_rtl(cell.paragraphs[0])

        file_path = f"free_doc_{user.id}.docx"
        doc.save(file_path)

        with open(file_path, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename="Dynamic_Table.docx",
                caption=f"✅ تفضل يا {user.first_name}! تم إنشاء الجدول بـ ({num_cols}) خانات بناءً على طلبك الحر وباتجاه RTL."
            )

        if os.path.exists(file_path):
            os.remove(file_path)
        return

    # التعديل الذكي الحر بناءً على السياق السابق
    if any(w in text_lower for w in ["عدل", "ضيف", "احذف", "غير", "ناقص", "حط"]):
        last_prompt = get_context(user.id)
        if not last_prompt:
            await update.message.reply_text("🤔 ما في طلب سابق مسجل لحتى أعدل عليه. اطلب مني جدولاً أو ملفاً أولاً!")
            return

        combined_text = last_prompt + " " + text
        await update.message.reply_text("🔄 جاري دمج التعديل وتحديث الجدول بناءً على إضافتك...")

        columns = extract_dynamic_columns(combined_text)

        doc = Document()
        p1 = doc.add_heading('المستند المحدث حرّاً', 0)
        set_rtl(p1)
        p2 = doc.add_paragraph
