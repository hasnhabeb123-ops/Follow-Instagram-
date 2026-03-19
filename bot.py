
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes

TOKEN = "8789485311:AAFHPlGDBu5Kr0sJe0y8JkHFDE0t9wOQXLg"
ADMIN_ID = 5859090443

# --- أسعار المتابعين ---
PRICES = {f"{i}_followers": f"{i*20}" for i in range(100, 10001, 100)}
PAGE_SIZE = 10

orders = {}
order_id_counter = 1

# --- Pagination ---
def get_page_keyboard(page=0):
    keys = list(PRICES.keys())
    start = page * PAGE_SIZE
    end = start + PAGE_SIZE
    kb = []
    for key in keys[start:end]:
        kb.append([InlineKeyboardButton(f"{key.split('_')[0]} متابع - {PRICES[key]} ليرة", callback_data=key)])
    nav = []
    if start > 0:
        nav.append(InlineKeyboardButton("⬅️", callback_data=f"prev_{page}"))
    if end < len(keys):
        nav.append(InlineKeyboardButton("➡️", callback_data=f"next_{page}"))
    if nav:
        kb.append(nav)
    return InlineKeyboardMarkup(kb)

# --- بدء ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # زر البداية
    kb_menu = [[KeyboardButton("/start")]]
    reply_m = ReplyKeyboardMarkup(kb_menu, resize_keyboard=True)

    await update.message.reply_text("اختر العرض:", reply_markup=get_page_keyboard(0))
    await update.message.reply_text("اضغط على /start للرجوع للبداية في أي وقت 👇", reply_markup=reply_m)

# --- أزرار ---
async def buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    # Pagination
    if data.startswith("next_") or data.startswith("prev_"):
        page = int(data.split("_")[1])
        page = page+1 if "next" in data else page-1
        await query.edit_message_reply_markup(reply_markup=get_page_keyboard(page))
        return

    # اختيار خدمة
    if "_followers" in data:
        context.user_data["service"] = data
        await query.edit_message_text("📎 أرسل رابط الحساب:")
        return

    # اختيار الدفع
    if data.startswith("pay_"):
        context.user_data["payment"] = data

        if data == "pay_sh":
            text = (
                "💰 الدفع عبر شام كاش\n"
                "حوّل المبلغ على الكود:\n"
                "`58c2f6fb44d2c954f450a8ed456dc5ab`\n"
                "📸 بعد التحويل:\n"
                "- أرسل صورة الإيصال\n"
                "- أرسل رقم العملية\n"
                "⚠️ أي خطأ بالمعلومات قد يؤدي لفقدان المبلغ"
            )
        else:  # pay_sy
            text = (
                "📱 الدفع عبر سيريتيل كاش\n"
                "حوّل المبلغ إلى الرقم:\n"
                "`0985718680`\n"
                "📸 بعد التحويل:\n"
                "- أرسل صورة الإيصال\n"
                "- أرسل رقم العملية\n"
                "⚠️ أي خطأ بالمعلومات قد يؤدي لفقدان المبلغ"
            )
        await query.edit_message_text(text, parse_mode='Markdown')
        return

    # أزرار الأدمن
    if data.startswith(("accept_", "reject_", "done_", "del_")):
        order_id = int(data.split("_")[1])
        if order_id not in orders:
            return

        if data.startswith("accept"):
            orders[order_id]["status"] = "تم الدفع"
            await query.edit_message_caption(query.message.caption + "\n✅ تم تأكيد الدفع")
            await context.bot.send_message(
                chat_id=orders[order_id]["user"],
                text=f"✅ تم تأكيد دفعك لطلب #{order_id}"
            )
        elif data.startswith("reject"):
            orders[order_id]["status"] = "مرفوض"
            await query.edit_message_caption(query.message.caption + "\n❌ تم رفض الدفع")
            await context.bot.send_message(
                chat_id=orders[order_id]["user"],
                text=f"❌ تم رفض طلب #{order_id} — أي خطأ بالمعلومات قد يؤدي لفقدان المبلغ"
            )
        elif data.startswith("done"):
            orders[order_id]["status"] = "تم التنفيذ"
            await query.edit_message_caption(query.message.caption + "\n🚀 تم التنفيذ")
            await context.bot.send_message(
                chat_id=orders[order_id]["user"],
                text=f"🚀 تم تنفيذ طلبك #{order_id}"
            )
        elif data.startswith("del"):
            orders[order_id]["status"] = "تم التسليم"
            await query.edit_message_caption(query.message.caption + "\n📦 تم التسليم")
            await context.bot.send_message(
                chat_id=orders[order_id]["user"],
                text=f"📦 تم تسليم طلبك #{order_id}"
            )
        return

# --- استقبال النص ---
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "service" in context.user_data and "link" not in context.user_data:
        context.user_data["link"] = update.message.text
        kb = [
            [InlineKeyboardButton("💰 شام كاش", callback_data="pay_sh")],
            [InlineKeyboardButton("📱 سيريتيل كاش", callback_data="pay_sy")]
        ]
        await update.message.reply_text("اختر طريقة الدفع:", reply_markup=InlineKeyboardMarkup(kb))
        return
    elif "payment" in context.user_data and "transaction" not in context.user_data:
        context.user_data["transaction"] = update.message.text
        await update.message.reply_text("📸 الآن أرسل صورة الإيصال:")
        return
    else:
        await update.message.reply_text("⚠️ أرسل صورة الإيصال فقط")

# --- استقبال الصورة ---
async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global order_id_counter
    if "transaction" not in context.user_data:
        await update.message.reply_text("⚠️ أرسل رقم العملية أولاً")
        return

    order_id = order_id_counter
    order_id_counter += 1

    orders[order_id] = {
        "user": update.message.from_user.id,
        "name": update.message.from_user.full_name,
        "service": context.user_data["service"],
        "link": context.user_data["link"],
        "payment": context.user_data["payment"],
        "transaction": context.user_data["transaction"],
        "status": "بانتظار الدفع"
    }

    caption = (
        f"🚨 طلب جديد #{order_id}\n"
        f"👤 {orders[order_id]['name']}\n"
        f"📦 خدمة: {orders[order_id]['service']}\n"
        f"🔗 رابط: {orders[order_id]['link']}\n"
        f"💳 العملية: {orders[order_id]['transaction']}\n"
        f"💰 الدفع: {'شام كاش' if context.user_data['payment']=='pay_sh' else 'سيريتيل كاش'}"
    )

    kb = [[
        InlineKeyboardButton("✅ تأكيد الدفع", callback_data=f"accept_{order_id}"),
        InlineKeyboardButton("❌ رفض الدفع", callback_data=f"reject_{order_id}")
    ],[
        InlineKeyboardButton("🚀 تم التنفيذ", callback_data=f"done_{order_id}"),
        InlineKeyboardButton("📦 تم التسليم", callback_data=f"del_{order_id}")
    ]]

    await context.bot.send_photo(
        chat_id=ADMIN_ID,
        photo=update.message.photo[-1].file_id,
        caption=caption,
        reply_markup=InlineKeyboardMarkup(kb)
    )

    await update.message.reply_text("⏳ تم إرسال طلبك، بانتظار التأكيد")
    context.user_data.clear()

# --- تشغيل البوت ---
def main():
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(buttons))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    print("🔥 البوت الكامل شغال")
    app.run_polling()

if __name__ == "__main__":
    main()