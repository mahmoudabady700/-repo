"""
بوت تيليجرام — سيستم إدارة باقات الإنترنت محمود / بودي
=====================================================
المتطلبات:
    pip install python-telegram-bot gspread google-auth
"""

import logging
import re
from datetime import datetime, timedelta
from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup,
    ReplyKeyboardMarkup, KeyboardButton
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler, ContextTypes,
    filters
)
import gspread
from google.oauth2.service_account import Credentials

# ══════════════════════════════════════════════
# إعدادات — عدّل هنا فقط
# ══════════════════════════════════════════════
import os
BOT_TOKEN = os.environ["BOT_TOKEN"]
SPREADSHEET_ID = "1zJF0_hdbgh63NcwuA5Qz0J-dYfMdQbeQ5BBnWMi3L5A"
ALLOWED_USERS  = [6843334319]

SHEET_CLIENTS  = " سجل العملاء"
SHEET_CONFIG   = "الإعدادات"
SHEET_CASH     = " تحويل الرصيد"

DATA_START_ROW = 4
TOTAL_ROW      = 503

# ══════════════════════════════════════════════
# Logging
# ══════════════════════════════════════════════
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════
# States
# ══════════════════════════════════════════════
(
    STATE_MAIN,
    # إضافة عميل
    STATE_ADD_TYPE, STATE_ADD_NAME, STATE_ADD_PHONE,
    STATE_ADD_PKG, STATE_ADD_GIGA, STATE_ADD_NOTES, STATE_ADD_CONFIRM,
    # تسجيل دفعة
    STATE_PAY_SEARCH, STATE_PAY_SELECT, STATE_PAY_CONFIRM,
    # تحويل رصيد للمدير
    STATE_TRANSFER_AMOUNT, STATE_TRANSFER_METHOD, STATE_TRANSFER_CONFIRM,
    # استلام رصيد من المدير
    STATE_RECEIVE_AMOUNT, STATE_RECEIVE_NOTE, STATE_RECEIVE_CONFIRM,
    # بحث عميل
    STATE_SEARCH_CLIENT,
) = range(18)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ══════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════
def to_int(val):
    if not val: return 0
    nums = re.findall(r'\d+', str(val))
    return int(nums[0]) if nums else 0

def to_float(val):
    if not val: return 0.0
    nums = re.findall(r'[\d.]+', str(val))
    return float(nums[0]) if nums else 0.0

def get_sheet(name: str):
    import json, os, base64
    raw = os.environ["GOOGLE_CREDENTIALS_B64"]
    creds_info = json.loads(base64.b64decode(raw).decode())
    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(name)

def is_allowed(update: Update) -> bool:
    return update.effective_user.id in ALLOWED_USERS

# ══════════════════════════════════════════════
# Google Sheets — قراءة
# ══════════════════════════════════════════════
def get_packages():
    ws = get_sheet(SHEET_CONFIG)
    rows = ws.get_values("A5:F11")
    pkgs = []
    for row in rows:
        if len(row) >= 2 and row[0]:
            price = to_int(row[1])
            comm  = to_int(row[4]) if len(row) > 4 else 20
            pkgs.append({
                "name":  row[0],
                "price": price,
                "giga":  to_int(row[2]) if len(row) > 2 else 0,
                "days":  to_int(row[3]) if len(row) > 3 else 31,
                "comm":  comm,
                "sell":  price + comm,
            })
    return pkgs

def get_extra_config():
    ws = get_sheet(SHEET_CONFIG)
    return {
        "price_per_g": to_float(ws.acell("B15").value) or 3,
        "comm":        to_float(ws.acell("B16").value) or 20,
    }

def get_transfer_config():
    """بيانات المدير من الإعدادات"""
    ws = get_sheet(SHEET_CONFIG)
    return {
        "name":   ws.acell("B19").value or "المدير",
        "phone":  ws.acell("B20").value or "",
        "method": ws.acell("B21").value or "محفظة",
    }

def get_all_clients():
    ws = get_sheet(SHEET_CLIENTS)
    rows = ws.get_values(f"A{DATA_START_ROW}:M200")
    clients = []
    for i, row in enumerate(rows, start=DATA_START_ROW):
        if len(row) > 1 and row[1]:
            clients.append({
                "row":        i,
                "date":       row[0]  if len(row) > 0  else "",
                "name":       row[1]  if len(row) > 1  else "",
                "phone":      row[2]  if len(row) > 2  else "",
                "package":    row[3]  if len(row) > 3  else "",
                "giga":       row[4]  if len(row) > 4  else "",
                "balance":    row[6]  if len(row) > 6  else "",
                "sell":       row[8]  if len(row) > 8  else "",
                "expiry":     row[9]  if len(row) > 9  else "",
                "pay_status": row[10] if len(row) > 10 else "غير مدفوع",
                "pay_date":   row[11] if len(row) > 11 else "",
                "notes":      row[12] if len(row) > 12 else "",
            })
    return clients

def get_summary():
    ws = get_sheet(SHEET_CASH)
    return {
        "received": to_float(ws.acell("H4").value),
        "charged":  to_float(ws.acell("H5").value),
        "transfer": to_float(ws.acell("H6").value),
        "cash":     to_float(ws.acell("H7").value),
        "we_bal":   to_float(ws.acell("H8").value),
    }

# ══════════════════════════════════════════════
# Google Sheets — كتابة
# ══════════════════════════════════════════════
def find_next_empty_row(ws) -> int:
    col_b = ws.col_values(2)
    for i, val in enumerate(col_b[DATA_START_ROW - 1:], start=DATA_START_ROW):
        if not val and i != TOTAL_ROW:
            return i
    return len(col_b) + 1

def add_client_row(data: dict) -> int:
    ws = get_sheet(SHEET_CLIENTS)
    row_num = find_next_empty_row(ws)
    today = datetime.now().strftime("%Y-%m-%d")

    if data["type"] == "package":
        pkg    = data["package"]
        expiry = (datetime.strptime(today, "%Y-%m-%d") + timedelta(days=pkg["days"])).strftime("%Y-%m-%d")
        row = [today, data["name"], data["phone"], pkg["name"],
               pkg["giga"], "", pkg["price"], pkg["comm"],
               pkg["price"] + pkg["comm"], expiry,
               "غير مدفوع", "", data.get("notes", "")]
    else:
        cfg  = data["extra_cfg"]
        giga = data["giga"]
        bal  = giga * cfg["price_per_g"]
        sell = bal + cfg["comm"]
        row = [today, data["name"], data["phone"], "شحن اضافي",
               giga, round(giga/5, 1), bal, cfg["comm"],
               sell, "لا ينتهي", "غير مدفوع", "", data.get("notes", "")]

    ws.update(range_name=f"A{row_num}:M{row_num}", values=[row])
    return row_num

def mark_as_paid(row_num: int):
    ws  = get_sheet(SHEET_CLIENTS)
    today = datetime.now().strftime("%Y-%m-%d")
    ws.update(range_name=f"K{row_num}", values=[["مدفوع"]])
    ws.update(range_name=f"L{row_num}", values=[[today]])

def add_transfer_row(amount: float, method: str, note: str = "") -> int:
    """يضيف صف في سجل التحويلات للمدير (D40:G80)"""
    ws = get_sheet(SHEET_CASH)
    cfg = get_transfer_config()
    today = datetime.now().strftime("%Y-%m-%d")

    # ابحث عن أول صف فاضي في عمود D (المبلغ) من 40 لـ 80
    col_d = ws.col_values(4)  # عمود D
    row_num = 40
    for i in range(39, min(80, len(col_d))):
        if not col_d[i]:
            row_num = i + 1
            break
    else:
        row_num = len(col_d) + 1
        if row_num < 40: row_num = 40

    ws.update(range_name=f"B{row_num}", values=[[cfg["name"]]])
    ws.update(range_name=f"C{row_num}", values=[[cfg["phone"]]])
    ws.update(range_name=f"D{row_num}", values=[[amount]])
    ws.update(range_name=f"E{row_num}", values=[[method]])
    ws.update(range_name=f"F{row_num}", values=[["تم التحويل"]])
    ws.update(range_name=f"G{row_num}", values=[[note or ""]])
    return row_num

def add_receive_row(amount: float, note: str = "") -> int:
    """يضيف صف في سجل استلام الرصيد من المدير (B12:B35)"""
    ws = get_sheet(SHEET_CASH)
    today = datetime.now().strftime("%Y-%m-%d")

    col_b = ws.col_values(2)  # عمود B
    row_num = 12
    for i in range(11, min(35, len(col_b))):
        if not col_b[i]:
            row_num = i + 1
            break
    else:
        row_num = 12

    ws.update(range_name=f"B{row_num}", values=[[amount]])
    if note:
        ws.update(range_name=f"C{row_num}", values=[[note]])
    return row_num

# ══════════════════════════════════════════════
# /start — القائمة الرئيسية
# ══════════════════════════════════════════════
MAIN_KEYBOARD = ReplyKeyboardMarkup([
    ["➕ إضافة عميل",     "✅ تسجيل دفعة"],
    ["💸 تحويل للمدير",   "📥 استلام رصيد"],
    ["👥 قائمة العملاء",  "🔍 بحث عميل"],
    ["📊 ملخص الحساب",   "❌ إلغاء"],
], resize_keyboard=True)

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("⛔ عيني عينك يا عم! مش مصرح ليك هنا 😤\nروح نام بقى يمحمود متبقاش رخم 😴")
        return ConversationHandler.END
    ctx.user_data.clear()
    await update.message.reply_text(
        "🌐 *سيستم باقات الإنترنت — محمود / بودي*\n\nاختار العملية ي عبادي:",
        reply_markup=MAIN_KEYBOARD,
        parse_mode="Markdown"
    )
    return STATE_MAIN

async def main_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text

    if text == "➕ إضافة عميل":
        kb = [
            [InlineKeyboardButton("📦 باقة عادية", callback_data="type_package")],
            [InlineKeyboardButton("🚀 شحن إضافي",  callback_data="type_extra")],
        ]
        await update.message.reply_text("!💪 يلا يا حودة اختار نوع الخدمة:",
                                        reply_markup=InlineKeyboardMarkup(kb))
        return STATE_ADD_TYPE

    elif text == "✅ تسجيل دفعة":
        await update.message.reply_text("🔍😎 يلا اكتب اسم العميل أو رقمه، هدور ليك فالحال يباشا ")
        return STATE_PAY_SEARCH

    elif text == "💸 تحويل للمدير":
        s = get_summary()
        cfg = get_transfer_config()
        await update.message.reply_text(
            f"💸 *تحويل رصيد للمدير*\n\n"
            f"👤 المستلم: {cfg['name']}\n"
            f"📱 الرقم: {cfg['phone']}\n"
            f"💳 طريقة التحويل الافتراضية: {cfg['method']}\n\n"
            f"💵 الكاش المتاح عندك: *{s['cash']} ج.م*\n\n"
            f" اكتب المبلغ يا معلم وأنا هسجله فالحال ي عمدة 💸",
            parse_mode="Markdown"
        )
        return STATE_TRANSFER_AMOUNT

    elif text == "📥 استلام رصيد":
        s = get_summary()
        await update.message.reply_text(
            f"📥 *استلام رصيد WE من المدير*\n\n"
            f"📶 رصيد WE الحالي عندك: *{s['we_bal']} ج.م*\n\n"
            f"ممتاز يا عبادي! كام استلمت من المدير؟ 📥",
            parse_mode="Markdown"
        )
        return STATE_RECEIVE_AMOUNT

    elif text == "👥 قائمة العملاء":
        await update.message.reply_text("⏳ لحظة بجيبلك القائمة يا هندسة... 📋")
        clients = get_all_clients()
        if not clients:
            await update.message.reply_text("😅 ماشيش عميل لسه يا معلم! ابدأ بإضافة أول عميل 🚀")
            return STATE_MAIN
        msg = "👥 *آخر 10 عملاء:*\n\n"
        for c in clients[-10:]:
            icon = "🟢" if c["pay_status"] == "مدفوع" else "🔴"
            exp  = c["expiry"] or "—"
            msg += (
                f"{icon} *{c['name']}* — {c['package']}\n"
                f"   📞 {c['phone']} | 💰 {c['sell']} ج.م\n"
                f"   📆 انتهاء: {exp} | {c['pay_status']}\n\n"
            )
        await update.message.reply_text(msg, parse_mode="Markdown")
        return STATE_MAIN

    elif text == "🔍 بحث عميل":
        await update.message.reply_text("🔍 يلا اكتب اسم العميل أو رقمه، هدور ليك فالحال 😎")
        return STATE_SEARCH_CLIENT

    elif text == "📊 ملخص الحساب":
        await update.message.reply_text("⏳ لحظة بجيبلك القائمة يا هندسة... 📋")
        s = get_summary()
        clients = get_all_clients()
        unpaid  = [c for c in clients if c["pay_status"] != "مدفوع"]
        msg = (
            f"📊 *ملخص الحساب*\n\n"
            f"💰 رصيد WE المستلم:    `{s['received']} ج.م`\n"
            f"📶 إجمالي الشحن:        `{s['charged']} ج.م`\n"
            f"📤 محوّل للمدير:        `{s['transfer']} ج.م`\n"
            f"━━━━━━━━━━━━━━━━\n"
            f"💵 الكاش عندك:          `{s['cash']} ج.م`\n"
            f"📶 رصيد WE المتبقي:     `{s['we_bal']} ج.م`\n"
            f"━━━━━━━━━━━━━━━━\n"
            f"🔴 عملاء غير مدفوعين:  `{len(unpaid)} عميل`\n"
        )
        if unpaid:
            msg += "\n*غير مدفوعين:*\n"
            for c in unpaid[:5]:
                msg += f"  • {c['name']} — {c['sell']} ج.م\n"
            if len(unpaid) > 5:
                msg += f"  _و {len(unpaid)-5} أكتر..._\n"
        await update.message.reply_text(msg, parse_mode="Markdown")
        return STATE_MAIN

    elif text == "❌ إلغاء":
        return await start(update, ctx)

    return STATE_MAIN

# ══════════════════════════════════════════════
# إضافة عميل
# ══════════════════════════════════════════════
async def add_type(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ctx.user_data["add_type"] = query.data
    await query.edit_message_text("تمام يا أسطى! 😎 اكتب اسم العميل:")
    return STATE_ADD_NAME

async def add_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["add_name"] = update.message.text.strip()
    await update.message.reply_text("👌 وصلت! دلوقتي اكتب رقم الهاتف يا معلم:")
    return STATE_ADD_PHONE

async def add_phone(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["add_phone"] = update.message.text.strip()
    if ctx.user_data["add_type"] == "type_package":
        await update.message.reply_text("⏳ لحظة يا صبر جميل، جاري جيب الباقات 🚀")
        pkgs = get_packages()
        ctx.user_data["packages"] = pkgs
        kb = [[InlineKeyboardButton(
            f"{p['name']} ◀ {p['sell']} ج.م ({p['giga']}G)",
            callback_data=f"pkg_{i}"
        )] for i, p in enumerate(pkgs)]
        await update.message.reply_text("📦 هوه ده اللي عندنا يا حودة، اختار اللي يعجبك:",
                                        reply_markup=InlineKeyboardMarkup(kb))
        return STATE_ADD_PKG
    else:
        await update.message.reply_text("🌐 تمام يا نور عيني! كام جيجا عايز تشحنله؟")
        return STATE_ADD_GIGA

async def add_pkg(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    pkg = ctx.user_data["packages"][int(query.data.replace("pkg_", ""))]
    ctx.user_data["add_pkg"] = pkg
    await query.edit_message_text(
        f"✅ *{pkg['name']}*\n"
        f"💰 {pkg['price']} ج.م | 🚚 عمولة {pkg['comm']} ج.م | 🏷️ بيع {pkg['sell']} ج.م\n"
        f"🌐 {pkg['giga']}G | 📆 {pkg['days']} يوم\n\n"
        f"📝 عندك ملاحظة على العميل؟\nلو مفيش اكتب /skip يا هندسة 😄",
        parse_mode="Markdown"
    )
    return STATE_ADD_NOTES

async def add_giga(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        giga = float(update.message.text.strip())
        if giga <= 0: raise ValueError
    except:
        await update.message.reply_text("❌ رقم غير صحيح، اكتب عدد الجيجا:")
        return STATE_ADD_GIGA
    cfg  = get_extra_config()
    bal  = giga * cfg["price_per_g"]
    sell = bal + cfg["comm"]
    ctx.user_data.update({"add_giga": giga, "extra_cfg": cfg})
    await update.message.reply_text(
        f"✅ *شحن إضافي*\n"
        f"🌐 {giga}G | كمية {round(giga/5,1)} وحدة\n"
        f"💰 {bal} ج.م | 🚚 عمولة {cfg['comm']} ج.م | 🏷️ بيع {sell} ج.م\n\n"
        f"📝 عندك ملاحظة على العميل؟\nلو مفيش اكتب /skip يا هندسة 😄",
        parse_mode="Markdown"
    )
    return STATE_ADD_NOTES

async def add_notes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    ctx.user_data["add_notes"] = "" if text == "/skip" else text.strip()
    d = ctx.user_data
    if d["add_type"] == "type_package":
        pkg = d["add_pkg"]
        summary = (f"👤 *{d['add_name']}*\n📞 {d['add_phone']}\n"
                   f"📦 {pkg['name']}\n🏷️ {pkg['sell']} ج.م\n"
                   f"📅 ينتهي بعد {pkg['days']} يوم\n"
                   f"📝 {d['add_notes'] or '—'}")
    else:
        cfg  = d["extra_cfg"]
        giga = d["add_giga"]
        sell = giga * cfg["price_per_g"] + cfg["comm"]
        summary = (f"👤 *{d['add_name']}*\n📞 {d['add_phone']}\n"
                   f"🚀 شحن {giga}G\n🏷️ {sell} ج.م\n"
                   f"📝 {d['add_notes'] or '—'}")
    kb = [[
        InlineKeyboardButton("✅ تأكيد", callback_data="confirm_add"),
        InlineKeyboardButton("❌ إلغاء", callback_data="cancel_add"),
    ]]
    await update.message.reply_text(
        f"📋 *ملخص العملية:*\n\n{summary}",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
    return STATE_ADD_CONFIRM

async def add_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "cancel_add":
        await query.edit_message_text("❌ تمام يا هندسة، ألغينا العملية 😄\nكلما تحتاج أنا هنا! 🫡")
        return ConversationHandler.END
    await query.edit_message_text("⏳ ثانية يا عم، بسجل في الشيت 📝")
    d = ctx.user_data
    data = {"type": "package" if d["add_type"] == "type_package" else "extra",
            "name": d["add_name"], "phone": d["add_phone"],
            "notes": d.get("add_notes", "")}
    if d["add_type"] == "type_package":
        data["package"] = d["add_pkg"]
    else:
        data["giga"] = d["add_giga"]
        data["extra_cfg"] = d["extra_cfg"]
    try:
        row_num = add_client_row(data)
        await query.edit_message_text(
            f"✅ *تم إضافة العميل!*\n\n"
            f"👤 {d['add_name']} | صف {row_num}\n"
            f"📅 {datetime.now().strftime('%Y-%m-%d')}\n"
            f"💳 حالة الدفع: غير مدفوع",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(e)
        await query.edit_message_text(f"❌ خطأ: {e}")
    return ConversationHandler.END

# ══════════════════════════════════════════════
# تسجيل دفعة
# ══════════════════════════════════════════════
async def pay_search(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    search  = update.message.text.strip().lower()
    await update.message.reply_text("⏳ بدور عليه دلوقتي يا سيدي... 🔍")
    clients = get_all_clients()
    results = [c for c in clients
               if (search in c["name"].lower() or search in c["phone"])
               and c["pay_status"] != "مدفوع"]
    if not results:
        await update.message.reply_text("🔍 مفيش عملاء غير مدفوعين بالاسم ده.\nجرب تاني أو /cancel.")
        return STATE_PAY_SEARCH
    ctx.user_data["pay_results"] = results
    kb = [[InlineKeyboardButton(
        f"{c['name']} — {c['package']} — {c['sell']} ج.م",
        callback_data=f"pay_{c['row']}"
    )] for c in results[:8]]
    await update.message.reply_text("لقيتهم! اختار العميل يا باشا: 👇", reply_markup=InlineKeyboardMarkup(kb))
    return STATE_PAY_SELECT

async def pay_select(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    row_num = int(query.data.replace("pay_", ""))
    client  = next((c for c in ctx.user_data["pay_results"] if c["row"] == row_num), None)
    if not client:
        await query.edit_message_text("❌ حصل خطأ.")
        return ConversationHandler.END
    ctx.user_data.update({"pay_client": client, "pay_row_num": row_num})
    kb = [[
        InlineKeyboardButton("✅ تأكيد الدفع", callback_data="confirm_pay"),
        InlineKeyboardButton("❌ إلغاء",        callback_data="cancel_pay"),
    ]]
    await query.edit_message_text(
        f"💳 *تأكيد استلام الدفعة*\n\n"
        f"👤 {client['name']}\n📞 {client['phone']}\n"
        f"📦 {client['package']}\n💰 {client['sell']} ج.م\n"
        f"📅 {datetime.now().strftime('%Y-%m-%d')}",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
    return STATE_PAY_CONFIRM

async def pay_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "cancel_pay":
        await query.edit_message_text("❌ تمام يا هندسة، ألغينا العملية 😄\nكلما تحتاج أنا هنا! 🫡")
        return ConversationHandler.END
    await query.edit_message_text("⏳ ثانية بسجل الدفعة يا نجم... 💳")
    try:
        mark_as_paid(ctx.user_data["pay_row_num"])
        c = ctx.user_data["pay_client"]
        await query.edit_message_text(
            f"✅ *تم تسجيل الدفعة!*\n\n"
            f"👤 {c['name']}\n💰 {c['sell']} ج.م\n"
            f"📅 {datetime.now().strftime('%Y-%m-%d')}",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(e)
        await query.edit_message_text(f"❌ خطأ: {e}")
    return ConversationHandler.END

# ══════════════════════════════════════════════
# تحويل رصيد للمدير
# ══════════════════════════════════════════════
async def transfer_amount(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        amount = float(update.message.text.strip())
        if amount <= 0: raise ValueError
    except:
        await update.message.reply_text("❌ رقم غير صحيح، اكتب المبلغ:")
        return STATE_TRANSFER_AMOUNT
    ctx.user_data["transfer_amount"] = amount
    cfg = get_transfer_config()
    kb = [
        [InlineKeyboardButton(cfg["method"],   callback_data=f"method_{cfg['method']}")],
        [InlineKeyboardButton("إنستاباي",       callback_data="method_إنستاباي")],
        [InlineKeyboardButton("تحويل بنكي",     callback_data="method_تحويل بنكي")],
        [InlineKeyboardButton("كاش يد بيد",     callback_data="method_كاش يد بيد")],
    ]
    await update.message.reply_text(
        f"💸 المبلغ: *{amount} ج.م*\n\nاختار طريقة التحويل:",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
    return STATE_TRANSFER_METHOD

async def transfer_method(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    method = query.data.replace("method_", "")
    ctx.user_data["transfer_method"] = method
    cfg    = get_transfer_config()
    amount = ctx.user_data["transfer_amount"]
    s      = get_summary()
    kb = [[
        InlineKeyboardButton("✅ تأكيد التحويل", callback_data="confirm_transfer"),
        InlineKeyboardButton("❌ إلغاء",          callback_data="cancel_transfer"),
    ]]
    await query.edit_message_text(
        f"💸 *ملخص التحويل*\n\n"
        f"👤 المستلم: {cfg['name']}\n"
        f"📱 الرقم: {cfg['phone']}\n"
        f"💳 الطريقة: {method}\n"
        f"💵 المبلغ: *{amount} ج.م*\n\n"
        f"💵 الكاش المتاح: {s['cash']} ج.م\n"
        f"💵 بعد التحويل: {s['cash'] - amount} ج.م",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
    return STATE_TRANSFER_CONFIRM

async def transfer_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "cancel_transfer":
        await query.edit_message_text("❌ تمام يا هندسة، ألغينا العملية 😄\nكلما تحتاج أنا هنا! 🫡")
        return ConversationHandler.END
    await query.edit_message_text("⏳ بسجل التحويل دلوقتي يا هندسة... ✍️")
    try:
        amount = ctx.user_data["transfer_amount"]
        method = ctx.user_data["transfer_method"]
        cfg    = get_transfer_config()
        row_num = add_transfer_row(amount, method)
        await query.edit_message_text(
            f"✅ *تم تسجيل التحويل!*\n\n"
            f"👤 {cfg['name']} — {cfg['phone']}\n"
            f"💵 {amount} ج.م\n"
            f"💳 {method}\n"
            f"📅 {datetime.now().strftime('%Y-%m-%d')}\n"
            f"📋 صف {row_num} في شيت تحويل الرصيد",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(e)
        await query.edit_message_text(f"❌ خطأ: {e}")
    return ConversationHandler.END

# ══════════════════════════════════════════════
# استلام رصيد من المدير
# ══════════════════════════════════════════════
async def receive_amount(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        amount = float(update.message.text.strip())
        if amount <= 0: raise ValueError
    except:
        await update.message.reply_text("❌ رقم غير صحيح، اكتب المبلغ:")
        return STATE_RECEIVE_AMOUNT
    ctx.user_data["receive_amount"] = amount
    await update.message.reply_text(
        f"📥 المبلغ المستلم: *{amount} ج.م*\n\n"
        f"📝 اكتب ملاحظة (مثلاً: واتساب / يد بيد) أو /skip:",
        parse_mode="Markdown"
    )
    return STATE_RECEIVE_NOTE

async def receive_note(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    ctx.user_data["receive_note"] = "" if text == "/skip" else text.strip()
    amount = ctx.user_data["receive_amount"]
    s      = get_summary()
    kb = [[
        InlineKeyboardButton("✅ تأكيد", callback_data="confirm_receive"),
        InlineKeyboardButton("❌ إلغاء", callback_data="cancel_receive"),
    ]]
    await update.message.reply_text(
        f"📥 *تأكيد استلام الرصيد*\n\n"
        f"💵 المبلغ: *{amount} ج.م*\n"
        f"📝 {ctx.user_data['receive_note'] or '—'}\n\n"
        f"📶 رصيد WE بعد الاستلام: {s['we_bal'] + amount} ج.م",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
    return STATE_RECEIVE_CONFIRM

async def receive_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "cancel_receive":
        await query.edit_message_text("❌ تمام يا هندسة، ألغينا العملية 😄\nكلما تحتاج أنا هنا! 🫡")
        return ConversationHandler.END
    await query.edit_message_text("⏳ بسجله دلوقتي... 📝")
    try:
        amount  = ctx.user_data["receive_amount"]
        note    = ctx.user_data.get("receive_note", "")
        row_num = add_receive_row(amount, note)
        await query.edit_message_text(
            f"✅ *تم تسجيل الاستلام!*\n\n"
            f"💵 {amount} ج.م\n"
            f"📅 {datetime.now().strftime('%Y-%m-%d')}\n"
            f"📋 صف {row_num} في شيت تحويل الرصيد",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(e)
        await query.edit_message_text(f"❌ خطأ: {e}")
    return ConversationHandler.END

# ══════════════════════════════════════════════
# بحث عميل
# ══════════════════════════════════════════════
async def search_client(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    search  = update.message.text.strip().lower()
    await update.message.reply_text("⏳ بدور عليه دلوقتي يا سيدي... 🔍")
    clients = get_all_clients()
    results = [c for c in clients
               if search in c["name"].lower() or search in c["phone"]]
    if not results:
        await update.message.reply_text("🤷 والله مش لاقيش حد بالاسم أو الرقم ده يا معلم! جرب تاني؟ 🔍")
        return STATE_MAIN
    msg = f"🔍 *نتائج البحث ({len(results)} عميل):*\n\n"
    for c in results[:10]:
        icon = "🟢" if c["pay_status"] == "مدفوع" else "🔴"
        msg += (
            f"{icon} *{c['name']}* — {c['package']}\n"
            f"   📞 {c['phone']} | 💰 {c['sell']} ج.م | {c['pay_status']}\n"
            f"   📆 انتهاء: {c['expiry'] or '—'} | 📅 أضيف: {c['date']}\n\n"
        )
    await update.message.reply_text(msg, parse_mode="Markdown")
    return STATE_MAIN

# ══════════════════════════════════════════════
# إلغاء
# ══════════════════════════════════════════════
async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text("❌ تمام يا هندسة، ألغينا العملية 😄\nكلما تحتاج أنا هنا! 🫡")
    return await start(update, ctx)

# ══════════════════════════════════════════════
# تشغيل البوت
# ══════════════════════════════════════════════
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.Regex("^(➕|✅|💸|📥|👥|🔍|📊|❌)"), main_menu),
        ],
        states={
            STATE_MAIN:             [MessageHandler(filters.TEXT & ~filters.COMMAND, main_menu)],
            # إضافة عميل
            STATE_ADD_TYPE:         [CallbackQueryHandler(add_type,         pattern="^type_")],
            STATE_ADD_NAME:         [MessageHandler(filters.TEXT & ~filters.COMMAND, add_name)],
            STATE_ADD_PHONE:        [MessageHandler(filters.TEXT & ~filters.COMMAND, add_phone)],
            STATE_ADD_PKG:          [CallbackQueryHandler(add_pkg,           pattern="^pkg_")],
            STATE_ADD_GIGA:         [MessageHandler(filters.TEXT & ~filters.COMMAND, add_giga)],
            STATE_ADD_NOTES:        [MessageHandler(filters.TEXT | filters.COMMAND,  add_notes)],
            STATE_ADD_CONFIRM:      [CallbackQueryHandler(add_confirm,       pattern="^(confirm|cancel)_add")],
            # دفعة
            STATE_PAY_SEARCH:       [MessageHandler(filters.TEXT & ~filters.COMMAND, pay_search)],
            STATE_PAY_SELECT:       [CallbackQueryHandler(pay_select,        pattern="^pay_")],
            STATE_PAY_CONFIRM:      [CallbackQueryHandler(pay_confirm,       pattern="^(confirm|cancel)_pay")],
            # تحويل للمدير
            STATE_TRANSFER_AMOUNT:  [MessageHandler(filters.TEXT & ~filters.COMMAND, transfer_amount)],
            STATE_TRANSFER_METHOD:  [CallbackQueryHandler(transfer_method,   pattern="^method_")],
            STATE_TRANSFER_CONFIRM: [CallbackQueryHandler(transfer_confirm,  pattern="^(confirm|cancel)_transfer")],
            # استلام رصيد
            STATE_RECEIVE_AMOUNT:   [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_amount)],
            STATE_RECEIVE_NOTE:     [MessageHandler(filters.TEXT | filters.COMMAND,  receive_note)],
            STATE_RECEIVE_CONFIRM:  [CallbackQueryHandler(receive_confirm,   pattern="^(confirm|cancel)_receive")],
            # بحث
            STATE_SEARCH_CLIENT:    [MessageHandler(filters.TEXT & ~filters.COMMAND, search_client)],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CommandHandler("start",  start),
        ],
        allow_reentry=True,
    )

    app.add_handler(conv)
    logger.info("🤖 Bot started!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
