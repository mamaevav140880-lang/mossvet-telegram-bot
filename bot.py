
import os, json, uuid, re
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, ContextTypes, filters
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from openpyxl import Workbook
from pypdf import PdfReader

load_dotenv()

BASE = Path(__file__).resolve().parent
DATA_ROOT = Path(os.getenv("DATA_ROOT", "/data" if os.getenv("RAILWAY_ENVIRONMENT") else str(BASE)))
DATA = DATA_ROOT / "data"
UPLOADS = DATA_ROOT / "uploads"
EXPORTS = DATA_ROOT / "exports"
for d in (DATA, UPLOADS, EXPORTS):
    d.mkdir(parents=True, exist_ok=True)

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "").strip()
SERVICE_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

CATALOG = json.loads((BASE/"catalog.json").read_text(encoding="utf-8"))

FIXED = {
    "organization": 'ГУП «Моссвет»',
    "representative": 'Ведущий инженер технического надзора Мамаев А.В.',
    "representative_short": 'Мамаев А.В.',
}

REQUIRED = ["object_name", "address", "inspection_date"]

def user_path(uid):
    return DATA / f"{uid}.json"

def empty_act():
    return {
        "id": uuid.uuid4().hex,
        "created": datetime.now().isoformat(timespec="seconds"),
        "act_number": "",
        "inspection_date": "",
        "object_name": "",
        "address": "",
        "customer": "",
        "basis": "",
        "commission_extra": "",
        "remarks": "",
        "equipment": [],
        "photos": [],
        "pdf_files": [],
        "pdf_text": ""
    }

def load_act(uid):
    p = user_path(uid)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    a = empty_act()
    save_act(uid, a)
    return a

def save_act(uid, act):
    user_path(uid).write_text(json.dumps(act, ensure_ascii=False, indent=2), encoding="utf-8")

def menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🆕 Новый акт", callback_data="new"),
         InlineKeyboardButton("🧭 Заполнить по шагам", callback_data="wizard")],
        [InlineKeyboardButton("💡 Оборудование", callback_data="equipment"),
         InlineKeyboardButton("📷 Фото / PDF", callback_data="uploads")],
        [InlineKeyboardButton("⚠️ Замечания", callback_data="remarks"),
         InlineKeyboardButton("✅ Проверить", callback_data="validate")],
        [InlineKeyboardButton("👁 Черновик", callback_data="preview")],
        [InlineKeyboardButton("📄 PDF", callback_data="make_pdf"),
         InlineKeyboardButton("📊 Excel", callback_data="make_excel")],
        [InlineKeyboardButton("☁️ Google Sheets", callback_data="sheet")]
    ])

def kb(options, prefix):
    rows = [[InlineKeyboardButton(x, callback_data=f"{prefix}:{i}")] for i, x in enumerate(options)]
    rows.append([InlineKeyboardButton("↩️ Меню", callback_data="menu")])
    return InlineKeyboardMarkup(rows)

def missing_fields(act):
    labels = {
        "object_name": "Наименование объекта",
        "address": "Адрес",
        "inspection_date": "Дата обследования",
    }
    return [labels[k] for k in REQUIRED if not str(act.get(k, "")).strip()]

def preview(act):
    eq = act.get("equipment") or []
    eq_text = "\n".join(
        f"{i+1}) №{x.get('number','—')} | {x.get('support','—')} | {x.get('luminaire','—')} | "
        f"{x.get('bracket','—')} | {x.get('base','—')}"
        for i, x in enumerate(eq)
    ) or "Не добавлено"
    return (
        f"АКТ ОБСЛЕДОВАНИЯ\n\n"
        f"№: {act.get('act_number') or '—'}\n"
        f"Дата: {act.get('inspection_date') or '—'}\n"
        f"Объект: {act.get('object_name') or '—'}\n"
        f"Адрес: {act.get('address') or '—'}\n"
        f"Заказчик: {act.get('customer') or '—'}\n"
        f"Основание: {act.get('basis') or '—'}\n\n"
        f"{FIXED['organization']}\n{FIXED['representative']}\n"
        f"Доп. комиссия: {act.get('commission_extra') or '—'}\n\n"
        f"Оборудование:\n{eq_text}\n\n"
        f"Замечания: {act.get('remarks') or '—'}\n"
        f"Фото: {len(act.get('photos',[]))}; PDF: {len(act.get('pdf_files',[]))}"
    )

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    act = load_act(update.effective_user.id)
    await update.message.reply_text(
        "Бот оформления акта обследования наружного освещения.\n\n"
        f"Постоянно закреплено:\n{FIXED['organization']}\n{FIXED['representative']}",
        reply_markup=menu()
    )

async def ask_field(message, context, field):
    prompts = {
        "object_name": "Введите наименование объекта:",
        "address": "Введите адрес объекта:",
        "act_number": "Введите номер акта (или '-' если пока нет):",
        "inspection_date": "Введите дату обследования, например 15.09.2026:",
        "customer": "Введите заказчика (или '-'):",
        "basis": "Введите основание/программу работ (или '-'):",
        "commission_extra": "Введите дополнительных представителей комиссии (или '-'):"
    }
    context.user_data["mode"] = f"field:{field}"
    await message.reply_text(prompts[field])

async def wizard_next(message, context, act):
    seq = ["object_name","address","act_number","inspection_date","customer","basis","commission_extra"]
    for f in seq:
        if not str(act.get(f,"")).strip():
            await ask_field(message, context, f)
            return
    context.user_data.pop("mode", None)
    await message.reply_text("Основные сведения заполнены.", reply_markup=menu())

async def on_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    act = load_act(uid)
    a = q.data

    if a == "menu":
        context.user_data.pop("mode", None)
        await q.message.reply_text("Главное меню:", reply_markup=menu())
        return
    if a == "new":
        act = empty_act()
        save_act(uid, act)
        context.user_data.clear()
        await q.message.reply_text("Создан новый акт.", reply_markup=menu())
        return
    if a == "wizard":
        context.user_data["wizard"] = True
        await wizard_next(q.message, context, act)
        return
    if a == "equipment":
        context.user_data["mode"] = "equip_number"
        context.user_data["equip"] = {}
        await q.message.reply_text("Введите номер опоры/позиции:")
        return
    if a == "uploads":
        context.user_data["mode"] = "uploads"
        await q.message.reply_text("Отправьте фото или PDF. Можно несколько файлов. Для возврата нажмите /start.")
        return
    if a == "remarks":
        context.user_data["mode"] = "remarks"
        await q.message.reply_text("Введите замечания. Если замечаний нет — напишите «нет».")
        return
    if a == "validate":
        miss = missing_fields(act)
        if miss:
            await q.message.reply_text("Не заполнены обязательные поля:\n• " + "\n• ".join(miss))
        else:
            await q.message.reply_text("✅ Обязательные поля заполнены. Акт готов к выпуску.")
        return
    if a == "preview":
        await q.message.reply_text(preview(act))
        return
    if a == "make_pdf":
        miss = missing_fields(act)
        if miss:
            await q.message.reply_text("Сначала заполните:\n• " + "\n• ".join(miss))
            return
        path = build_pdf(act)
        with open(path, "rb") as f:
            await q.message.reply_document(f, filename=path.name)
        return
    if a == "make_excel":
        path = build_excel(act)
        with open(path, "rb") as f:
            await q.message.reply_document(f, filename=path.name)
        return
    if a == "sheet":
        ok, msg = append_sheet(act)
        await q.message.reply_text(msg)
        return

    if a.startswith("support:"):
        opt = CATALOG["support_types"][int(a.split(":")[1])]
        context.user_data["equip"]["support"] = opt
        context.user_data["mode"] = "equip_luminaire"
        await q.message.reply_text("Выберите тип/производителя светильника:", reply_markup=kb(CATALOG["luminaire_types"], "luminaire"))
        return
    if a.startswith("luminaire:"):
        opt = CATALOG["luminaire_types"][int(a.split(":")[1])]
        context.user_data["equip"]["luminaire"] = opt
        context.user_data["mode"] = "equip_bracket"
        await q.message.reply_text("Выберите кронштейн:", reply_markup=kb(CATALOG["bracket_types"], "bracket"))
        return
    if a.startswith("bracket:"):
        opt = CATALOG["bracket_types"][int(a.split(":")[1])]
        context.user_data["equip"]["bracket"] = opt
        context.user_data["mode"] = "equip_base"
        await q.message.reply_text("Выберите цоколь:", reply_markup=kb(CATALOG["base_types"], "base"))
        return
    if a.startswith("base:"):
        opt = CATALOG["base_types"][int(a.split(":")[1])]
        context.user_data["equip"]["base"] = opt
        act["equipment"].append(context.user_data["equip"])
        save_act(uid, act)
        context.user_data.pop("equip", None)
        context.user_data.pop("mode", None)
        await q.message.reply_text("✅ Оборудование добавлено.", reply_markup=menu())
        return

async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    act = load_act(uid)
    mode = context.user_data.get("mode", "")
    text = update.message.text.strip()

    if mode.startswith("field:"):
        field = mode.split(":",1)[1]
        act[field] = "" if text == "-" else text
        save_act(uid, act)
        context.user_data.pop("mode", None)
        if context.user_data.get("wizard"):
            await wizard_next(update.message, context, act)
        else:
            await update.message.reply_text("Сохранено.", reply_markup=menu())
        return

    if mode == "equip_number":
        context.user_data["equip"]["number"] = text
        context.user_data["mode"] = "equip_support"
        await update.message.reply_text("Выберите тип опоры:", reply_markup=kb(CATALOG["support_types"], "support"))
        return

    if mode == "remarks":
        act["remarks"] = "Замечаний нет." if text.lower() == "нет" else text
        save_act(uid, act)
        context.user_data.pop("mode", None)
        await update.message.reply_text("Замечания сохранены.", reply_markup=menu())
        return

    await update.message.reply_text("Используйте /start для открытия меню.")

async def on_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    act = load_act(uid)
    photo = update.message.photo[-1]
    f = await photo.get_file()
    path = UPLOADS / f"{uid}_{uuid.uuid4().hex}.jpg"
    await f.download_to_drive(custom_path=str(path))
    act["photos"].append(str(path))
    save_act(uid, act)
    await update.message.reply_text(f"Фото сохранено. Всего: {len(act['photos'])}")

async def on_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    act = load_act(uid)
    doc = update.message.document
    name = doc.file_name or "document"
    suffix = Path(name).suffix.lower()
    f = await doc.get_file()
    path = UPLOADS / f"{uid}_{uuid.uuid4().hex}{suffix}"
    await f.download_to_drive(custom_path=str(path))

    if suffix == ".pdf":
        act["pdf_files"].append(str(path))
        extracted = extract_pdf_text(path)
        if extracted:
            act["pdf_text"] = (act.get("pdf_text","") + "\n" + extracted)[:30000]
            hints = guess_fields(extracted)
            changed = []
            for k,v in hints.items():
                if v and not act.get(k):
                    act[k] = v
                    changed.append(k)
            save_act(uid, act)
            txt = "PDF сохранён. Текст извлечён."
            if changed:
                txt += "\nАвтоматически предварительно заполнено: " + ", ".join(changed)
            await update.message.reply_text(txt)
        else:
            save_act(uid, act)
            await update.message.reply_text("PDF сохранён. Текстовый слой не найден — вероятно, это скан.")
    else:
        await update.message.reply_text("Файл сохранён, но автоматический разбор сейчас поддерживается для PDF.")

def extract_pdf_text(path):
    try:
        reader = PdfReader(str(path))
        parts = []
        for p in reader.pages[:30]:
            parts.append(p.extract_text() or "")
        return "\n".join(parts).strip()
    except Exception:
        return ""

def guess_fields(text):
    out = {}
    m = re.search(r"(?:адрес|по адресу)\s*[:\-]?\s*([^\n]{5,120})", text, re.I)
    if m: out["address"] = m.group(1).strip()
    m = re.search(r"(?:акт\s*№|№\s*акта)\s*([A-Za-zА-Яа-я0-9\-\/]+)", text, re.I)
    if m: out["act_number"] = m.group(1).strip()
    m = re.search(r"\b(\d{2}\.\d{2}\.\d{4})\b", text)
    if m: out["inspection_date"] = m.group(1)
    return out

def font_name():
    for p in ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf","C:/Windows/Fonts/arial.ttf"]:
        if Path(p).exists():
            pdfmetrics.registerFont(TTFont("RU", p))
            return "RU"
    return "Helvetica"

def build_pdf(act):
    fnt = font_name()
    path = EXPORTS / f"akt_{act.get('act_number') or act['id'][:8]}.pdf"
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    normal = ParagraphStyle("n", parent=styles["Normal"], fontName=fnt, fontSize=9, leading=12)
    title = ParagraphStyle("t", parent=styles["Title"], fontName=fnt, fontSize=14, leading=17)
    story = [
        Paragraph("АКТ ОБСЛЕДОВАНИЯ ОБЪЕКТА НАРУЖНОГО ОСВЕЩЕНИЯ", title),
        Spacer(1,8),
        Paragraph(FIXED["organization"], normal),
        Paragraph(FIXED["representative"], normal),
        Spacer(1,10),
    ]
    info = [
        ["№ акта", act.get("act_number","")],
        ["Дата", act.get("inspection_date","")],
        ["Объект", act.get("object_name","")],
        ["Адрес", act.get("address","")],
        ["Заказчик", act.get("customer","")],
        ["Основание", act.get("basis","")],
    ]
    t = Table(info, colWidths=[120,380])
    t.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),fnt),("FONTSIZE",(0,0),(-1,-1),9),
        ("GRID",(0,0),(-1,-1),0.5,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP"),
        ("BACKGROUND",(0,0),(0,-1),colors.whitesmoke),("PADDING",(0,0),(-1,-1),5)
    ]))
    story += [t, Spacer(1,12), Paragraph("Оборудование", normal)]
    rows = [["№","Тип опоры","Светильник","Кронштейн","Цоколь"]]
    for x in act.get("equipment",[]):
        rows.append([x.get("number",""),x.get("support",""),x.get("luminaire",""),x.get("bracket",""),x.get("base","")])
    if len(rows)==1: rows.append(["","","","",""])
    et = Table(rows, repeatRows=1, colWidths=[45,105,125,115,90])
    et.setStyle(TableStyle([
        ("FONTNAME",(0,0),(-1,-1),fnt),("FONTSIZE",(0,0),(-1,-1),8),
        ("GRID",(0,0),(-1,-1),0.5,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP"),
        ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke)
    ]))
    story += [et, Spacer(1,12), Paragraph("Замечания", normal),
              Paragraph(act.get("remarks") or "Не указано.", normal), Spacer(1,12)]
    for p in act.get("photos",[])[:8]:
        try:
            story += [RLImage(p, width=220, height=165), Spacer(1,5)]
        except Exception:
            pass
    story += [
        Spacer(1,12),
        Paragraph("Подпись эксплуатирующей организации: ____________________", normal),
        Spacer(1,10),
        Paragraph("Представитель ГУП «Моссвет» Мамаев А.В.: ____________________", normal)
    ]
    doc.build(story)
    return path

def build_excel(act):
    path = EXPORTS / f"akt_{act.get('act_number') or act['id'][:8]}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Акт"
    for row in [
        ["Поле","Значение"],
        ["Номер акта",act.get("act_number","")],
        ["Дата",act.get("inspection_date","")],
        ["Объект",act.get("object_name","")],
        ["Адрес",act.get("address","")],
        ["Заказчик",act.get("customer","")],
        ["Основание",act.get("basis","")],
        ["Организация",FIXED["organization"]],
        ["Представитель",FIXED["representative"]],
        ["Комиссия",act.get("commission_extra","")],
        ["Замечания",act.get("remarks","")],
        ["Фото",len(act.get("photos",[]))],
        ["PDF",len(act.get("pdf_files",[]))],
    ]:
        ws.append(row)
    eq = wb.create_sheet("Оборудование")
    eq.append(["№","Тип опоры","Светильник","Кронштейн","Цоколь"])
    for x in act.get("equipment",[]):
        eq.append([x.get("number",""),x.get("support",""),x.get("luminaire",""),x.get("bracket",""),x.get("base","")])
    wb.save(path)
    return path

def append_sheet(act):
    if not SHEET_ID or not SERVICE_JSON:
        return False, "Google Sheets не настроен. Заполните GOOGLE_SHEET_ID и GOOGLE_SERVICE_ACCOUNT_JSON в .env."
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file(SERVICE_JSON, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SHEET_ID)
        ws = sh.sheet1
        if not ws.row_values(1):
            ws.append_row(["Дата записи","№ акта","Дата обследования","Объект","Адрес","Заказчик","Основание","Замечания","Кол-во оборудования","Кол-во фото"])
        ws.append_row([
            datetime.now().strftime("%d.%m.%Y %H:%M"),
            act.get("act_number",""), act.get("inspection_date",""), act.get("object_name",""),
            act.get("address",""), act.get("customer",""), act.get("basis",""),
            act.get("remarks",""), len(act.get("equipment",[])), len(act.get("photos",[]))
        ])
        return True, "✅ Акт записан в Google Sheets."
    except Exception as e:
        return False, f"Ошибка Google Sheets: {e}"

def main():
    if not TOKEN:
        raise RuntimeError("В .env не задан TELEGRAM_BOT_TOKEN")
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(on_button))
    app.add_handler(MessageHandler(filters.PHOTO, on_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    print("Mossvet bot v2 started")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
