import os
import re
import logging
import httpx
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ['TELEGRAM_TOKEN']
GROQ_API_KEY   = os.environ['GROQ_API_KEY']
GROQ_URL       = "https://api.groq.com/openai/v1/chat/completions"

user_data: dict = {}

SYSTEM_PROMPT = """Ты — профессиональный AI‑ассистент для управления объектами охранно‑пожарной сигнализации (ОПС), систем оповещения и эвакуации (СОУЭ), монтажных и сервисных работ.

Твоя задача:
- фиксировать информацию по объектам
- вести историю работ
- сохранять проблемы и причины остановки работ
- учитывать сотрудников и объемы выполненных работ

Когда пользователь сообщает о выполненных работах, структурируй ответ строго в формате:

✔ Объект: [название]
✔ Сотрудники: [список]
✔ Выполнено:
• [работа — объём]

⚠ Проблема (если есть):
• [описание]

📅 Запись сохранена.

После ответа добавь строку для автообработки:
МЕТА: кабель=[число]м устройства=[число]шт проблема=[да/нет]

Отвечай кратко, профессионально, используй инженерную терминологию ОПС/СОУЭ.
Если данных недостаточно — задавай уточняющие вопросы.
Все ответы только на русском языке."""

def get_user(uid):
    if uid not in user_data:
        user_data[uid] = {"objects": {}, "current": None}
    return user_data[uid]

def get_current_object(uid):
    u = get_user(uid)
    if u["current"] and u["current"] in u["objects"]:
        return u["objects"][u["current"]]
    return None

def fmt_date(ts):
    return datetime.fromtimestamp(ts).strftime("%d.%m.%Y %H:%M")

def status_emoji(s):
    return {"active": "🟢", "paused": "🟡", "problem": "🔴", "done": "✅"}.get(s, "⚪")

def status_label(s):
    return {"active": "В работе", "paused": "Приостановлен", "problem": "Проблема", "done": "Завершён"}.get(s, s)

def objects_keyboard(uid):
    u = get_user(uid)
    buttons = []
    for name, obj in u["objects"].items():
        buttons.append([InlineKeyboardButton(f"{status_emoji(obj['status'])} {name}", callback_data=f"select:{name}")])
    buttons.append([InlineKeyboardButton("➕ Новый объект", callback_data="new_object")])
    return InlineKeyboardMarkup(buttons)

def parse_meta(text):
    meta = {"cable": 0, "devices": 0, "has_problem": False}
    for line in text.split("\n"):
        if line.startswith("МЕТА:"):
            c = re.search(r'кабель=(\d+)', line)
            d = re.search(r'устройства=(\d+)', line)
            p = re.search(r'проблема=(да|нет)', line)
            if c: meta["cable"] = int(c.group(1))
            if d: meta["devices"] = int(d.group(1))
            if p: meta["has_problem"] = p.group(1) == "да"
    return meta

def clean_response(text):
    return "\n".join(l for l in text.split("\n") if not l.startswith("МЕТА:")).strip()

async def ask_groq(uid, user_message):
    u = get_user(uid)
    obj = get_current_object(uid)

    context = ""
    if obj:
        problems_open = [p for p in obj.get("problems", []) if p["status"] != "solved"]
        context = (
            f"Текущий объект: «{obj['name']}»\n"
            f"Адрес: {obj.get('address', 'не указан')}\n"
            f"Заказчик: {obj.get('customer', 'не указан')}\n"
            f"Инженер: {obj.get('engineer', 'не указан')}\n"
            f"Статус: {status_label(obj['status'])}\n"
            f"Кабеля: {obj.get('total_cable', 0)} м | Устройств: {obj.get('total_devices', 0)} шт\n\n"
            f"Последние записи:\n"
            + ("\n".join(f"• {fmt_date(h['ts'])}: {h['text']}" for h in obj.get("history", [])[-5:]) or "нет")
            + "\n\nОткрытые проблемы:\n"
            + ("\n".join(f"• {p['title']}" for p in problems_open) or "нет")
        )

    chat_key = f"chat_{u['current']}" if u["current"] else "chat_general"
    if chat_key not in u:
        u[chat_key] = []

    messages = [{"role": "system", "content": SYSTEM_PROMPT}]
    if context:
        messages.append({"role": "user", "content": context})
        messages.append({"role": "assistant", "content": "Понял контекст объекта. Готов к работе."})
    messages += u[chat_key][-10:]
    messages.append({"role": "user", "content": user_message})

    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            GROQ_URL,
            headers={"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"},
            json={"model": "llama-3.3-70b-versatile", "messages": messages, "max_tokens": 1024, "temperature": 0.3}
        )
        resp.raise_for_status()
        data = resp.json()

    reply = data["choices"][0]["message"]["content"]
    u[chat_key].append({"role": "user", "content": user_message})
    u[chat_key].append({"role": "assistant", "content": reply})

    meta = parse_meta(reply)
    clean = clean_response(reply)

    if obj and ("📅" in reply or "✔" in reply):
        obj.setdefault("history", []).append({"ts": datetime.now().timestamp(), "text": user_message})
        obj["total_cable"]   = obj.get("total_cable", 0)   + meta["cable"]
        obj["total_devices"] = obj.get("total_devices", 0) + meta["devices"]
        if meta["has_problem"]:
            obj.setdefault("problems", []).append({
                "id": str(len(obj.get("problems", [])) + 1),
                "title": f"Проблема от {fmt_date(datetime.now().timestamp())}",
                "status": "open", "ts": datetime.now().timestamp(),
            })

    return clean, meta

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    u = get_user(uid)
    name = update.effective_user.first_name or "коллега"
    text = (
        f"👋 Привет, {name}!\n\nЯ — AI-диспетчер объектов ОПС.\n\n"
        "Помогу:\n• фиксировать ежедневные работы\n• хранить историю\n"
        "• отслеживать проблемы\n• формировать отчёты\n\n"
    )
    if u["objects"]:
        await update.message.reply_text(text + "Выберите объект:", reply_markup=objects_keyboard(uid))
    else:
        await update.message.reply_text(text + "Создайте первый объект: /new")

async def cmd_new(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📋 Введите название нового объекта:")
    ctx.user_data["awaiting"] = "object_name"

async def cmd_objects(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    u = get_user(uid)
    if not u["objects"]:
        await update.message.reply_text("Объектов нет. Создайте первый: /new")
        return
    await update.message.reply_text("📂 Ваши объекты:", reply_markup=objects_keyboard(uid))

async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    obj = get_current_object(uid)
    if not obj:
        await update.message.reply_text("Объект не выбран. Используйте /objects")
        return
    probs_open = [p for p in obj.get("problems", []) if p["status"] != "solved"]
    last = obj.get("history", [])
    text = (
        f"{status_emoji(obj['status'])} *{obj['name']}*\n"
        f"📍 {obj.get('address','не указан')}\n👤 {obj.get('engineer','не указан')}\n"
        f"🏢 {obj.get('customer','не указан')}\n\n"
        f"Статус: *{status_label(obj['status'])}*\n"
        f"🔌 Кабель: {obj.get('total_cable',0)} м\n📡 Устройства: {obj.get('total_devices',0)} шт\n"
        f"⚠️ Открытых проблем: {len(probs_open)}"
        + (f"\n📝 Последняя запись: {fmt_date(last[-1]['ts'])}" if last else "")
    )
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🟢 В работе", callback_data="status:active"),
         InlineKeyboardButton("🟡 Пауза",    callback_data="status:paused")],
        [InlineKeyboardButton("🔴 Проблема", callback_data="status:problem"),
         InlineKeyboardButton("✅ Завершён",  callback_data="status:done")],
        [InlineKeyboardButton("📋 История",  callback_data="history")],
        [InlineKeyboardButton("⚠️ Проблемы", callback_data="problems")],
    ])
    await update.message.reply_text(text, parse_mode="Markdown", reply_markup=keyboard)

async def cmd_history(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message or update.callback_query.message
    uid = update.effective_user.id
    obj = get_current_object(uid)
    if not obj:
        await msg.reply_text("Объект не выбран. Используйте /objects")
        return
    history = obj.get("history", [])
    if not history:
        await msg.reply_text(f"📋 История по «{obj['name']}» пуста.")
        return
    lines = [f"📋 *История: {obj['name']}*\n"]
    for h in history[-10:]:
        lines.append(f"📅 {fmt_date(h['ts'])}\n{h['text']}\n")
    await msg.reply_text("\n".join(lines), parse_mode="Markdown")

async def cmd_problems(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message or update.callback_query.message
    uid = update.effective_user.id
    obj = get_current_object(uid)
    if not obj:
        await msg.reply_text("Объект не выбран. Используйте /objects")
        return
    problems = obj.get("problems", [])
    if not problems:
        await msg.reply_text(f"✅ Проблем по «{obj['name']}» нет.")
        return
    icons = {"open": "🔴", "inwork": "🟡", "solved": "🟢"}
    names = {"open": "Открыта", "inwork": "В работе", "solved": "Решена"}
    lines = [f"⚠️ *Проблемы: {obj['name']}*\n"]
    for p in problems:
        lines.append(f"{icons.get(p['status'],'⚪')} {p['title']}\n   {names.get(p['status'],p['status'])} · {fmt_date(p['ts'])}\n")
    await msg.reply_text("\n".join(lines), parse_mode="Markdown")

async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    obj = get_current_object(uid)
    if not obj:
        await update.message.reply_text("Объект не выбран. Используйте /objects")
        return
    await update.message.reply_text("⏳ Формирую отчёт...")
    try:
        reply, _ = await ask_groq(uid, "Сформируй краткий отчёт: что сделано, проблемы, текущий статус.")
        await update.message.reply_text(reply)
    except Exception as e:
        logger.error(f"Groq error: {e}")
        await update.message.reply_text("⚠️ Ошибка AI. Попробуйте позже.")

async def cmd_help(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = (
        "📖 *Команды:*\n\n/start — начало\n/new — создать объект\n/objects — список объектов\n"
        "/status — статус объекта\n/history — история работ\n/problems — проблемы\n/report — отчёт\n\n"
        "💬 *Просто пишите:*\n_«Иван протянул 300 м кабеля на 3 этаже, остановились — нет лотков»_"
    )
    await update.message.reply_text(text, parse_mode="Markdown")

async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    text = update.message.text.strip()
    awaiting = ctx.user_data.get("awaiting")

    if awaiting == "object_name":
        u = get_user(uid)
        u["objects"][text] = {
            "name": text, "address": "", "customer": "", "engineer": "",
            "status": "active", "history": [], "problems": [],
            "total_cable": 0, "total_devices": 0, "created": datetime.now().timestamp(),
        }
        u["current"] = text
        ctx.user_data["awaiting"] = "object_address"
        await update.message.reply_text(f"✅ Объект «{text}» создан!\n\nВведите адрес (или /skip):")
        return

    if awaiting == "object_address":
        obj = get_current_object(uid)
        if obj and text != "/skip": obj["address"] = text
        ctx.user_data["awaiting"] = "object_customer"
        await update.message.reply_text("Введите заказчика (или /skip):")
        return

    if awaiting == "object_customer":
        obj = get_current_object(uid)
        if obj and text != "/skip": obj["customer"] = text
        ctx.user_data["awaiting"] = "object_engineer"
        await update.message.reply_text("Введите ответственного инженера (или /skip):")
        return

    if awaiting == "object_engineer":
        obj = get_current_object(uid)
        if obj and text != "/skip": obj["engineer"] = text
        ctx.user_data.pop("awaiting", None)
        await update.message.reply_text(
            "✅ Готово! Теперь просто пишите что было сделано.\n\n"
            "_Например: «Иван протянул 200 м кабеля на 3 этаже»_", parse_mode="Markdown"
        )
        return

    if not get_current_object(uid):
        u = get_user(uid)
        if u["objects"]:
            await update.message.reply_text("Выберите объект:", reply_markup=objects_keyboard(uid))
        else:
            await update.message.reply_text("Объектов нет. Создайте первый: /new")
        return

    await update.message.chat.send_action("typing")
    try:
        reply, _ = await ask_groq(uid, text)
        await update.message.reply_text(reply)
    except Exception as e:
        logger.error(f"Groq error: {e}")
        await update.message.reply_text("⚠️ Ошибка AI. Проверьте GROQ_API_KEY или попробуйте позже.")

async def handle_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    uid = query.from_user.id
    data = query.data

    if data == "new_object":
        await query.message.reply_text("Введите название нового объекта:")
        ctx.user_data["awaiting"] = "object_name"
    elif data.startswith("select:"):
        name = data[7:]
        u = get_user(uid)
        if name in u["objects"]:
            u["current"] = name
            obj = u["objects"][name]
            await query.message.reply_text(
                f"✅ Выбран: *{name}*\nСтатус: {status_emoji(obj['status'])} {status_label(obj['status'])}\n\nПишите что было сделано.",
                parse_mode="Markdown"
            )
    elif data.startswith("status:"):
        new_status = data[7:]
        obj = get_current_object(uid)
        if obj:
            obj["status"] = new_status
            await query.message.reply_text(f"✅ Статус: {status_emoji(new_status)} {status_label(new_status)}")
    elif data == "history":
        await cmd_history(update, ctx)
    elif data == "problems":
        await cmd_problems(update, ctx)

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start",    cmd_start))
    app.add_handler(CommandHandler("new",      cmd_new))
    app.add_handler(CommandHandler("objects",  cmd_objects))
    app.add_handler(CommandHandler("status",   cmd_status))
    app.add_handler(CommandHandler("history",  cmd_history))
    app.add_handler(CommandHandler("problems", cmd_problems))
    app.add_handler(CommandHandler("report",   cmd_report))
    app.add_handler(CommandHandler("help",     cmd_help))
    app.add_handler(CommandHandler("skip",     handle_message))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    logger.info("Бот запущен на Groq (бесплатно)...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
