import os
import re
import logging
import httpx
import asyncio
from datetime import datetime, timedelta
from io import BytesIO
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
import psycopg2
from psycopg2.extras import RealDictCursor
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ['TELEGRAM_TOKEN']
GROQ_API_KEY   = os.environ['GROQ_API_KEY']
DATABASE_URL   = os.environ['DATABASE_URL']
ADMIN_IDS      = [int(x) for x in os.environ.get('ADMIN_IDS', '').split(',') if x.strip()]
GROQ_URL       = "https://api.groq.com/openai/v1/chat/completions"

SYSTEM_PROMPT = """Ты — профессиональный AI‑ассистент для управления объектами охранно‑пожарной сигнализации (ОПС), систем оповещения и эвакуации (СОУЭ), монтажных и сервисных работ.

Когда пользователь сообщает о выполненных работах, структурируй ответ строго в формате:

✔ Объект: [название]
✔ Сотрудники: [список]
✔ Выполнено:
• [работа — объём]

⚠ Проблема (если есть):
• [описание]

📅 Запись сохранена.

После ответа добавь строку:
МЕТА: кабель=[число]м устройства=[число]шт проблема=[да/нет]

Отвечай кратко, профессионально, используй инженерную терминологию ОПС/СОУЭ.
Если данных недостаточно — задавай уточняющие вопросы.
Все ответы только на русском языке."""

# ─── Database ────────────────────────────────────────────────────────────────

def get_db():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)

def init_db():
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    uid BIGINT PRIMARY KEY,
                    name TEXT,
                    role TEXT DEFAULT 'worker',
                    created_at TIMESTAMP DEFAULT NOW()
                );
                CREATE TABLE IF NOT EXISTS objects (
                    id SERIAL PRIMARY KEY,
                    name TEXT NOT NULL,
                    address TEXT DEFAULT '',
                    customer TEXT DEFAULT '',
                    engineer TEXT DEFAULT '',
                    status TEXT DEFAULT 'active',
                    total_cable INT DEFAULT 0,
                    total_devices INT DEFAULT 0,
                    notes TEXT DEFAULT '',
                    owner_uid BIGINT,
                    created_at TIMESTAMP DEFAULT NOW()
                );
                CREATE TABLE IF NOT EXISTS history (
                    id SERIAL PRIMARY KEY,
                    object_id INT REFERENCES objects(id) ON DELETE CASCADE,
                    uid BIGINT,
                    text TEXT,
                    created_at TIMESTAMP DEFAULT NOW()
                );
                CREATE TABLE IF NOT EXISTS problems (
                    id SERIAL PRIMARY KEY,
                    object_id INT REFERENCES objects(id) ON DELETE CASCADE,
                    title TEXT,
                    status TEXT DEFAULT 'open',
                    created_at TIMESTAMP DEFAULT NOW()
                );
                CREATE TABLE IF NOT EXISTS photos (
                    id SERIAL PRIMARY KEY,
                    object_id INT REFERENCES objects(id) ON DELETE CASCADE,
                    uid BIGINT,
                    file_id TEXT,
                    caption TEXT DEFAULT '',
                    created_at TIMESTAMP DEFAULT NOW()
                );
                CREATE TABLE IF NOT EXISTS user_objects (
                    uid BIGINT,
                    object_id INT REFERENCES objects(id) ON DELETE CASCADE,
                    current_object_id INT,
                    PRIMARY KEY (uid)
                );
                CREATE TABLE IF NOT EXISTS chat_history (
                    id SERIAL PRIMARY KEY,
                    uid BIGINT,
                    object_id INT,
                    role TEXT,
                    content TEXT,
                    created_at TIMESTAMP DEFAULT NOW()
                );
            """)
            conn.commit()
    logger.info("БД инициализирована")

# ─── DB helpers ──────────────────────────────────────────────────────────────

def db_get_user(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE uid=%s", (uid,))
            return cur.fetchone()

def db_upsert_user(uid, name):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO users (uid, name) VALUES (%s, %s)
                ON CONFLICT (uid) DO UPDATE SET name=EXCLUDED.name
            """, (uid, name))
            conn.commit()

def db_set_role(uid, role):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE users SET role=%s WHERE uid=%s", (role, uid))
            conn.commit()

def db_is_admin(uid):
    if uid in ADMIN_IDS:
        return True
    user = db_get_user(uid)
    return user and user['role'] == 'admin'

def db_get_objects(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            if db_is_admin(uid):
                cur.execute("SELECT * FROM objects ORDER BY created_at DESC")
            else:
                cur.execute("SELECT * FROM objects WHERE owner_uid=%s ORDER BY created_at DESC", (uid,))
            return cur.fetchall()

def db_get_object(obj_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM objects WHERE id=%s", (obj_id,))
            return cur.fetchone()

def db_create_object(uid, name):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO objects (name, owner_uid) VALUES (%s, %s) RETURNING id
            """, (name, uid))
            obj_id = cur.fetchone()['id']
            cur.execute("""
                INSERT INTO user_objects (uid, current_object_id) VALUES (%s, %s)
                ON CONFLICT (uid) DO UPDATE SET current_object_id=EXCLUDED.current_object_id
            """, (uid, obj_id))
            conn.commit()
            return obj_id

def db_update_object(obj_id, **kwargs):
    if not kwargs:
        return
    fields = ", ".join(f"{k}=%s" for k in kwargs)
    vals = list(kwargs.values()) + [obj_id]
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute(f"UPDATE objects SET {fields} WHERE id=%s", vals)
            conn.commit()

def db_get_current_object(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT current_object_id FROM user_objects WHERE uid=%s", (uid,))
            row = cur.fetchone()
            if not row or not row['current_object_id']:
                return None
            return db_get_object(row['current_object_id'])

def db_set_current_object(uid, obj_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO user_objects (uid, current_object_id) VALUES (%s, %s)
                ON CONFLICT (uid) DO UPDATE SET current_object_id=EXCLUDED.current_object_id
            """, (uid, obj_id))
            conn.commit()

def db_add_history(obj_id, uid, text):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO history (object_id, uid, text) VALUES (%s, %s, %s)", (obj_id, uid, text))
            conn.commit()

def db_get_history(obj_id, limit=10):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM history WHERE object_id=%s ORDER BY created_at DESC LIMIT %s", (obj_id, limit))
            return list(reversed(cur.fetchall()))

def db_add_problem(obj_id, title):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO problems (object_id, title) VALUES (%s, %s)", (obj_id, title))
            conn.commit()

def db_get_problems(obj_id, status=None):
    with get_db() as conn:
        with conn.cursor() as cur:
            if status:
                cur.execute("SELECT * FROM problems WHERE object_id=%s AND status=%s ORDER BY created_at DESC", (obj_id, status))
            else:
                cur.execute("SELECT * FROM problems WHERE object_id=%s ORDER BY created_at DESC", (obj_id,))
            return cur.fetchall()

def db_solve_problem(prob_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE problems SET status='solved' WHERE id=%s", (prob_id,))
            conn.commit()

def db_add_photo(obj_id, uid, file_id, caption):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO photos (object_id, uid, file_id, caption) VALUES (%s,%s,%s,%s)", (obj_id, uid, file_id, caption))
            conn.commit()

def db_get_photos(obj_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM photos WHERE object_id=%s ORDER BY created_at DESC", (obj_id,))
            return cur.fetchall()

def db_get_chat_history(uid, obj_id, limit=10):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT role, content FROM chat_history
                WHERE uid=%s AND object_id=%s
                ORDER BY created_at DESC LIMIT %s
            """, (uid, obj_id, limit))
            return list(reversed(cur.fetchall()))

def db_add_chat_message(uid, obj_id, role, content):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO chat_history (uid, object_id, role, content) VALUES (%s,%s,%s,%s)", (uid, obj_id, role, content))
            conn.commit()

def db_get_all_objects_with_open_problems():
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT o.*, COUNT(p.id) as open_problems
                FROM objects o
                JOIN problems p ON p.object_id=o.id AND p.status='open'
                GROUP BY o.id
                HAVING COUNT(p.id) > 0
            """)
            return cur.fetchall()

# ─── Helpers ─────────────────────────────────────────────────────────────────

def fmt_date(dt):
    if isinstance(dt, datetime):
        return dt.strftime("%d.%m.%Y %H:%M")
    return str(dt)

def status_emoji(s):
    return {"active": "🟢", "paused": "🟡", "problem": "🔴", "done": "✅"}.get(s, "⚪")

def status_label(s):
    return {"active": "В работе", "paused": "Приостановлен", "problem": "Проблема", "done": "Завершён"}.get(s, s)

def objects_keyboard(uid):
    objects = db_get_objects(uid)
    buttons = []
    for o in objects:
        buttons.append([InlineKeyboardButton(
            f"{status_emoji(o['status'])} {o['name']}", callback_data=f"select:{o['id']}"
        )])
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

# ─── Groq AI ─────────────────────────────────────────────────────────────────

async def ask_groq(uid, obj, user_message):
    problems_open = db_get_problems(obj['id'], status='open')
    history = db_get_history(obj['id'], limit=5)

    context = (
        f"Текущий объект: «{obj['name']}»\n"
        f"Адрес: {obj.get('address') or 'не указан'}\n"
        f"Заказчик: {obj.get('customer') or 'не указан'}\n"
        f"Инженер: {obj.get('engineer') or 'не указан'}\n"
        f"Статус: {status_label(obj['status'])}\n"
        f"Кабеля: {obj['total_cable']} м | Устройств: {obj['total_devices']} шт\n\n"
        f"Последние записи:\n"
        + ("\n".join(f"• {fmt_date(h['created_at'])}: {h['text']}" for h in history) or "нет")
        + "\n\nОткрытые проблемы:\n"
        + ("\n".join(f"• {p['title']}" for p in problems_open) or "нет")
    )

    chat_hist = db_get_chat_history(uid, obj['id'], limit=10)
    messages = [{"role": "system", "content": SYSTEM_PROMPT}]
    messages.append({"role": "user", "content": context})
    messages.append({"role": "assistant", "content": "Понял контекст объекта. Готов к работе."})
    for m in chat_hist:
        messages.append({"role": m['role'], "content": m['content']})
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
    db_add_chat_message(uid, obj['id'], "user", user_message)
    db_add_chat_message(uid, obj['id'], "assistant", reply)

    meta = parse_meta(reply)
    clean = clean_response(reply)

    if "📅" in reply or "✔" in reply:
        db_add_history(obj['id'], uid, user_message)
        db_update_object(obj['id'],
            total_cable=obj['total_cable'] + meta['cable'],
            total_devices=obj['total_devices'] + meta['devices']
        )
        if meta['has_problem']:
            db_add_problem(obj['id'], f"Проблема от {fmt_date(datetime.now())}")

    return clean, meta

# ─── PDF Export ───────────────────────────────────────────────────────────────

def generate_pdf(obj):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle('title', parent=styles['Title'], fontSize=16, spaceAfter=6)
    h2_style    = ParagraphStyle('h2',    parent=styles['Heading2'], fontSize=12, spaceAfter=4)
    normal      = ParagraphStyle('n',     parent=styles['Normal'], fontSize=10, spaceAfter=3)

    story.append(Paragraph(f"Отчёт по объекту: {obj['name']}", title_style))
    story.append(Paragraph(f"Сформирован: {fmt_date(datetime.now())}", normal))
    story.append(Spacer(1, 12))

    info = [
        ["Адрес", obj.get('address') or '—'],
        ["Заказчик", obj.get('customer') or '—'],
        ["Инженер", obj.get('engineer') or '—'],
        ["Статус", status_label(obj['status'])],
        ["Кабель протянут", f"{obj['total_cable']} м"],
        ["Устройств установлено", f"{obj['total_devices']} шт"],
    ]
    t = Table(info, colWidths=[150, 330])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#f0f4ff')),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t)
    story.append(Spacer(1, 16))

    history = db_get_history(obj['id'], limit=50)
    story.append(Paragraph("История работ", h2_style))
    if history:
        for h in history:
            story.append(Paragraph(f"<b>{fmt_date(h['created_at'])}</b> — {h['text']}", normal))
    else:
        story.append(Paragraph("Записей нет", normal))
    story.append(Spacer(1, 12))

    problems = db_get_problems(obj['id'])
    story.append(Paragraph("Проблемы", h2_style))
    icons = {"open": "🔴", "inwork": "🟡", "solved": "🟢"}
    names = {"open": "Открыта", "inwork": "В работе", "solved": "Решена"}
    if problems:
        for p in problems:
            story.append(Paragraph(
                f"{icons.get(p['status'],'?')} {p['title']} — {names.get(p['status'], p['status'])} ({fmt_date(p['created_at'])})",
                normal
            ))
    else:
        story.append(Paragraph("Проблем нет", normal))

    doc.build(story)
    buf.seek(0)
    return buf

# ─── Excel Export ────────────────────────────────────────────────────────────

def generate_excel(obj):
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Объект"
    ws1.column_dimensions['A'].width = 25
    ws1.column_dimensions['B'].width = 45
    header_fill = PatternFill("solid", fgColor="1a3a6e")
    bold_white  = Font(bold=True, color="FFFFFF")
    ws1.append(["Отчёт по объекту", obj['name']])
    ws1.append(["Дата отчёта", fmt_date(datetime.now())])
    ws1.append([])
    for row in [
        ["Адрес", obj.get('address') or '—'],
        ["Заказчик", obj.get('customer') or '—'],
        ["Инженер", obj.get('engineer') or '—'],
        ["Статус", status_label(obj['status'])],
        ["Кабель", f"{obj['total_cable']} м"],
        ["Устройства", f"{obj['total_devices']} шт"],
    ]:
        ws1.append(row)
        ws1.cell(row=ws1.max_row, column=1).font = Font(bold=True)

    ws2 = wb.create_sheet("История работ")
    ws2.column_dimensions['A'].width = 20
    ws2.column_dimensions['B'].width = 70
    ws2.append(["Дата", "Запись"])
    ws2['A1'].fill = header_fill
    ws2['A1'].font = bold_white
    ws2['B1'].fill = header_fill
    ws2['B1'].font = bold_white
    for h in db_get_history(obj['id'], limit=100):
        ws2.append([fmt_date(h['created_at']), h['text']])

    ws3 = wb.create_sheet("Проблемы")
    ws3.column_dimensions['A'].width = 50
    ws3.column_dimensions['B'].width = 15
    ws3.column_dimensions['C'].width = 20
    ws3.append(["Проблема", "Статус", "Дата"])
    ws3['A1'].fill = header_fill; ws3['A1'].font = bold_white
    ws3['B1'].fill = header_fill; ws3['B1'].font = bold_white
    ws3['C1'].fill = header_fill; ws3['C1'].font = bold_white
    names = {"open": "Открыта", "inwork": "В работе", "solved": "Решена"}
    for p in db_get_problems(obj['id']):
        ws3.append([p['title'], names.get(p['status'], p['status']), fmt_date(p['created_at'])])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─── Notifications ────────────────────────────────────────────────────────────

async def send_notifications(app):
    while True:
        await asyncio.sleep(3600 * 24)  # раз в сутки
        try:
            objects_with_problems = db_get_all_objects_with_open_problems()
            if not objects_with_problems:
                continue
            with get_db() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT uid FROM users WHERE role='admin'")
                    admins = cur.fetchall()
            for admin in admins:
                lines = ["⚠️ *Напоминание: открытые проблемы*\n"]
                for o in objects_with_problems:
                    lines.append(f"🔴 {o['name']} — {o['open_problems']} проблем(ы)")
                try:
                    await app.bot.send_message(admin['uid'], "\n".join(lines), parse_mode="Markdown")
                except Exception as e:
                    logger.error(f"Notification error for {admin['uid']}: {e}")
        except Exception as e:
            logger.error(f"Notification loop error: {e}")

# ─── Handlers ────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid  = update.effective_user.id
    name = update.effective_user.first_name or "коллега"
    db_upsert_user(uid, name)

    if uid in ADMIN_IDS:
        db_set_role(uid, 'admin')

    objects = db_get_objects(uid)
    text = (
        f"👋 Привет, {name}!\n\nЯ — AI-диспетчер объектов ОПС.\n\n"
        "Помогу:\n• фиксировать ежедневные работы\n"
        "• хранить историю\n• отслеживать проблемы\n"
        "• формировать отчёты PDF/Excel\n\n"
    )
    if objects:
        await update.message.reply_text(text + "Выберите объект:", reply_markup=objects_keyboard(uid))
    else:
        await update.message.reply_text(text + "Создайте первый объект: /new")

async def cmd_new(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📋 Введите название нового объекта:")
    ctx.user_data["awaiting"] = "object_name"

async def cmd_objects(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    objects = db_get_objects(uid)
    if not objects:
        await update.message.reply_text("Объектов нет. Создайте первый: /new")
        return
    await update.message.reply_text("📂 Ваши объекты:", reply_markup=objects_keyboard(uid))

async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    obj = db_get_current_object(uid)
    if not obj:
        await update.message.reply_text("Объект не выбран. Используйте /objects")
        return
    probs_open = db_get_problems(obj['id'], status='open')
    history    = db_get_history(obj['id'], limit=1)
    last_str   = f"\n📝 Последняя запись: {fmt_date(history[0]['created_at'])}" if history else ""
    text = (
        f"{status_emoji(obj['status'])} *{obj['name']}*\n"
        f"📍 {obj.get('address') or 'не указан'}\n"
        f"👤 {obj.get('engineer') or 'не указан'}\n"
        f"🏢 {obj.get('customer') or 'не указан'}\n\n"
        f"Статус: *{status_label(obj['status'])}*\n"
        f"🔌 Кабель: {obj['total_cable']} м\n"
        f"📡 Устройства: {obj['total_devices']} шт\n"
        f"⚠️ Открытых проблем: {len(probs_open)}"
        f"{last_str}"
    )
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🟢 В работе",  callback_data="status:active"),
         InlineKeyboardButton("🟡 Пауза",     callback_data="status:paused")],
        [InlineKeyboardButton("🔴 Проблема",  callback_data="status:problem"),
         InlineKeyboardButton("✅ Завершён",   callback_data="status:done")],
        [InlineKeyboardButton("📋 История",   callback_data="history"),
         InlineKeyboardButton("⚠️ Проблемы",  callback_data="problems")],
        [InlineKeyboardButton("📄 PDF отчёт", callback_data="export:pdf"),
         InlineKeyboardButton("📊 Excel",     callback_data="export:excel")],
        [InlineKeyboardButton("🖼 Фото",      callback_data="photos")],
    ])
    await update.message.reply_text(text, parse_mode="Markdown", reply_markup=keyboard)

async def cmd_history(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message or update.callback_query.message
    uid = update.effective_user.id
    obj = db_get_current_object(uid)
    if not obj:
        await msg.reply_text("Объект не выбран. Используйте /objects")
        return
    history = db_get_history(obj['id'], limit=10)
    if not history:
        await msg.reply_text(f"📋 История по «{obj['name']}» пуста.")
        return
    lines = [f"📋 *История: {obj['name']}*\n"]
    for h in history:
        lines.append(f"📅 {fmt_date(h['created_at'])}\n{h['text']}\n")
    await msg.reply_text("\n".join(lines), parse_mode="Markdown")

async def cmd_problems(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message or update.callback_query.message
    uid = update.effective_user.id
    obj = db_get_current_object(uid)
    if not obj:
        await msg.reply_text("Объект не выбран. Используйте /objects")
        return
    problems = db_get_problems(obj['id'])
    if not problems:
        await msg.reply_text(f"✅ Проблем по «{obj['name']}» нет.")
        return
    icons = {"open": "🔴", "inwork": "🟡", "solved": "🟢"}
    names = {"open": "Открыта", "inwork": "В работе", "solved": "Решена"}
    lines = [f"⚠️ *Проблемы: {obj['name']}*\n"]
    buttons = []
    for p in problems:
        lines.append(f"{icons.get(p['status'],'⚪')} {p['title']}\n   {names.get(p['status'])} · {fmt_date(p['created_at'])}\n")
        if p['status'] != 'solved':
            buttons.append([InlineKeyboardButton(f"✅ Закрыть: {p['title'][:30]}", callback_data=f"solve:{p['id']}")])
    kb = InlineKeyboardMarkup(buttons) if buttons else None
    await msg.reply_text("\n".join(lines), parse_mode="Markdown", reply_markup=kb)

async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    obj = db_get_current_object(uid)
    if not obj:
        await update.message.reply_text("Объект не выбран. Используйте /objects")
        return
    await update.message.reply_text("⏳ Формирую отчёт...")
    try:
        reply, _ = await ask_groq(uid, obj, "Сформируй краткий отчёт: что сделано, проблемы, текущий статус.")
        await update.message.reply_text(reply)
    except Exception as e:
        logger.error(f"Groq error: {e}")
        await update.message.reply_text("⚠️ Ошибка AI. Попробуйте позже.")

async def cmd_addadmin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not db_is_admin(uid):
        await update.message.reply_text("⛔ Нет доступа.")
        return
    args = ctx.args
    if not args:
        await update.message.reply_text("Использование: /addadmin [uid]")
        return
    try:
        target_uid = int(args[0])
        db_set_role(target_uid, 'admin')
        await update.message.reply_text(f"✅ Пользователь {target_uid} назначен администратором.")
    except ValueError:
        await update.message.reply_text("Неверный формат uid.")

async def cmd_myid(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Ваш Telegram ID: `{update.effective_user.id}`", parse_mode="Markdown")

async def cmd_help(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    is_admin = db_is_admin(uid)
    text = (
        "📖 *Команды:*\n\n"
        "/start — начало работы\n"
        "/new — создать объект\n"
        "/objects — список объектов\n"
        "/status — статус объекта\n"
        "/history — история работ\n"
        "/problems — проблемы\n"
        "/report — AI-отчёт\n"
        "/myid — узнать свой ID\n"
    )
    if is_admin:
        text += "\n👑 *Админ:*\n/addadmin [uid] — назначить администратора\n"
    text += "\n💬 Просто пишите что было сделано — бот сохранит.\n📷 Отправьте фото — оно привяжется к объекту."
    await update.message.reply_text(text, parse_mode="Markdown")

# ─── Photo handler ────────────────────────────────────────────────────────────

async def handle_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    obj = db_get_current_object(uid)
    if not obj:
        await update.message.reply_text("Сначала выберите объект: /objects")
        return
    photo   = update.message.photo[-1]
    caption = update.message.caption or ""
    db_add_photo(obj['id'], uid, photo.file_id, caption)
    await update.message.reply_text(
        f"📷 Фото сохранено\n"
        f"🏗 Объект: *{obj['name']}*\n"
        f"📝 Комментарий: {caption or 'нет'}\n"
        f"📅 {fmt_date(datetime.now())}",
        parse_mode="Markdown"
    )

# ─── Message handler ──────────────────────────────────────────────────────────

async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid      = update.effective_user.id
    text     = update.message.text.strip()
    awaiting = ctx.user_data.get("awaiting")
    obj      = db_get_current_object(uid)

    if awaiting == "object_name":
        obj_id = db_create_object(uid, text)
        ctx.user_data["awaiting"] = "object_address"
        ctx.user_data["new_obj_id"] = obj_id
        await update.message.reply_text(f"✅ Объект «{text}» создан!\n\nВведите адрес (или /skip):")
        return

    if awaiting == "object_address":
        if text != "/skip":
            db_update_object(ctx.user_data.get("new_obj_id"), address=text)
        ctx.user_data["awaiting"] = "object_customer"
        await update.message.reply_text("Введите заказчика (или /skip):")
        return

    if awaiting == "object_customer":
        if text != "/skip":
            db_update_object(ctx.user_data.get("new_obj_id"), customer=text)
        ctx.user_data["awaiting"] = "object_engineer"
        await update.message.reply_text("Введите ответственного инженера (или /skip):")
        return

    if awaiting == "object_engineer":
        if text != "/skip":
            db_update_object(ctx.user_data.get("new_obj_id"), engineer=text)
        ctx.user_data.pop("awaiting", None)
        ctx.user_data.pop("new_obj_id", None)
        await update.message.reply_text(
            "✅ Готово! Теперь пишите что было сделано.\n\n"
            "_Например: «Иван протянул 200 м кабеля на 3 этаже»_",
            parse_mode="Markdown"
        )
        return

    if not obj:
        objects = db_get_objects(uid)
        if objects:
            await update.message.reply_text("Выберите объект:", reply_markup=objects_keyboard(uid))
        else:
            await update.message.reply_text("Объектов нет. Создайте первый: /new")
        return

    await update.message.chat.send_action("typing")
    try:
        reply, _ = await ask_groq(uid, obj, text)
        await update.message.reply_text(reply)
    except Exception as e:
        logger.error(f"Groq error: {e}")
        await update.message.reply_text("⚠️ Ошибка AI. Попробуйте позже.")

# ─── Callback handler ────────────────────────────────────────────────────────

async def handle_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    uid  = query.from_user.id
    data = query.data

    if data == "new_object":
        await query.message.reply_text("Введите название нового объекта:")
        ctx.user_data["awaiting"] = "object_name"

    elif data.startswith("select:"):
        obj_id = int(data[7:])
        db_set_current_object(uid, obj_id)
        obj = db_get_object(obj_id)
        if obj:
            await query.message.reply_text(
                f"✅ Выбран: *{obj['name']}*\n"
                f"Статус: {status_emoji(obj['status'])} {status_label(obj['status'])}\n\n"
                "Пишите что было сделано или отправьте фото.",
                parse_mode="Markdown"
            )

    elif data.startswith("status:"):
        new_status = data[7:]
        obj = db_get_current_object(uid)
        if obj:
            db_update_object(obj['id'], status=new_status)
            await query.message.reply_text(f"✅ Статус: {status_emoji(new_status)} {status_label(new_status)}")

    elif data.startswith("solve:"):
        prob_id = int(data[6:])
        db_solve_problem(prob_id)
        await query.message.reply_text("✅ Проблема отмечена как решённая.")

    elif data == "history":
        await cmd_history(update, ctx)

    elif data == "problems":
        await cmd_problems(update, ctx)

    elif data == "photos":
        obj = db_get_current_object(uid)
        if not obj:
            await query.message.reply_text("Объект не выбран.")
            return
        photos = db_get_photos(obj['id'])
        if not photos:
            await query.message.reply_text(f"🖼 Фото по «{obj['name']}» нет.")
            return
        await query.message.reply_text(f"🖼 Фото по «{obj['name']}»: {len(photos)} шт.\nОтправляю последние 5...")
        for p in photos[:5]:
            await query.message.reply_photo(p['file_id'], caption=f"📅 {fmt_date(p['created_at'])}\n{p['caption']}")

    elif data.startswith("export:"):
        fmt = data[7:]
        obj = db_get_current_object(uid)
        if not obj:
            await query.message.reply_text("Объект не выбран.")
            return
        await query.message.reply_text("⏳ Генерирую файл...")
        try:
            if fmt == "pdf":
                buf = generate_pdf(obj)
                fname = f"report_{obj['name'].replace(' ','_')}.pdf"
                await query.message.reply_document(document=buf, filename=fname, caption=f"📄 Отчёт: {obj['name']}")
            elif fmt == "excel":
                buf = generate_excel(obj)
                fname = f"report_{obj['name'].replace(' ','_')}.xlsx"
                await query.message.reply_document(document=buf, filename=fname, caption=f"📊 Отчёт: {obj['name']}")
        except Exception as e:
            logger.error(f"Export error: {e}")
            await query.message.reply_text("⚠️ Ошибка при генерации файла.")

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    init_db()
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start",    cmd_start))
    app.add_handler(CommandHandler("new",      cmd_new))
    app.add_handler(CommandHandler("objects",  cmd_objects))
    app.add_handler(CommandHandler("status",   cmd_status))
    app.add_handler(CommandHandler("history",  cmd_history))
    app.add_handler(CommandHandler("problems", cmd_problems))
    app.add_handler(CommandHandler("report",   cmd_report))
    app.add_handler(CommandHandler("help",     cmd_help))
    app.add_handler(CommandHandler("myid",     cmd_myid))
    app.add_handler(CommandHandler("addadmin", cmd_addadmin))
    app.add_handler(CommandHandler("skip",     handle_message))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    loop = asyncio.get_event_loop()
    loop.create_task(send_notifications(app))

    logger.info("Бот v2 запущен!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
