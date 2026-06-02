import os
import re
import logging
import httpx
import asyncio
import secrets
from datetime import datetime
from io import BytesIO
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
import psycopg
from psycopg.rows import dict_row
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import openpyxl
from openpyxl.styles import Font, PatternFill

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ['TELEGRAM_TOKEN']
GROQ_API_KEY   = os.environ['GROQ_API_KEY']
DATABASE_URL   = os.environ['DATABASE_URL']
ADMIN_IDS      = [int(x) for x in os.environ.get('ADMIN_IDS', '').split(',') if x.strip()]
BOT_USERNAME   = os.environ.get('BOT_USERNAME', '')  # например: ops_my_bot
GROQ_URL       = "https://api.groq.com/openai/v1/chat/completions"

SYSTEM_PROMPT = """Ты — профессиональный AI‑ассистент для управления объектами охранно‑пожарной сигнализации (ОПС), систем оповещения и эвакуации (СОУЭ).

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

Отвечай кратко, профессионально. Все ответы только на русском языке."""

# ─── Database ────────────────────────────────────────────────────────────────

def get_db():
    return psycopg.connect(DATABASE_URL, row_factory=dict_row)

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
                CREATE TABLE IF NOT EXISTS invites (
                    token TEXT PRIMARY KEY,
                    created_by BIGINT,
                    used_by BIGINT DEFAULT NULL,
                    used_at TIMESTAMP DEFAULT NULL,
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
                    owner_uid BIGINT,
                    created_at TIMESTAMP DEFAULT NOW()
                );
                CREATE TABLE IF NOT EXISTS object_workers (
                    object_id INT REFERENCES objects(id) ON DELETE CASCADE,
                    uid BIGINT,
                    PRIMARY KEY (object_id, uid)
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
                CREATE TABLE IF NOT EXISTS user_state (
                    uid BIGINT PRIMARY KEY,
                    current_object_id INT,
                    state TEXT DEFAULT ''
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

def db_upsert_user(uid, name, role='worker'):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO users (uid, name, role) VALUES (%s, %s, %s)
                ON CONFLICT (uid) DO UPDATE SET name=EXCLUDED.name
            """, (uid, name, role))
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

def db_is_registered(uid):
    return db_get_user(uid) is not None

def db_get_objects(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            if db_is_admin(uid):
                cur.execute("SELECT * FROM objects ORDER BY created_at DESC")
            else:
                cur.execute("""
                    SELECT DISTINCT o.* FROM objects o
                    LEFT JOIN object_workers ow ON ow.object_id=o.id
                    WHERE o.owner_uid=%s OR ow.uid=%s
                    ORDER BY o.created_at DESC
                """, (uid, uid))
            return cur.fetchall()

def db_get_object(obj_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM objects WHERE id=%s", (obj_id,))
            return cur.fetchone()

def db_create_object(uid, name):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO objects (name, owner_uid) VALUES (%s,%s) RETURNING id", (name, uid))
            obj_id = cur.fetchone()['id']
            cur.execute("""
                INSERT INTO user_state (uid, current_object_id) VALUES (%s,%s)
                ON CONFLICT (uid) DO UPDATE SET current_object_id=EXCLUDED.current_object_id
            """, (uid, obj_id))
            conn.commit()
            return obj_id

def db_update_object(obj_id, **kwargs):
    if not kwargs: return
    fields = ", ".join(f"{k}=%s" for k in kwargs)
    vals = list(kwargs.values()) + [obj_id]
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute(f"UPDATE objects SET {fields} WHERE id=%s", vals)
            conn.commit()

def db_get_current_object(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT current_object_id FROM user_state WHERE uid=%s", (uid,))
            row = cur.fetchone()
            if not row or not row['current_object_id']: return None
            return db_get_object(row['current_object_id'])

def db_set_current_object(uid, obj_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO user_state (uid, current_object_id) VALUES (%s,%s)
                ON CONFLICT (uid) DO UPDATE SET current_object_id=EXCLUDED.current_object_id
            """, (uid, obj_id))
            conn.commit()

def db_get_state(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT state FROM user_state WHERE uid=%s", (uid,))
            row = cur.fetchone()
            return row['state'] if row else ''

def db_set_state(uid, state):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO user_state (uid, state) VALUES (%s,%s)
                ON CONFLICT (uid) DO UPDATE SET state=EXCLUDED.state
            """, (uid, state))
            conn.commit()

def db_add_history(obj_id, uid, text):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO history (object_id,uid,text) VALUES (%s,%s,%s)", (obj_id, uid, text))
            conn.commit()

def db_get_history(obj_id, limit=10):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM history WHERE object_id=%s ORDER BY created_at DESC LIMIT %s", (obj_id, limit))
            return list(reversed(cur.fetchall()))

def db_add_problem(obj_id, title):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO problems (object_id,title) VALUES (%s,%s)", (obj_id, title))
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
            cur.execute("INSERT INTO photos (object_id,uid,file_id,caption) VALUES (%s,%s,%s,%s)", (obj_id, uid, file_id, caption))
            conn.commit()

def db_get_photos(obj_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM photos WHERE object_id=%s ORDER BY created_at DESC", (obj_id,))
            return cur.fetchall()

def db_get_workers(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE role='worker' ORDER BY name")
            return cur.fetchall()

def db_assign_worker(obj_id, worker_uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO object_workers (object_id,uid) VALUES (%s,%s) ON CONFLICT DO NOTHING", (obj_id, worker_uid))
            conn.commit()

def db_create_invite(admin_uid):
    token = secrets.token_urlsafe(16)
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO invites (token,created_by) VALUES (%s,%s)", (token, admin_uid))
            conn.commit()
    return token

def db_use_invite(token, uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM invites WHERE token=%s AND used_by IS NULL", (token,))
            inv = cur.fetchone()
            if not inv: return False
            cur.execute("UPDATE invites SET used_by=%s, used_at=NOW() WHERE token=%s", (uid, token))
            conn.commit()
            return True

def db_get_chat_history(uid, obj_id, limit=10):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT role, content FROM chat_history
                WHERE uid=%s AND object_id=%s
                ORDER BY created_at DESC LIMIT %s
            """, (uid, obj_id, limit))
            return list(reversed(cur.fetchall()))

def db_add_chat(uid, obj_id, role, content):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO chat_history (uid,object_id,role,content) VALUES (%s,%s,%s,%s)", (uid, obj_id, role, content))
            conn.commit()

# ─── Helpers ─────────────────────────────────────────────────────────────────

def fmt_date(dt):
    if isinstance(dt, datetime): return dt.strftime("%d.%m.%Y %H:%M")
    return str(dt)

def status_emoji(s):
    return {"active":"🟢","paused":"🟡","problem":"🔴","done":"✅"}.get(s,"⚪")

def status_label(s):
    return {"active":"В работе","paused":"Приостановлен","problem":"Проблема","done":"Завершён"}.get(s,s)

def parse_meta(text):
    meta = {"cable":0,"devices":0,"has_problem":False}
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

# ─── Keyboards ───────────────────────────────────────────────────────────────

def main_keyboard_admin():
    return ReplyKeyboardMarkup([
        ["📂 Объекты",      "➕ Новый объект"],
        ["👷 Сотрудники",   "🔗 Пригласить"],
        ["⚠️ Проблемы",     "📊 Статус"],
        ["📄 PDF отчёт",    "📋 Excel отчёт"],
    ], resize_keyboard=True)

def main_keyboard_worker():
    return ReplyKeyboardMarkup([
        ["📋 Внести работы",  "📷 Отправить фото"],
        ["📂 Мои объекты",    "⚠️ Проблема"],
        ["📊 Статус объекта", "📜 История"],
    ], resize_keyboard=True)

def objects_inline(uid):
    objects = db_get_objects(uid)
    buttons = []
    for o in objects:
        buttons.append([InlineKeyboardButton(
            f"{status_emoji(o['status'])} {o['name']}", callback_data=f"sel:{o['id']}"
        )])
    return InlineKeyboardMarkup(buttons)

# ─── Groq AI ─────────────────────────────────────────────────────────────────

async def ask_groq(uid, obj, user_message):
    problems_open = db_get_problems(obj['id'], status='open')
    history = db_get_history(obj['id'], limit=5)
    context = (
        f"Объект: «{obj['name']}» | Адрес: {obj.get('address') or '—'}\n"
        f"Заказчик: {obj.get('customer') or '—'} | Инженер: {obj.get('engineer') or '—'}\n"
        f"Статус: {status_label(obj['status'])} | Кабель: {obj['total_cable']}м | Устройства: {obj['total_devices']}шт\n\n"
        "Последние записи:\n"
        + ("\n".join(f"• {fmt_date(h['created_at'])}: {h['text']}" for h in history) or "нет")
        + "\nОткрытые проблемы:\n"
        + ("\n".join(f"• {p['title']}" for p in problems_open) or "нет")
    )
    chat_hist = db_get_chat_history(uid, obj['id'], limit=10)
    messages = [{"role":"system","content":SYSTEM_PROMPT}]
    messages.append({"role":"user","content":context})
    messages.append({"role":"assistant","content":"Понял контекст. Готов к работе."})
    for m in chat_hist:
        messages.append({"role":m['role'],"content":m['content']})
    messages.append({"role":"user","content":user_message})

    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            GROQ_URL,
            headers={"Authorization":f"Bearer {GROQ_API_KEY}","Content-Type":"application/json"},
            json={"model":"llama-3.3-70b-versatile","messages":messages,"max_tokens":1024,"temperature":0.3}
        )
        resp.raise_for_status()
        data = resp.json()

    reply = data["choices"][0]["message"]["content"]
    db_add_chat(uid, obj['id'], "user", user_message)
    db_add_chat(uid, obj['id'], "assistant", reply)

    meta = parse_meta(reply)
    clean = clean_response(reply)

    if "📅" in reply or "✔" in reply:
        db_add_history(obj['id'], uid, user_message)
        db_update_object(obj['id'],
            total_cable=obj['total_cable']+meta['cable'],
            total_devices=obj['total_devices']+meta['devices']
        )
        if meta['has_problem']:
            db_add_problem(obj['id'], f"Проблема от {fmt_date(datetime.now())}")

    return clean, meta

# ─── Export ───────────────────────────────────────────────────────────────────

def generate_pdf(obj):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    normal = ParagraphStyle('n', parent=styles['Normal'], fontSize=10, spaceAfter=3)
    h2 = ParagraphStyle('h2', parent=styles['Heading2'], fontSize=12, spaceAfter=4)
    story = []
    story.append(Paragraph(f"Отчёт: {obj['name']}", styles['Title']))
    story.append(Paragraph(f"Сформирован: {fmt_date(datetime.now())}", normal))
    story.append(Spacer(1,12))
    info = [["Адрес",obj.get('address') or '—'],["Заказчик",obj.get('customer') or '—'],
            ["Инженер",obj.get('engineer') or '—'],["Статус",status_label(obj['status'])],
            ["Кабель",f"{obj['total_cable']} м"],["Устройства",f"{obj['total_devices']} шт"]]
    t = Table(info, colWidths=[150,330])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(0,-1),colors.HexColor('#f0f4ff')),
                            ('GRID',(0,0),(-1,-1),0.5,colors.grey),('PADDING',(0,0),(-1,-1),6)]))
    story.append(t); story.append(Spacer(1,12))
    story.append(Paragraph("История работ", h2))
    for h in db_get_history(obj['id'], limit=50):
        story.append(Paragraph(f"<b>{fmt_date(h['created_at'])}</b> — {h['text']}", normal))
    story.append(Spacer(1,12))
    story.append(Paragraph("Проблемы", h2))
    icons = {"open":"🔴","inwork":"🟡","solved":"🟢"}
    names = {"open":"Открыта","inwork":"В работе","solved":"Решена"}
    for p in db_get_problems(obj['id']):
        story.append(Paragraph(f"{icons.get(p['status'],'?')} {p['title']} — {names.get(p['status'])} ({fmt_date(p['created_at'])})", normal))
    doc.build(story); buf.seek(0)
    return buf

def generate_excel(obj):
    wb = openpyxl.Workbook()
    ws1 = wb.active; ws1.title = "Объект"
    ws1.column_dimensions['A'].width = 25; ws1.column_dimensions['B'].width = 45
    hf = PatternFill("solid", fgColor="1a3a6e"); bw = Font(bold=True, color="FFFFFF")
    for row in [["Объект",obj['name']],["Дата",fmt_date(datetime.now())],[],
                ["Адрес",obj.get('address') or '—'],["Заказчик",obj.get('customer') or '—'],
                ["Инженер",obj.get('engineer') or '—'],["Статус",status_label(obj['status'])],
                ["Кабель",f"{obj['total_cable']} м"],["Устройства",f"{obj['total_devices']} шт"]]:
        ws1.append(row)
        if row: ws1.cell(ws1.max_row,1).font = Font(bold=True)
    ws2 = wb.create_sheet("История"); ws2.column_dimensions['A'].width=20; ws2.column_dimensions['B'].width=70
    ws2.append(["Дата","Запись"]); ws2['A1'].fill=hf; ws2['A1'].font=bw; ws2['B1'].fill=hf; ws2['B1'].font=bw
    for h in db_get_history(obj['id'],limit=100): ws2.append([fmt_date(h['created_at']),h['text']])
    ws3 = wb.create_sheet("Проблемы"); ws3.column_dimensions['A'].width=50; ws3.column_dimensions['B'].width=15; ws3.column_dimensions['C'].width=20
    ws3.append(["Проблема","Статус","Дата"])
    for c in ['A1','B1','C1']: ws3[c].fill=hf; ws3[c].font=bw
    names = {"open":"Открыта","inwork":"В работе","solved":"Решена"}
    for p in db_get_problems(obj['id']): ws3.append([p['title'],names.get(p['status']),fmt_date(p['created_at'])])
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─── Handlers ────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid  = update.effective_user.id
    name = update.effective_user.first_name or "коллега"
    args = ctx.args

    # Обработка инвайт-ссылки: /start invite_TOKEN
    if args and args[0].startswith("invite_"):
        token = args[0][7:]
        if db_use_invite(token, uid):
            db_upsert_user(uid, name, role='worker')
            await update.message.reply_text(
                f"👋 Добро пожаловать, {name}!\n\nВы успешно зарегистрированы как монтажник.\nАдминистратор назначит вам объекты.",
                reply_markup=main_keyboard_worker()
            )
            return
        else:
            await update.message.reply_text("❌ Ссылка недействительна или уже использована.")
            return

    # Уже зарегистрирован
    if db_is_registered(uid):
        is_admin = db_is_admin(uid)
        kb = main_keyboard_admin() if is_admin else main_keyboard_worker()
        role_text = "администратор" if is_admin else "монтажник"
        await update.message.reply_text(
            f"👋 С возвращением, {name}!\nВы вошли как *{role_text}*.",
            parse_mode="Markdown", reply_markup=kb
        )
        return

    # Новый пользователь без инвайта
    if uid in ADMIN_IDS:
        db_upsert_user(uid, name, role='admin')
        await update.message.reply_text(
            f"👋 Привет, {name}! Вы — администратор системы.\n\nИспользуйте меню для управления объектами.",
            reply_markup=main_keyboard_admin()
        )
    else:
        await update.message.reply_text(
            "⛔ Доступ только по приглашению.\n\nПопросите администратора прислать вам ссылку-приглашение."
        )

async def cmd_invite(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not db_is_admin(uid):
        await update.message.reply_text("⛔ Только для администраторов.")
        return
    token = db_create_invite(uid)
    bot_username = BOT_USERNAME or (await ctx.bot.get_me()).username
    link = f"https://t.me/{bot_username}?start=invite_{token}"
    await update.message.reply_text(
        f"🔗 *Ссылка-приглашение для монтажника:*\n\n`{link}`\n\n"
        f"⚠️ Одноразовая — после использования сгорает.\n"
        f"Перешлите монтажнику в личку.",
        parse_mode="Markdown"
    )

async def cmd_myid(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Ваш Telegram ID: `{update.effective_user.id}`", parse_mode="Markdown")

async def show_objects(update, uid):
    objects = db_get_objects(uid)
    if not objects:
        await update.message.reply_text("Объектов нет." + (" Создайте: нажмите '➕ Новый объект'" if db_is_admin(uid) else " Ожидайте назначения от администратора."))
        return
    await update.message.reply_text("📂 Выберите объект:", reply_markup=objects_inline(uid))

async def show_status(update, uid):
    obj = db_get_current_object(uid)
    if not obj:
        await update.message.reply_text("Объект не выбран. Сначала выберите из списка.")
        return
    probs = db_get_problems(obj['id'], status='open')
    hist  = db_get_history(obj['id'], limit=1)
    last  = f"\n📝 Последняя запись: {fmt_date(hist[0]['created_at'])}" if hist else ""
    text  = (
        f"{status_emoji(obj['status'])} *{obj['name']}*\n"
        f"📍 {obj.get('address') or '—'}\n👤 {obj.get('engineer') or '—'}\n🏢 {obj.get('customer') or '—'}\n\n"
        f"Статус: *{status_label(obj['status'])}*\n"
        f"🔌 Кабель: {obj['total_cable']} м\n📡 Устройства: {obj['total_devices']} шт\n"
        f"⚠️ Открытых проблем: {len(probs)}{last}"
    )
    buttons = [
        [InlineKeyboardButton("🟢 В работе",  callback_data="st:active"),
         InlineKeyboardButton("🟡 Пауза",     callback_data="st:paused")],
        [InlineKeyboardButton("🔴 Проблема",  callback_data="st:problem"),
         InlineKeyboardButton("✅ Завершён",   callback_data="st:done")],
        [InlineKeyboardButton("📜 История",   callback_data="history"),
         InlineKeyboardButton("⚠️ Проблемы",  callback_data="problems")],
    ]
    if db_is_admin(uid):
        buttons.append([
            InlineKeyboardButton("📄 PDF", callback_data="exp:pdf"),
            InlineKeyboardButton("📊 Excel", callback_data="exp:excel"),
            InlineKeyboardButton("👷 Назначить", callback_data="assign"),
        ])
    await update.message.reply_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(buttons))

async def show_history(msg, uid):
    obj = db_get_current_object(uid)
    if not obj:
        await msg.reply_text("Объект не выбран.")
        return
    history = db_get_history(obj['id'], limit=10)
    if not history:
        await msg.reply_text(f"📜 История по «{obj['name']}» пуста.")
        return
    lines = [f"📜 *История: {obj['name']}*\n"]
    for h in history:
        lines.append(f"📅 {fmt_date(h['created_at'])}\n{h['text']}\n")
    await msg.reply_text("\n".join(lines), parse_mode="Markdown")

async def show_problems(msg, uid):
    obj = db_get_current_object(uid)
    if not obj:
        await msg.reply_text("Объект не выбран.")
        return
    problems = db_get_problems(obj['id'])
    if not problems:
        await msg.reply_text(f"✅ Проблем по «{obj['name']}» нет.")
        return
    icons = {"open":"🔴","inwork":"🟡","solved":"🟢"}
    names = {"open":"Открыта","inwork":"В работе","solved":"Решена"}
    lines = [f"⚠️ *Проблемы: {obj['name']}*\n"]
    buttons = []
    for p in problems:
        lines.append(f"{icons.get(p['status'],'⚪')} {p['title']}\n   {names.get(p['status'])} · {fmt_date(p['created_at'])}\n")
        if p['status'] != 'solved':
            buttons.append([InlineKeyboardButton(f"✅ {p['title'][:35]}", callback_data=f"solve:{p['id']}")])
    await msg.reply_text("\n".join(lines), parse_mode="Markdown",
                         reply_markup=InlineKeyboardMarkup(buttons) if buttons else None)

# ─── Photo handler ────────────────────────────────────────────────────────────

async def handle_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not db_is_registered(uid):
        await update.message.reply_text("⛔ Нет доступа.")
        return
    obj = db_get_current_object(uid)
    if not obj:
        await update.message.reply_text("Сначала выберите объект.")
        return
    photo   = update.message.photo[-1]
    caption = update.message.caption or ""
    db_add_photo(obj['id'], uid, photo.file_id, caption)
    await update.message.reply_text(
        f"📷 Фото сохранено!\n🏗 *{obj['name']}*\n📝 {caption or 'без комментария'}\n📅 {fmt_date(datetime.now())}",
        parse_mode="Markdown"
    )

# ─── Message handler ──────────────────────────────────────────────────────────

async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid  = update.effective_user.id
    text = update.message.text.strip()

    if not db_is_registered(uid):
        await update.message.reply_text("⛔ Доступ только по приглашению от администратора.")
        return

    is_admin = db_is_admin(uid)
    state    = db_get_state(uid)

    # ── Состояния (ввод данных) ───────────────────────────────────────────────
    if state == "new_object_name":
        obj_id = db_create_object(uid, text)
        ctx.user_data["new_obj_id"] = obj_id
        db_set_state(uid, "new_object_address")
        await update.message.reply_text(f"✅ Объект «{text}» создан!\n\nВведите адрес (или /skip):")
        return

    if state == "new_object_address":
        if text != "/skip": db_update_object(ctx.user_data.get("new_obj_id"), address=text)
        db_set_state(uid, "new_object_customer")
        await update.message.reply_text("Введите заказчика (или /skip):")
        return

    if state == "new_object_customer":
        if text != "/skip": db_update_object(ctx.user_data.get("new_obj_id"), customer=text)
        db_set_state(uid, "new_object_engineer")
        await update.message.reply_text("Введите ответственного инженера (или /skip):")
        return

    if state == "new_object_engineer":
        if text != "/skip": db_update_object(ctx.user_data.get("new_obj_id"), engineer=text)
        db_set_state(uid, "")
        await update.message.reply_text(
            "✅ Объект создан и настроен!\n\nТеперь пишите что было сделано — я всё зафиксирую.",
            reply_markup=main_keyboard_admin() if is_admin else main_keyboard_worker()
        )
        return

    if state == "report_problem":
        obj = db_get_current_object(uid)
        if obj:
            db_add_problem(obj['id'], text)
            db_set_state(uid, "")
            await update.message.reply_text(f"⚠️ Проблема зафиксирована:\n«{text}»")
        return

    # ── Кнопки меню ──────────────────────────────────────────────────────────

    if text == "📂 Объекты" or text == "📂 Мои объекты":
        await show_objects(update, uid)
        return

    if text == "➕ Новый объект":
        if not is_admin:
            await update.message.reply_text("⛔ Только администратор может создавать объекты.")
            return
        db_set_state(uid, "new_object_name")
        await update.message.reply_text("📋 Введите название нового объекта:")
        return

    if text == "📊 Статус" or text == "📊 Статус объекта":
        await show_status(update, uid)
        return

    if text == "📜 История":
        obj = db_get_current_object(uid)
        if obj: await show_history(update.message, uid)
        else: await update.message.reply_text("Объект не выбран.")
        return

    if text == "⚠️ Проблемы" or text == "⚠️ Проблема":
        obj = db_get_current_object(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.")
            return
        if is_admin:
            await show_problems(update.message, uid)
        else:
            db_set_state(uid, "report_problem")
            await update.message.reply_text("Опишите проблему:")
        return

    if text == "🔗 Пригласить":
        if not is_admin:
            await update.message.reply_text("⛔ Только для администраторов.")
            return
        await cmd_invite(update, ctx)
        return

    if text == "👷 Сотрудники":
        if not is_admin:
            return
        workers = db_get_workers(uid)
        if not workers:
            await update.message.reply_text("Монтажников ещё нет. Пригласите через 🔗 Пригласить")
            return
        lines = ["👷 *Зарегистрированные монтажники:*\n"]
        for w in workers:
            lines.append(f"• {w['name']} (ID: {w['uid']})")
        await update.message.reply_text("\n".join(lines), parse_mode="Markdown")
        return

    if text == "📄 PDF отчёт":
        obj = db_get_current_object(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.")
            return
        await update.message.reply_text("⏳ Генерирую PDF...")
        buf = generate_pdf(obj)
        await update.message.reply_document(buf, filename=f"report_{obj['name']}.pdf", caption=f"📄 {obj['name']}")
        return

    if text == "📋 Excel отчёт":
        obj = db_get_current_object(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.")
            return
        await update.message.reply_text("⏳ Генерирую Excel...")
        buf = generate_excel(obj)
        await update.message.reply_document(buf, filename=f"report_{obj['name']}.xlsx", caption=f"📊 {obj['name']}")
        return

    if text == "📋 Внести работы":
        obj = db_get_current_object(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект из списка.", reply_markup=objects_inline(uid))
            return
        await update.message.reply_text(f"✍️ Пишите что было сделано на объекте *{obj['name']}*:", parse_mode="Markdown")
        return

    if text == "📷 Отправить фото":
        obj = db_get_current_object(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.")
            return
        await update.message.reply_text(f"📷 Отправьте фото — оно привяжется к *{obj['name']}*", parse_mode="Markdown")
        return

    # ── AI обработка свободного текста ───────────────────────────────────────
    obj = db_get_current_object(uid)
    if not obj:
        await update.message.reply_text("Сначала выберите объект:", reply_markup=objects_inline(uid))
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
    is_admin = db_is_admin(uid)

    if data.startswith("sel:"):
        obj_id = int(data[4:])
        db_set_current_object(uid, obj_id)
        obj = db_get_object(obj_id)
        if obj:
            await query.message.reply_text(
                f"✅ Выбран: *{obj['name']}*\n{status_emoji(obj['status'])} {status_label(obj['status'])}",
                parse_mode="Markdown",
                reply_markup=main_keyboard_admin() if is_admin else main_keyboard_worker()
            )

    elif data.startswith("st:"):
        obj = db_get_current_object(uid)
        if obj:
            db_update_object(obj['id'], status=data[3:])
            await query.message.reply_text(f"✅ Статус: {status_emoji(data[3:])} {status_label(data[3:])}")

    elif data.startswith("solve:"):
        db_solve_problem(int(data[6:]))
        await query.message.reply_text("✅ Проблема закрыта.")

    elif data == "history":
        await show_history(query.message, uid)

    elif data == "problems":
        await show_problems(query.message, uid)

    elif data.startswith("exp:"):
        obj = db_get_current_object(uid)
        if not obj or not is_admin: return
        await query.message.reply_text("⏳ Генерирую...")
        if data[4:] == "pdf":
            buf = generate_pdf(obj)
            await query.message.reply_document(buf, filename=f"report_{obj['name']}.pdf")
        else:
            buf = generate_excel(obj)
            await query.message.reply_document(buf, filename=f"report_{obj['name']}.xlsx")

    elif data == "assign":
        if not is_admin: return
        workers = db_get_workers(uid)
        if not workers:
            await query.message.reply_text("Монтажников нет. Пригласите через 🔗 Пригласить")
            return
        obj = db_get_current_object(uid)
        buttons = [[InlineKeyboardButton(w['name'], callback_data=f"asgn:{obj['id']}:{w['uid']}")] for w in workers]
        await query.message.reply_text("Выберите монтажника для назначения на объект:", reply_markup=InlineKeyboardMarkup(buttons))

    elif data.startswith("asgn:"):
        _, obj_id, worker_uid = data.split(":")
        db_assign_worker(int(obj_id), int(worker_uid))
        obj = db_get_object(int(obj_id))
        worker = db_get_user(int(worker_uid))
        await query.message.reply_text(f"✅ {worker['name']} назначен на объект «{obj['name']}»")
        try:
            await ctx.bot.send_message(
                int(worker_uid),
                f"📋 Вам назначен объект: *{obj['name']}*\n📍 {obj.get('address') or '—'}",
                parse_mode="Markdown"
            )
        except: pass

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    init_db()
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start",  cmd_start))
    app.add_handler(CommandHandler("invite", cmd_invite))
    app.add_handler(CommandHandler("myid",   cmd_myid))
    app.add_handler(CommandHandler("skip",   handle_message))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    logger.info("Бот v3 запущен!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
