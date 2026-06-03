import os
import re
import logging
import httpx
import asyncio
import secrets
import tempfile
from datetime import datetime, timedelta
from io import BytesIO
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
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

TELEGRAM_TOKEN  = os.environ['TELEGRAM_TOKEN']
GROQ_API_KEY    = os.environ['GROQ_API_KEY']
DATABASE_URL    = os.environ['DATABASE_URL']
ADMIN_IDS       = [int(x) for x in os.environ.get('ADMIN_IDS', '').split(',') if x.strip()]
BOT_USERNAME    = os.environ.get('BOT_USERNAME', '')
DIGEST_HOUR     = int(os.environ.get('DIGEST_HOUR', '9'))   # время дайджеста (UTC+3 → UTC = -3)
TIMEZONE_OFFSET = int(os.environ.get('TIMEZONE_OFFSET', '3'))  # часовой пояс

GROQ_URL        = "https://api.groq.com/openai/v1/chat/completions"
GROQ_AUDIO_URL  = "https://api.groq.com/openai/v1/audio/transcriptions"

SYSTEM_PROMPT = """Ты — профессиональный AI‑ассистент для управления объектами охранно‑пожарной сигнализации (ОПС).

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
                    plan_cable INT DEFAULT 0,
                    plan_devices INT DEFAULT 0,
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
                    cable INT DEFAULT 0,
                    devices INT DEFAULT 0,
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
                CREATE TABLE IF NOT EXISTS tasks (
                    id SERIAL PRIMARY KEY,
                    object_id INT REFERENCES objects(id) ON DELETE CASCADE,
                    worker_uid BIGINT,
                    created_by BIGINT,
                    title TEXT,
                    status TEXT DEFAULT 'open',
                    created_at TIMESTAMP DEFAULT NOW()
                );
                CREATE TABLE IF NOT EXISTS material_requests (
                    id SERIAL PRIMARY KEY,
                    object_id INT REFERENCES objects(id) ON DELETE CASCADE,
                    worker_uid BIGINT,
                    text TEXT,
                    status TEXT DEFAULT 'new',
                    created_at TIMESTAMP DEFAULT NOW()
                );
                CREATE TABLE IF NOT EXISTS locations (
                    id SERIAL PRIMARY KEY,
                    object_id INT REFERENCES objects(id) ON DELETE CASCADE,
                    uid BIGINT,
                    lat FLOAT,
                    lon FLOAT,
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
    logger.info("БД инициализирована v4")

# ─── DB helpers ──────────────────────────────────────────────────────────────

def get_db_user(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE uid=%s", (uid,))
            return cur.fetchone()

def upsert_user(uid, name, role='worker'):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO users (uid, name, role) VALUES (%s,%s,%s)
                ON CONFLICT (uid) DO UPDATE SET name=EXCLUDED.name
            """, (uid, name, role))
            conn.commit()

def set_role(uid, role):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE users SET role=%s WHERE uid=%s", (role, uid))
            conn.commit()

def is_admin(uid):
    if uid in ADMIN_IDS: return True
    u = get_db_user(uid)
    return u and u['role'] == 'admin'

def is_registered(uid):
    return get_db_user(uid) is not None

def get_objects(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            if is_admin(uid):
                cur.execute("SELECT * FROM objects ORDER BY created_at DESC")
            else:
                cur.execute("""
                    SELECT DISTINCT o.* FROM objects o
                    LEFT JOIN object_workers ow ON ow.object_id=o.id
                    WHERE o.owner_uid=%s OR ow.uid=%s ORDER BY o.created_at DESC
                """, (uid, uid))
            return cur.fetchall()

def get_object(obj_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM objects WHERE id=%s", (obj_id,))
            return cur.fetchone()

def create_object(uid, name):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO objects (name,owner_uid) VALUES (%s,%s) RETURNING id", (name, uid))
            obj_id = cur.fetchone()['id']
            cur.execute("""
                INSERT INTO user_state (uid, current_object_id) VALUES (%s,%s)
                ON CONFLICT (uid) DO UPDATE SET current_object_id=EXCLUDED.current_object_id
            """, (uid, obj_id))
            conn.commit()
            return obj_id

def update_object(obj_id, **kwargs):
    if not kwargs: return
    fields = ", ".join(f"{k}=%s" for k in kwargs)
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute(f"UPDATE objects SET {fields} WHERE id=%s", list(kwargs.values())+[obj_id])
            conn.commit()

def get_current_obj(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT current_object_id FROM user_state WHERE uid=%s", (uid,))
            row = cur.fetchone()
            if not row or not row['current_object_id']: return None
            return get_object(row['current_object_id'])

def set_current_obj(uid, obj_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO user_state (uid, current_object_id) VALUES (%s,%s)
                ON CONFLICT (uid) DO UPDATE SET current_object_id=EXCLUDED.current_object_id
            """, (uid, obj_id))
            conn.commit()

def get_state(uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT state FROM user_state WHERE uid=%s", (uid,))
            row = cur.fetchone()
            return row['state'] if row else ''

def set_state(uid, state):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO user_state (uid, state) VALUES (%s,%s)
                ON CONFLICT (uid) DO UPDATE SET state=EXCLUDED.state
            """, (uid, state))
            conn.commit()

def add_history(obj_id, uid, text, cable=0, devices=0):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO history (object_id,uid,text,cable,devices) VALUES (%s,%s,%s,%s,%s)",
                        (obj_id, uid, text, cable, devices))
            conn.commit()

def get_history(obj_id, limit=10):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM history WHERE object_id=%s ORDER BY created_at DESC LIMIT %s", (obj_id, limit))
            return list(reversed(cur.fetchall()))

def get_history_since(obj_id, since_dt):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM history WHERE object_id=%s AND created_at>=%s ORDER BY created_at", (obj_id, since_dt))
            return cur.fetchall()

def add_problem(obj_id, title):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO problems (object_id,title) VALUES (%s,%s)", (obj_id, title))
            conn.commit()

def get_problems(obj_id, status=None):
    with get_db() as conn:
        with conn.cursor() as cur:
            if status:
                cur.execute("SELECT * FROM problems WHERE object_id=%s AND status=%s ORDER BY created_at DESC", (obj_id, status))
            else:
                cur.execute("SELECT * FROM problems WHERE object_id=%s ORDER BY created_at DESC", (obj_id,))
            return cur.fetchall()

def solve_problem(prob_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE problems SET status='solved' WHERE id=%s", (prob_id,))
            conn.commit()

def add_photo(obj_id, uid, file_id, caption):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO photos (object_id,uid,file_id,caption) VALUES (%s,%s,%s,%s)", (obj_id, uid, file_id, caption))
            conn.commit()

def add_location(obj_id, uid, lat, lon):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO locations (object_id,uid,lat,lon) VALUES (%s,%s,%s,%s)", (obj_id, uid, lat, lon))
            conn.commit()

def get_workers():
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM users WHERE role='worker' ORDER BY name")
            return cur.fetchall()

def assign_worker(obj_id, worker_uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO object_workers (object_id,uid) VALUES (%s,%s) ON CONFLICT DO NOTHING", (obj_id, worker_uid))
            conn.commit()

def create_invite(admin_uid):
    token = secrets.token_urlsafe(16)
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO invites (token,created_by) VALUES (%s,%s)", (token, admin_uid))
            conn.commit()
    return token

def use_invite(token, uid):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM invites WHERE token=%s AND used_by IS NULL", (token,))
            inv = cur.fetchone()
            if not inv: return False
            cur.execute("UPDATE invites SET used_by=%s, used_at=NOW() WHERE token=%s", (uid, token))
            conn.commit()
            return True

def get_chat_hist(uid, obj_id, limit=10):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT role, content FROM chat_history
                WHERE uid=%s AND object_id=%s ORDER BY created_at DESC LIMIT %s
            """, (uid, obj_id, limit))
            return list(reversed(cur.fetchall()))

def add_chat(uid, obj_id, role, content):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO chat_history (uid,object_id,role,content) VALUES (%s,%s,%s,%s)", (uid, obj_id, role, content))
            conn.commit()

def create_task(obj_id, worker_uid, created_by, title):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO tasks (object_id,worker_uid,created_by,title) VALUES (%s,%s,%s,%s) RETURNING id",
                        (obj_id, worker_uid, created_by, title))
            task_id = cur.fetchone()['id']
            conn.commit()
            return task_id

def get_tasks(worker_uid, status=None):
    with get_db() as conn:
        with conn.cursor() as cur:
            if status:
                cur.execute("""
                    SELECT t.*, o.name as obj_name FROM tasks t
                    JOIN objects o ON o.id=t.object_id
                    WHERE t.worker_uid=%s AND t.status=%s ORDER BY t.created_at DESC
                """, (worker_uid, status))
            else:
                cur.execute("""
                    SELECT t.*, o.name as obj_name FROM tasks t
                    JOIN objects o ON o.id=t.object_id
                    WHERE t.worker_uid=%s ORDER BY t.created_at DESC
                """, (worker_uid,))
            return cur.fetchall()

def get_all_tasks_for_admin(obj_id=None):
    with get_db() as conn:
        with conn.cursor() as cur:
            if obj_id:
                cur.execute("""
                    SELECT t.*, o.name as obj_name, u.name as worker_name FROM tasks t
                    JOIN objects o ON o.id=t.object_id
                    JOIN users u ON u.uid=t.worker_uid
                    WHERE t.object_id=%s ORDER BY t.created_at DESC
                """, (obj_id,))
            else:
                cur.execute("""
                    SELECT t.*, o.name as obj_name, u.name as worker_name FROM tasks t
                    JOIN objects o ON o.id=t.object_id
                    JOIN users u ON u.uid=t.worker_uid
                    ORDER BY t.created_at DESC LIMIT 30
                """)
            return cur.fetchall()

def complete_task(task_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE tasks SET status='done' WHERE id=%s", (task_id,))
            conn.commit()

def add_material_request(obj_id, worker_uid, text):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO material_requests (object_id,worker_uid,text) VALUES (%s,%s,%s)", (obj_id, worker_uid, text))
            conn.commit()

def get_material_requests(status=None):
    with get_db() as conn:
        with conn.cursor() as cur:
            if status:
                cur.execute("""
                    SELECT mr.*, o.name as obj_name, u.name as worker_name FROM material_requests mr
                    JOIN objects o ON o.id=mr.object_id
                    JOIN users u ON u.uid=mr.worker_uid
                    WHERE mr.status=%s ORDER BY mr.created_at DESC
                """, (status,))
            else:
                cur.execute("""
                    SELECT mr.*, o.name as obj_name, u.name as worker_name FROM material_requests mr
                    JOIN objects o ON o.id=mr.object_id
                    JOIN users u ON u.uid=mr.worker_uid
                    ORDER BY mr.created_at DESC LIMIT 30
                """)
            return cur.fetchall()

def close_material_request(req_id):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE material_requests SET status='done' WHERE id=%s", (req_id,))
            conn.commit()

def get_all_admins():
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT uid FROM users WHERE role='admin'")
            rows = cur.fetchall()
            admin_uids = [r['uid'] for r in rows]
            for aid in ADMIN_IDS:
                if aid not in admin_uids:
                    admin_uids.append(aid)
            return admin_uids

def get_worker_stats(uid, days=7):
    since = datetime.now() - timedelta(days=days)
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT SUM(cable) as total_cable, SUM(devices) as total_devices, COUNT(*) as records
                FROM history WHERE uid=%s AND created_at>=%s
            """, (uid, since))
            return cur.fetchone()

def get_all_workers_stats(days=7):
    since = datetime.now() - timedelta(days=days)
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT u.name, u.uid,
                       COALESCE(SUM(h.cable),0) as total_cable,
                       COALESCE(SUM(h.devices),0) as total_devices,
                       COUNT(h.id) as records
                FROM users u
                LEFT JOIN history h ON h.uid=u.uid AND h.created_at>=%s
                WHERE u.role='worker'
                GROUP BY u.uid, u.name
                ORDER BY total_cable DESC
            """, (since,))
            return cur.fetchall()

# ─── Helpers ─────────────────────────────────────────────────────────────────

def fmt(dt):
    if isinstance(dt, datetime): return dt.strftime("%d.%m.%Y %H:%M")
    return str(dt)

def fmt_short(dt):
    if isinstance(dt, datetime): return dt.strftime("%d.%m")
    return str(dt)

def st_emoji(s): return {"active":"🟢","paused":"🟡","problem":"🔴","done":"✅"}.get(s,"⚪")
def st_label(s): return {"active":"В работе","paused":"Приостановлен","problem":"Проблема","done":"Завершён"}.get(s,s)

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

def clean(text):
    return "\n".join(l for l in text.split("\n") if not l.startswith("МЕТА:")).strip()

def progress_bar(done, total, length=10):
    if total <= 0: return "нет плана"
    pct = min(done / total, 1.0)
    filled = int(pct * length)
    bar = "█" * filled + "░" * (length - filled)
    return f"{bar} {int(pct*100)}%"

# ─── Keyboards ───────────────────────────────────────────────────────────────

def kb_admin():
    return ReplyKeyboardMarkup([
        ["📂 Объекты",       "➕ Новый объект"],
        ["👷 Сотрудники",    "🔗 Пригласить"],
        ["⚠️ Проблемы",      "📊 Статус"],
        ["📄 PDF отчёт",     "📋 Excel отчёт"],
        ["📈 Статистика",    "🗓 Дайджест"],
        ["📝 Задачи",        "📦 Заявки на материал"],
    ], resize_keyboard=True)

def kb_worker():
    return ReplyKeyboardMarkup([
        ["📋 Внести работы",    "📷 Отправить фото"],
        ["📂 Мои объекты",      "⚠️ Проблема"],
        ["✅ Мой чек-лист",     "📈 Моя статистика"],
        ["📦 Запросить материал","📊 Статус объекта"],
        ["📜 История",          "📍 Геолокация"],
    ], resize_keyboard=True)

def kb_objects_inline(uid):
    objs = get_objects(uid)
    buttons = [[InlineKeyboardButton(f"{st_emoji(o['status'])} {o['name']}", callback_data=f"sel:{o['id']}")] for o in objs]
    return InlineKeyboardMarkup(buttons)

# ─── Groq AI ─────────────────────────────────────────────────────────────────

async def ask_ai(uid, obj, user_message):
    probs = get_problems(obj['id'], status='open')
    hist  = get_history(obj['id'], limit=5)
    context = (
        f"Объект: «{obj['name']}» | {obj.get('address') or '—'}\n"
        f"Заказчик: {obj.get('customer') or '—'} | Инженер: {obj.get('engineer') or '—'}\n"
        f"Статус: {st_label(obj['status'])} | Кабель: {obj['total_cable']}м | Устройства: {obj['total_devices']}шт\n\n"
        "Последние записи:\n" + ("\n".join(f"• {fmt(h['created_at'])}: {h['text']}" for h in hist) or "нет") +
        "\nОткрытые проблемы:\n" + ("\n".join(f"• {p['title']}" for p in probs) or "нет")
    )
    ch = get_chat_hist(uid, obj['id'], limit=10)
    messages = [{"role":"system","content":SYSTEM_PROMPT},
                {"role":"user","content":context},
                {"role":"assistant","content":"Понял. Готов к работе."}]
    messages += [{"role":m['role'],"content":m['content']} for m in ch]
    messages.append({"role":"user","content":user_message})

    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(GROQ_URL,
            headers={"Authorization":f"Bearer {GROQ_API_KEY}","Content-Type":"application/json"},
            json={"model":"llama-3.3-70b-versatile","messages":messages,"max_tokens":1024,"temperature":0.3})
        resp.raise_for_status()
        data = resp.json()

    reply = data["choices"][0]["message"]["content"]
    add_chat(uid, obj['id'], "user", user_message)
    add_chat(uid, obj['id'], "assistant", reply)

    meta = parse_meta(reply)
    result = clean(reply)

    if "📅" in reply or "✔" in reply:
        add_history(obj['id'], uid, user_message, meta['cable'], meta['devices'])
        update_object(obj['id'],
            total_cable=obj['total_cable']+meta['cable'],
            total_devices=obj['total_devices']+meta['devices'])
        if meta['has_problem']:
            add_problem(obj['id'], f"Проблема от {fmt(datetime.now())}")

    return result, meta

# ─── Voice transcription ─────────────────────────────────────────────────────

async def transcribe_voice(file_path: str) -> str:
    async with httpx.AsyncClient(timeout=60) as client:
        with open(file_path, 'rb') as f:
            resp = await client.post(
                GROQ_AUDIO_URL,
                headers={"Authorization": f"Bearer {GROQ_API_KEY}"},
                files={"file": ("audio.ogg", f, "audio/ogg")},
                data={"model": "whisper-large-v3-turbo", "language": "ru"}
            )
            resp.raise_for_status()
            return resp.json().get("text", "")

# ─── Daily digest ─────────────────────────────────────────────────────────────

async def send_digest(app):
    """Отправляет ежедневный дайджест всем администраторам"""
    while True:
        now = datetime.utcnow() + timedelta(hours=TIMEZONE_OFFSET)
        # Следующий запуск в DIGEST_HOUR:00
        target = now.replace(hour=DIGEST_HOUR, minute=0, second=0, microsecond=0)
        if now >= target:
            target += timedelta(days=1)
        wait_seconds = (target - now).total_seconds()
        logger.info(f"Дайджест через {wait_seconds/3600:.1f} часов")
        await asyncio.sleep(wait_seconds)

        try:
            await build_and_send_digest(app)
        except Exception as e:
            logger.error(f"Digest error: {e}")

async def build_and_send_digest(app):
    yesterday = datetime.now() - timedelta(days=1)
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM objects WHERE status != 'done'")
            objects = cur.fetchall()

    if not objects:
        return

    lines = [f"🗓 *Утренний дайджест — {datetime.now().strftime('%d.%m.%Y')}*\n"]

    # Работы за вчера
    lines.append("📋 *Работы за вчера:*")
    has_work = False
    for obj in objects:
        hist = get_history_since(obj['id'], yesterday)
        if hist:
            has_work = True
            total_cable   = sum(h['cable'] for h in hist)
            total_devices = sum(h['devices'] for h in hist)
            lines.append(f"\n{st_emoji(obj['status'])} *{obj['name']}*")
            lines.append(f"  Записей: {len(hist)}")
            if total_cable:   lines.append(f"  Кабель: {total_cable} м")
            if total_devices: lines.append(f"  Устройства: {total_devices} шт")
    if not has_work:
        lines.append("  Вчера записей не было")

    # Открытые проблемы
    lines.append("\n⚠️ *Открытые проблемы:*")
    has_problems = False
    for obj in objects:
        probs = get_problems(obj['id'], status='open')
        if probs:
            has_problems = True
            lines.append(f"\n🔴 *{obj['name']}* — {len(probs)} проблем(ы)")
            for p in probs[:3]:
                lines.append(f"  • {p['title'][:50]}")
    if not has_problems:
        lines.append("  Открытых проблем нет ✅")

    # Статус объектов
    lines.append("\n📂 *Активные объекты:*")
    for obj in objects:
        if obj['status'] == 'active':
            progress = ""
            if obj.get('plan_cable', 0) > 0:
                progress = f" | {progress_bar(obj['total_cable'], obj['plan_cable'])}"
            lines.append(f"  🟢 {obj['name']}{progress}")

    text = "\n".join(lines)
    admin_uids = get_all_admins()
    for admin_uid in admin_uids:
        try:
            await app.bot.send_message(admin_uid, text, parse_mode="Markdown")
        except Exception as e:
            logger.error(f"Digest send error for {admin_uid}: {e}")

# ─── Export ───────────────────────────────────────────────────────────────────

def gen_pdf(obj):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    normal = ParagraphStyle('n', parent=styles['Normal'], fontSize=10, spaceAfter=3)
    h2     = ParagraphStyle('h2', parent=styles['Heading2'], fontSize=12, spaceAfter=4)
    story  = [Paragraph(f"Отчёт: {obj['name']}", styles['Title']),
              Paragraph(f"Сформирован: {fmt(datetime.now())}", normal), Spacer(1,12)]

    # Прогресс
    if obj.get('plan_cable', 0) > 0:
        pct_c = min(obj['total_cable']/obj['plan_cable']*100, 100)
        story.append(Paragraph(f"Прогресс кабеля: {obj['total_cable']}/{obj['plan_cable']} м ({pct_c:.0f}%)", normal))
    if obj.get('plan_devices', 0) > 0:
        pct_d = min(obj['total_devices']/obj['plan_devices']*100, 100)
        story.append(Paragraph(f"Прогресс устройств: {obj['total_devices']}/{obj['plan_devices']} шт ({pct_d:.0f}%)", normal))
    story.append(Spacer(1,8))

    info = [["Адрес",obj.get('address') or '—'],["Заказчик",obj.get('customer') or '—'],
            ["Инженер",obj.get('engineer') or '—'],["Статус",st_label(obj['status'])],
            ["Кабель",f"{obj['total_cable']} м"],["Устройства",f"{obj['total_devices']} шт"]]
    t = Table(info, colWidths=[150,330])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(0,-1),colors.HexColor('#f0f4ff')),
                            ('GRID',(0,0),(-1,-1),0.5,colors.grey),('PADDING',(0,0),(-1,-1),6)]))
    story += [t, Spacer(1,12), Paragraph("История работ", h2)]
    for h in get_history(obj['id'], limit=100):
        story.append(Paragraph(f"<b>{fmt(h['created_at'])}</b> — {h['text']}", normal))
    story += [Spacer(1,12), Paragraph("Проблемы", h2)]
    icons = {"open":"🔴","inwork":"🟡","solved":"🟢"}
    names = {"open":"Открыта","inwork":"В работе","solved":"Решена"}
    for p in get_problems(obj['id']):
        story.append(Paragraph(f"{icons.get(p['status'],'?')} {p['title']} — {names.get(p['status'])} ({fmt(p['created_at'])})", normal))
    doc.build(story); buf.seek(0)
    return buf

def gen_excel(obj):
    wb = openpyxl.Workbook()
    ws1 = wb.active; ws1.title = "Объект"
    ws1.column_dimensions['A'].width = 25; ws1.column_dimensions['B'].width = 45
    hf = PatternFill("solid", fgColor="1a3a6e"); bw = Font(bold=True, color="FFFFFF")
    for row in [["Объект",obj['name']],["Дата",fmt(datetime.now())],[],
                ["Адрес",obj.get('address') or '—'],["Заказчик",obj.get('customer') or '—'],
                ["Инженер",obj.get('engineer') or '—'],["Статус",st_label(obj['status'])],
                ["Кабель (факт)",f"{obj['total_cable']} м"],["Кабель (план)",f"{obj.get('plan_cable',0)} м"],
                ["Устройства (факт)",f"{obj['total_devices']} шт"],["Устройства (план)",f"{obj.get('plan_devices',0)} шт"]]:
        ws1.append(row)
        if row: ws1.cell(ws1.max_row,1).font = Font(bold=True)
    ws2 = wb.create_sheet("История")
    ws2.column_dimensions['A'].width=20; ws2.column_dimensions['B'].width=60
    ws2.column_dimensions['C'].width=10; ws2.column_dimensions['D'].width=10
    ws2.append(["Дата","Запись","Кабель(м)","Устройств"])
    for c in ['A1','B1','C1','D1']: ws2[c].fill=hf; ws2[c].font=bw
    for h in get_history(obj['id'], limit=100):
        ws2.append([fmt(h['created_at']), h['text'], h.get('cable',0), h.get('devices',0)])
    ws3 = wb.create_sheet("Проблемы")
    ws3.column_dimensions['A'].width=50; ws3.column_dimensions['B'].width=15; ws3.column_dimensions['C'].width=20
    ws3.append(["Проблема","Статус","Дата"])
    for c in ['A1','B1','C1']: ws3[c].fill=hf; ws3[c].font=bw
    names = {"open":"Открыта","inwork":"В работе","solved":"Решена"}
    for p in get_problems(obj['id']):
        ws3.append([p['title'],names.get(p['status']),fmt(p['created_at'])])
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─── Handlers ────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid  = update.effective_user.id
    name = update.effective_user.first_name or "коллега"
    args = ctx.args

    if args and args[0].startswith("invite_"):
        token = args[0][7:]
        if use_invite(token, uid):
            upsert_user(uid, name, role='worker')
            await update.message.reply_text(
                f"👋 Добро пожаловать, {name}!\n\nВы зарегистрированы как монтажник.\nАдминистратор назначит вам объекты.",
                reply_markup=kb_worker()
            )
        else:
            await update.message.reply_text("❌ Ссылка недействительна или уже использована.")
        return

    if is_registered(uid):
        kb = kb_admin() if is_admin(uid) else kb_worker()
        role = "администратор" if is_admin(uid) else "монтажник"
        await update.message.reply_text(f"👋 С возвращением, {name}! Вы — *{role}*.", parse_mode="Markdown", reply_markup=kb)
        return

    if uid in ADMIN_IDS:
        upsert_user(uid, name, role='admin')
        await update.message.reply_text(f"👋 Привет, {name}! Вы — администратор.", reply_markup=kb_admin())
    else:
        await update.message.reply_text("⛔ Доступ только по приглашению от администратора.")

async def cmd_invite(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not is_admin(uid):
        await update.message.reply_text("⛔ Только для администраторов.")
        return
    token = create_invite(uid)
    bot_user = BOT_USERNAME or (await ctx.bot.get_me()).username
    link = f"https://t.me/{bot_user}?start=invite_{token}"
    await update.message.reply_text(
        f"🔗 *Ссылка-приглашение:*\n\n`{link}`\n\n⚠️ Одноразовая. Перешлите монтажнику.",
        parse_mode="Markdown"
    )

async def cmd_myid(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Ваш ID: `{update.effective_user.id}`", parse_mode="Markdown")

# ─── Show helpers ─────────────────────────────────────────────────────────────

async def show_status(update, uid):
    obj = get_current_obj(uid)
    if not obj:
        await update.message.reply_text("Объект не выбран.")
        return
    probs = get_problems(obj['id'], status='open')
    hist  = get_history(obj['id'], limit=1)
    last  = f"\n📝 {fmt(hist[0]['created_at'])}" if hist else ""

    # Прогресс
    progress = ""
    if obj.get('plan_cable', 0) > 0:
        progress += f"\n🔌 Кабель: {obj['total_cable']}/{obj['plan_cable']} м {progress_bar(obj['total_cable'], obj['plan_cable'])}"
    else:
        progress += f"\n🔌 Кабель: {obj['total_cable']} м"
    if obj.get('plan_devices', 0) > 0:
        progress += f"\n📡 Устройства: {obj['total_devices']}/{obj['plan_devices']} шт {progress_bar(obj['total_devices'], obj['plan_devices'])}"
    else:
        progress += f"\n📡 Устройства: {obj['total_devices']} шт"

    text = (
        f"{st_emoji(obj['status'])} *{obj['name']}*\n"
        f"📍 {obj.get('address') or '—'}\n"
        f"👤 {obj.get('engineer') or '—'}\n"
        f"🏢 {obj.get('customer') or '—'}\n\n"
        f"Статус: *{st_label(obj['status'])}*"
        f"{progress}\n"
        f"⚠️ Открытых проблем: {len(probs)}{last}"
    )
    buttons = [
        [InlineKeyboardButton("🟢 В работе", callback_data="st:active"),
         InlineKeyboardButton("🟡 Пауза",    callback_data="st:paused")],
        [InlineKeyboardButton("🔴 Проблема", callback_data="st:problem"),
         InlineKeyboardButton("✅ Завершён",  callback_data="st:done")],
        [InlineKeyboardButton("📜 История",  callback_data="history"),
         InlineKeyboardButton("⚠️ Проблемы", callback_data="problems")],
    ]
    if is_admin(uid):
        buttons.append([
            InlineKeyboardButton("📄 PDF",       callback_data="exp:pdf"),
            InlineKeyboardButton("📊 Excel",     callback_data="exp:excel"),
            InlineKeyboardButton("👷 Назначить", callback_data="assign"),
        ])
        buttons.append([InlineKeyboardButton("🎯 Установить план", callback_data="set_plan")])
    await update.message.reply_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(buttons))

async def show_history(msg, uid):
    obj = get_current_obj(uid)
    if not obj:
        await msg.reply_text("Объект не выбран.")
        return
    hist = get_history(obj['id'], limit=10)
    if not hist:
        await msg.reply_text(f"📜 История по «{obj['name']}» пуста.")
        return
    lines = [f"📜 *История: {obj['name']}*\n"]
    for h in hist:
        lines.append(f"📅 {fmt(h['created_at'])}\n{h['text']}\n")
    await msg.reply_text("\n".join(lines), parse_mode="Markdown")

async def show_problems(msg, uid):
    obj = get_current_obj(uid)
    if not obj:
        await msg.reply_text("Объект не выбран.")
        return
    probs = get_problems(obj['id'])
    if not probs:
        await msg.reply_text(f"✅ Проблем нет.")
        return
    icons = {"open":"🔴","inwork":"🟡","solved":"🟢"}
    names = {"open":"Открыта","inwork":"В работе","solved":"Решена"}
    lines = [f"⚠️ *Проблемы: {obj['name']}*\n"]
    buttons = []
    for p in probs:
        lines.append(f"{icons.get(p['status'],'⚪')} {p['title']}\n   {names.get(p['status'])} · {fmt(p['created_at'])}\n")
        if p['status'] != 'solved':
            buttons.append([InlineKeyboardButton(f"✅ {p['title'][:35]}", callback_data=f"solve:{p['id']}")])
    await msg.reply_text("\n".join(lines), parse_mode="Markdown",
                         reply_markup=InlineKeyboardMarkup(buttons) if buttons else None)

async def show_stats(update, uid):
    if not is_admin(uid):
        stats = get_worker_stats(uid, days=7)
        await update.message.reply_text(
            f"📈 *Ваша статистика за 7 дней:*\n\n"
            f"🔌 Кабель: {stats['total_cable'] or 0} м\n"
            f"📡 Устройства: {stats['total_devices'] or 0} шт\n"
            f"📋 Записей: {stats['records'] or 0}",
            parse_mode="Markdown"
        )
        return

    workers = get_all_workers_stats(days=7)
    if not workers:
        await update.message.reply_text("Монтажников ещё нет.")
        return
    lines = ["📈 *Статистика монтажников за 7 дней:*\n"]
    for i, w in enumerate(workers, 1):
        medal = ["🥇","🥈","🥉"].pop(0) if i <= 3 else f"{i}."
        lines.append(
            f"{medal} *{w['name']}*\n"
            f"   🔌 {w['total_cable']} м | 📡 {w['total_devices']} шт | 📋 {w['records']} записей\n"
        )
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

# ─── Photo handler ────────────────────────────────────────────────────────────

async def handle_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not is_registered(uid):
        await update.message.reply_text("⛔ Нет доступа.")
        return
    obj = get_current_obj(uid)
    if not obj:
        await update.message.reply_text("Сначала выберите объект.")
        return
    photo   = update.message.photo[-1]
    caption = update.message.caption or ""
    add_photo(obj['id'], uid, photo.file_id, caption)
    await update.message.reply_text(
        f"📷 Фото сохранено!\n🏗 *{obj['name']}*\n📝 {caption or 'без комментария'}",
        parse_mode="Markdown"
    )

# ─── Voice handler ────────────────────────────────────────────────────────────

async def handle_voice(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not is_registered(uid):
        await update.message.reply_text("⛔ Нет доступа.")
        return
    obj = get_current_obj(uid)
    if not obj:
        await update.message.reply_text("Сначала выберите объект.")
        return

    await update.message.reply_text("🎤 Распознаю речь...")
    try:
        voice_file = await update.message.voice.get_file()
        with tempfile.NamedTemporaryFile(suffix='.ogg', delete=False) as tmp:
            await voice_file.download_to_drive(tmp.name)
            text = await transcribe_voice(tmp.name)
            os.unlink(tmp.name)

        if not text:
            await update.message.reply_text("Не удалось распознать речь. Попробуйте ещё раз.")
            return

        await update.message.reply_text(f"📝 Распознано:\n_{text}_", parse_mode="Markdown")
        await update.message.chat.send_action("typing")
        reply, _ = await ask_ai(uid, obj, text)
        await update.message.reply_text(reply)

    except Exception as e:
        logger.error(f"Voice error: {e}")
        await update.message.reply_text("⚠️ Ошибка распознавания голоса.")

# ─── Location handler ─────────────────────────────────────────────────────────

async def handle_location(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not is_registered(uid):
        return
    obj = get_current_obj(uid)
    if not obj:
        await update.message.reply_text("Сначала выберите объект.")
        return
    loc = update.message.location
    add_location(obj['id'], uid, loc.latitude, loc.longitude)
    user = get_db_user(uid)
    name = user['name'] if user else "Монтажник"
    maps_link = f"https://maps.google.com/?q={loc.latitude},{loc.longitude}"
    await update.message.reply_text(
        f"📍 Геолокация зафиксирована!\n🏗 *{obj['name']}*\n👤 {name}\n📅 {fmt(datetime.now())}\n[Открыть на карте]({maps_link})",
        parse_mode="Markdown"
    )
    # Уведомить администраторов
    for admin_uid in get_all_admins():
        if admin_uid != uid:
            try:
                await ctx.bot.send_message(
                    admin_uid,
                    f"📍 *{name}* отметился на объекте *{obj['name']}*\n[Открыть на карте]({maps_link})",
                    parse_mode="Markdown"
                )
            except: pass

# ─── Message handler ──────────────────────────────────────────────────────────

async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid  = update.effective_user.id
    text = update.message.text.strip()

    if not is_registered(uid):
        await update.message.reply_text("⛔ Доступ только по приглашению.")
        return

    admin  = is_admin(uid)
    state  = get_state(uid)

    # ── Состояния ─────────────────────────────────────────────────────────────
    if state == "new_obj_name":
        obj_id = create_object(uid, text)
        ctx.user_data["noid"] = obj_id
        set_state(uid, "new_obj_address")
        await update.message.reply_text(f"✅ «{text}» создан!\n\nВведите адрес (или /skip):")
        return

    if state == "new_obj_address":
        if text != "/skip": update_object(ctx.user_data.get("noid"), address=text)
        set_state(uid, "new_obj_customer")
        await update.message.reply_text("Заказчик (или /skip):")
        return

    if state == "new_obj_customer":
        if text != "/skip": update_object(ctx.user_data.get("noid"), customer=text)
        set_state(uid, "new_obj_engineer")
        await update.message.reply_text("Ответственный инженер (или /skip):")
        return

    if state == "new_obj_engineer":
        if text != "/skip": update_object(ctx.user_data.get("noid"), engineer=text)
        set_state(uid, "new_obj_plan_cable")
        await update.message.reply_text("Плановый объём кабеля в метрах (или /skip):")
        return

    if state == "new_obj_plan_cable":
        if text != "/skip":
            try: update_object(ctx.user_data.get("noid"), plan_cable=int(text))
            except: pass
        set_state(uid, "new_obj_plan_devices")
        await update.message.reply_text("Плановое количество устройств (или /skip):")
        return

    if state == "new_obj_plan_devices":
        if text != "/skip":
            try: update_object(ctx.user_data.get("noid"), plan_devices=int(text))
            except: pass
        set_state(uid, "")
        await update.message.reply_text("✅ Объект полностью настроен!", reply_markup=kb_admin())
        return

    if state == "material_request":
        obj = get_current_obj(uid)
        if obj:
            add_material_request(obj['id'], uid, text)
            set_state(uid, "")
            await update.message.reply_text(f"📦 Заявка отправлена администратору:\n«{text}»")
            # Уведомить администраторов
            user = get_db_user(uid)
            for admin_uid in get_all_admins():
                try:
                    await ctx.bot.send_message(
                        admin_uid,
                        f"📦 *Новая заявка на материал*\n"
                        f"👤 {user['name']} | 🏗 {obj['name']}\n\n"
                        f"«{text}»",
                        parse_mode="Markdown"
                    )
                except: pass
        return

    if state == "new_task_title":
        ctx.user_data["task_title"] = text
        set_state(uid, "new_task_worker")
        workers = get_workers()
        if not workers:
            await update.message.reply_text("Монтажников нет.")
            set_state(uid, "")
            return
        buttons = [[InlineKeyboardButton(w['name'], callback_data=f"task_worker:{w['uid']}")] for w in workers]
        await update.message.reply_text("Выберите монтажника:", reply_markup=InlineKeyboardMarkup(buttons))
        return


        obj = get_current_obj(uid)
        if obj:
            add_problem(obj['id'], text)
            set_state(uid, "")
            await update.message.reply_text(f"⚠️ Проблема зафиксирована:\n«{text}»")
        return

    if state == "set_plan_cable":
        obj = get_current_obj(uid)
        if obj:
            try: update_object(obj['id'], plan_cable=int(text))
            except: pass
            set_state(uid, "set_plan_devices")
            await update.message.reply_text("Плановое количество устройств:")
        return

    if state == "set_plan_devices":
        obj = get_current_obj(uid)
        if obj:
            try: update_object(obj['id'], plan_devices=int(text))
            except: pass
            set_state(uid, "")
            await update.message.reply_text("✅ План установлен!")
        return

    # ── Кнопки меню ───────────────────────────────────────────────────────────
    if text in ("📂 Объекты", "📂 Мои объекты"):
        objs = get_objects(uid)
        if not objs:
            await update.message.reply_text("Объектов нет.")
            return
        await update.message.reply_text("📂 Выберите объект:", reply_markup=kb_objects_inline(uid))
        return

    if text == "➕ Новый объект":
        if not admin:
            await update.message.reply_text("⛔ Только администратор.")
            return
        set_state(uid, "new_obj_name")
        await update.message.reply_text("📋 Введите название объекта:")
        return

    if text in ("📊 Статус", "📊 Статус объекта"):
        await show_status(update, uid)
        return

    if text == "📜 История":
        await show_history(update.message, uid)
        return

    if text in ("⚠️ Проблемы", "⚠️ Проблема"):
        obj = get_current_obj(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.")
            return
        if admin:
            await show_problems(update.message, uid)
        else:
            set_state(uid, "report_problem")
            await update.message.reply_text("Опишите проблему:")
        return

    if text == "🔗 Пригласить":
        await cmd_invite(update, ctx)
        return

    if text == "👷 Сотрудники":
        workers = get_workers()
        if not workers:
            await update.message.reply_text("Монтажников нет. Пригласите через 🔗 Пригласить")
            return
        lines = ["👷 *Монтажники:*\n"]
        for w in workers:
            stats = get_worker_stats(w['uid'], days=7)
            lines.append(f"• {w['name']} — {stats['total_cable'] or 0}м / {stats['total_devices'] or 0}шт за 7 дней")
        await update.message.reply_text("\n".join(lines), parse_mode="Markdown")
        return

    if text == "📈 Статистика":
        await show_stats(update, uid)
        return

    if text == "🗓 Дайджест":
        if not admin:
            return
        await update.message.reply_text("⏳ Формирую дайджест...")
        await build_and_send_digest(ctx.application)
        return

    if text == "📄 PDF отчёт":
        obj = get_current_obj(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.")
            return
        await update.message.reply_text("⏳ Генерирую PDF...")
        buf = gen_pdf(obj)
        await update.message.reply_document(buf, filename=f"report_{obj['name']}.pdf")
        return

    if text == "📋 Excel отчёт":
        obj = get_current_obj(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.")
            return
        await update.message.reply_text("⏳ Генерирую Excel...")
        buf = gen_excel(obj)
        await update.message.reply_document(buf, filename=f"report_{obj['name']}.xlsx")
        return

    if text == "📋 Внести работы":
        obj = get_current_obj(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.", reply_markup=kb_objects_inline(uid))
            return
        await update.message.reply_text(f"✍️ Пишите что сделано на *{obj['name']}*:", parse_mode="Markdown")
        return

    if text == "📷 Отправить фото":
        obj = get_current_obj(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.")
            return
        await update.message.reply_text(f"📷 Отправьте фото — привяжется к *{obj['name']}*", parse_mode="Markdown")
        return

    if text == "✅ Мой чек-лист":
        tasks = get_tasks(uid, status='open')
        if not tasks:
            await update.message.reply_text(
                "✅ На сегодня задач нет.\n\nАдминистратор назначит задачи через кнопку 📝 Задачи."
            )
            return
        lines = ["✅ *Ваши задачи:*\n"]
        buttons = []
        for t in tasks:
            lines.append(f"🔲 {t['title']}\n   📍 {t['obj_name']}")
            buttons.append([InlineKeyboardButton(f"✅ Выполнено: {t['title'][:35]}", callback_data=f"done_task:{t['id']}")])
        await update.message.reply_text("\n".join(lines), parse_mode="Markdown",
                                        reply_markup=InlineKeyboardMarkup(buttons))
        return

    if text == "📈 Моя статистика":
        stats_7  = get_worker_stats(uid, days=7)
        stats_30 = get_worker_stats(uid, days=30)
        tasks_done = len([t for t in get_tasks(uid) if t['status'] == 'done'])
        tasks_open = len([t for t in get_tasks(uid) if t['status'] == 'open'])
        await update.message.reply_text(
            f"📈 *Моя статистика*\n\n"
            f"*За 7 дней:*\n"
            f"🔌 Кабель: {stats_7['total_cable'] or 0} м\n"
            f"📡 Устройства: {stats_7['total_devices'] or 0} шт\n"
            f"📋 Записей: {stats_7['records'] or 0}\n\n"
            f"*За 30 дней:*\n"
            f"🔌 Кабель: {stats_30['total_cable'] or 0} м\n"
            f"📡 Устройства: {stats_30['total_devices'] or 0} шт\n"
            f"📋 Записей: {stats_30['records'] or 0}\n\n"
            f"*Задачи:*\n"
            f"✅ Выполнено: {tasks_done}\n"
            f"🔲 Открытых: {tasks_open}",
            parse_mode="Markdown"
        )
        return

    if text == "📦 Запросить материал":
        obj = get_current_obj(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.")
            return
        set_state(uid, "material_request")
        await update.message.reply_text(
            f"📦 Введите что нужно заказать для объекта *{obj['name']}*:\n\n"
            f"Например: «Кабель КПСнг 500м, дымовые датчики ИП212 — 20шт»",
            parse_mode="Markdown"
        )
        return

    if text == "📝 Задачи":
        if not admin: return
        workers = get_workers()
        if not workers:
            await update.message.reply_text("Монтажников нет. Сначала пригласите через 🔗 Пригласить")
            return
        # Показать все открытые задачи + кнопка создать новую
        all_tasks = get_all_tasks_for_admin()
        open_tasks = [t for t in all_tasks if t['status'] == 'open']
        lines = ["📝 *Активные задачи:*\n"]
        if open_tasks:
            for t in open_tasks:
                lines.append(f"🔲 *{t['worker_name']}* — {t['title']}\n   📍 {t['obj_name']}\n")
        else:
            lines.append("Открытых задач нет.\n")
        buttons = [[InlineKeyboardButton("➕ Создать задачу", callback_data="new_task")]]
        await update.message.reply_text("\n".join(lines), parse_mode="Markdown",
                                        reply_markup=InlineKeyboardMarkup(buttons))
        return

    if text == "📦 Заявки на материал":
        if not admin: return
        requests = get_material_requests(status='new')
        if not requests:
            await update.message.reply_text("📦 Новых заявок на материал нет.")
            return
        lines = ["📦 *Заявки на материал:*\n"]
        buttons = []
        for r in requests:
            lines.append(f"🔸 *{r['worker_name']}* | {r['obj_name']}\n   {r['text']}\n   📅 {fmt(r['created_at'])}\n")
            buttons.append([InlineKeyboardButton(f"✅ Закрыть: {r['text'][:35]}", callback_data=f"close_req:{r['id']}")])
        await update.message.reply_text("\n".join(lines), parse_mode="Markdown",
                                        reply_markup=InlineKeyboardMarkup(buttons))
        return

    if text == "📍 Геолокация":
        obj = get_current_obj(uid)
        if not obj:
            await update.message.reply_text("Сначала выберите объект.")
            return
        kb = ReplyKeyboardMarkup(
            [[KeyboardButton("📍 Отправить геолокацию", request_location=True)]],
            resize_keyboard=True, one_time_keyboard=True
        )
        await update.message.reply_text("Нажмите кнопку:", reply_markup=kb)
        return

    if text == "🎤 Голосовое":
        await update.message.reply_text("🎤 Просто отправьте голосовое сообщение.")
        return

    # ── AI ────────────────────────────────────────────────────────────────────
    obj = get_current_obj(uid)
    if not obj:
        await update.message.reply_text("Выберите объект:", reply_markup=kb_objects_inline(uid))
        return

    await update.message.chat.send_action("typing")
    try:
        reply, _ = await ask_ai(uid, obj, text)
        await update.message.reply_text(reply)
    except Exception as e:
        logger.error(f"AI error: {e}")
        await update.message.reply_text("⚠️ Ошибка AI. Попробуйте позже.")

# ─── Callback handler ────────────────────────────────────────────────────────

async def handle_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid  = q.from_user.id
    data = q.data
    admin = is_admin(uid)

    if data.startswith("sel:"):
        obj_id = int(data[4:])
        set_current_obj(uid, obj_id)
        obj = get_object(obj_id)
        if obj:
            await q.message.reply_text(
                f"✅ Выбран: *{obj['name']}*\n{st_emoji(obj['status'])} {st_label(obj['status'])}",
                parse_mode="Markdown",
                reply_markup=kb_admin() if admin else kb_worker()
            )

    elif data.startswith("st:"):
        obj = get_current_obj(uid)
        if obj:
            update_object(obj['id'], status=data[3:])
            await q.message.reply_text(f"✅ Статус: {st_emoji(data[3:])} {st_label(data[3:])}")

    elif data.startswith("solve:"):
        solve_problem(int(data[6:]))
        await q.message.reply_text("✅ Проблема закрыта.")

    elif data == "history":
        await show_history(q.message, uid)

    elif data == "problems":
        await show_problems(q.message, uid)

    elif data.startswith("exp:"):
        obj = get_current_obj(uid)
        if not obj or not admin: return
        await q.message.reply_text("⏳ Генерирую...")
        if data[4:] == "pdf":
            buf = gen_pdf(obj)
            await q.message.reply_document(buf, filename=f"report_{obj['name']}.pdf")
        else:
            buf = gen_excel(obj)
            await q.message.reply_document(buf, filename=f"report_{obj['name']}.xlsx")

    elif data == "assign":
        if not admin: return
        obj = get_current_obj(uid)
        if not obj:
            await q.message.reply_text(
                "⚠️ Объект не выбран.\n\n"
                "Сначала выберите объект: нажмите 📂 Объекты → выберите нужный → затем 📊 Статус → 👷 Назначить"
            )
            return
        workers = get_workers()
        if not workers:
            await q.message.reply_text(
                "👷 Монтажников пока нет.\n\n"
                "Пригласите через кнопку 🔗 Пригласить — после регистрации они появятся здесь."
            )
            return
        lines = [f"👷 Назначить на *{obj['name']}*:\n"]
        for w in workers:
            stats = get_worker_stats(w['uid'], days=7)
            lines.append(f"• {w['name']} — {stats['total_cable'] or 0}м за 7 дней")
        buttons = [[InlineKeyboardButton(w['name'], callback_data=f"asgn:{obj['id']}:{w['uid']}")] for w in workers]
        await q.message.reply_text("\n".join(lines), parse_mode="Markdown",
                                   reply_markup=InlineKeyboardMarkup(buttons))

    elif data.startswith("asgn:"):
        _, obj_id, worker_uid = data.split(":")
        assign_worker(int(obj_id), int(worker_uid))
        obj    = get_object(int(obj_id))
        worker = get_db_user(int(worker_uid))
        await q.message.reply_text(f"✅ {worker['name']} назначен на «{obj['name']}»")
        try:
            await ctx.bot.send_message(int(worker_uid),
                f"📋 Вам назначен объект: *{obj['name']}*\n📍 {obj.get('address') or '—'}",
                parse_mode="Markdown")
        except: pass

    elif data == "new_task":
        if not admin: return
        set_state(uid, "new_task_title")
        await q.message.reply_text("📝 Введите текст задачи:")

    elif data.startswith("task_worker:"):
        worker_uid = int(data[12:])
        obj = get_current_obj(uid)
        title = ctx.user_data.get("task_title", "")
        if obj and title:
            task_id = create_task(obj['id'], worker_uid, uid, title)
            worker = get_db_user(worker_uid)
            set_state(uid, "")
            await q.message.reply_text(f"✅ Задача создана для {worker['name']}:\n«{title}»")
            try:
                await ctx.bot.send_message(
                    worker_uid,
                    f"📋 *Новая задача от администратора:*\n\n"
                    f"🔲 {title}\n📍 {obj['name']}\n\n"
                    f"Нажмите ✅ Мой чек-лист чтобы увидеть все задачи.",
                    parse_mode="Markdown"
                )
            except: pass

    elif data.startswith("done_task:"):
        task_id = int(data[10:])
        complete_task(task_id)
        await q.message.reply_text("✅ Задача выполнена! Молодец!")

    elif data.startswith("close_req:"):
        req_id = int(data[10:])
        close_material_request(req_id)
        await q.message.reply_text("✅ Заявка закрыта.")

    elif data == "set_plan":
        if not admin: return
        set_state(uid, "set_plan_cable")
        await q.message.reply_text("Введите плановый объём кабеля (в метрах):")

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    init_db()
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start",  cmd_start))
    app.add_handler(CommandHandler("invite", cmd_invite))
    app.add_handler(CommandHandler("myid",   cmd_myid))
    app.add_handler(CommandHandler("skip",   handle_message))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.PHOTO,    handle_photo))
    app.add_handler(MessageHandler(filters.VOICE,    handle_voice))
    app.add_handler(MessageHandler(filters.LOCATION, handle_location))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    loop = asyncio.get_event_loop()
    loop.create_task(send_digest(app))

    logger.info("Бот v4 запущен!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
