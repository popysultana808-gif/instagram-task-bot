import asyncio
import datetime
import io
import json
import logging
import os
import random
import re
import sqlite3
import uuid
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from aiohttp import web
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
    Update,
)
from telegram.constants import ParseMode
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
except ImportError:
    psycopg2 = None
    RealDictCursor = None

# ==============================================================================
# CONFIGURATION & LOGGING
# ==============================================================================
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
PORT = int(os.getenv("PORT", "10000"))
OWNER_ID = 8919985167
BD_TZ = datetime.timezone(datetime.timedelta(hours=6))

FIRST_NAMES = [
    "Scott", "James", "Robert", "John", "Michael", "David", "William", "Richard",
    "Joseph", "Thomas", "Charles", "Christopher", "Daniel", "Matthew", "Anthony",
    "Mark", "Donald", "Steven", "Paul", "Andrew", "Joshua", "Kenneth", "Kevin",
    "Brian", "George", "Edward", "Ronald", "Timothy", "Jason", "Jeffrey",
    "Ryan", "Jacob", "Gary", "Nicholas", "Eric", "Jonathan", "Stephen", "Larry",
    "Emily", "Emma", "Olivia", "Sophia", "Ava", "Isabella", "Mia", "Charlotte",
    "Amelia", "Harper", "Evelyn", "Abigail", "Ella", "Scarlett", "Grace", "Chloe",
]
LAST_NAMES = [
    "Campbell", "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia",
    "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez",
    "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin",
    "Lee", "Perez", "Thompson", "White", "Harris", "Sanchez", "Clark", "Ramirez",
    "Lewis", "Robinson", "Walker", "Young", "Allen", "King", "Wright", "Scott",
    "Torres", "Nguyen", "Hill", "Flores", "Green", "Adams", "Nelson", "Baker",
]

DEFAULT_MAIN_BUTTONS = [
    "Work",
    "File Submit",
    "Balance",
    "Withdraw",
    "My Referral",
    "Support",
    "আমি নতুন",
]

# ==============================================================================
# DATABASE LAYER (PostgreSQL with SQLite fallback)
# ==============================================================================
class Database:
    def __init__(self, db_url: str = ""):
        self.db_url = db_url
        self.use_pg = bool(db_url and psycopg2 is not None)
        self.sqlite_path = os.getenv("SQLITE_PATH", "bot_data.db")
        self._init_db()

    def _connect(self):
        if self.use_pg:
            return psycopg2.connect(self.db_url, sslmode="require")
        conn = sqlite3.connect(self.sqlite_path, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        return conn

    def _ph(self, query: str) -> str:
        return query.replace("?", "%s") if self.use_pg else query

    def execute(self, query: str, params: Tuple = ()) -> None:
        conn = self._connect()
        try:
            cur = conn.cursor()
            cur.execute(self._ph(query), params)
            conn.commit()
        finally:
            conn.close()

    def execute_insert_returning_id(self, query: str, params: Tuple = ()) -> int:
        conn = self._connect()
        try:
            cur = conn.cursor()
            if self.use_pg:
                cur.execute(self._ph(query) + " RETURNING id", params)
                row_id = cur.fetchone()[0]
            else:
                cur.execute(query, params)
                row_id = cur.lastrowid
            conn.commit()
            return row_id
        finally:
            conn.close()

    def fetchone(self, query: str, params: Tuple = ()) -> Optional[Dict[str, Any]]:
        conn = self._connect()
        try:
            if self.use_pg:
                cur = conn.cursor(cursor_factory=RealDictCursor)
                cur.execute(self._ph(query), params)
                row = cur.fetchone()
                return dict(row) if row else None
            else:
                cur = conn.cursor()
                cur.execute(query, params)
                row = cur.fetchone()
                return dict(row) if row else None
        finally:
            conn.close()

    def fetchall(self, query: str, params: Tuple = ()) -> List[Dict[str, Any]]:
        conn = self._connect()
        try:
            if self.use_pg:
                cur = conn.cursor(cursor_factory=RealDictCursor)
                cur.execute(self._ph(query), params)
                rows = cur.fetchall()
                return [dict(r) for r in rows]
            else:
                cur = conn.cursor()
                cur.execute(query, params)
                rows = cur.fetchall()
                return [dict(r) for r in rows]
        finally:
            conn.close()

    def _init_db(self) -> None:
        serial_pk = "SERIAL PRIMARY KEY" if self.use_pg else "INTEGER PRIMARY KEY AUTOINCREMENT"
        statements = [
            f"""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS admins (
                user_id BIGINT PRIMARY KEY,
                added_by BIGINT,
                created_at TEXT
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS users (
                user_id BIGINT PRIMARY KEY,
                username TEXT,
                full_name TEXT,
                balance REAL DEFAULT 0,
                pending_withdraw REAL DEFAULT 0,
                total_income REAL DEFAULT 0,
                completed_jobs INTEGER DEFAULT 0,
                rejected_jobs INTEGER DEFAULT 0,
                referred_by BIGINT,
                referral_paid INTEGER DEFAULT 0,
                referral_count INTEGER DEFAULT 0,
                referral_earnings REAL DEFAULT 0,
                joined_at TEXT
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS categories (
                id {serial_pk},
                name TEXT NOT NULL UNIQUE,
                is_active INTEGER DEFAULT 1
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS subcategories (
                id {serial_pk},
                category_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                rate REAL NOT NULL DEFAULT 5.0,
                steps_json TEXT NOT NULL,
                is_active INTEGER DEFAULT 1
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS used_usernames (
                username TEXT PRIMARY KEY,
                created_at TEXT
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS batches (
                id {serial_pk},
                user_id BIGINT NOT NULL,
                subcategory_id INTEGER NOT NULL,
                subcategory_name TEXT NOT NULL,
                rate REAL NOT NULL,
                total_count INTEGER NOT NULL,
                approved_count INTEGER DEFAULT 0,
                rejected_count INTEGER DEFAULT 0,
                status TEXT DEFAULT 'received',
                received_at TEXT,
                reported_at TEXT
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS submissions (
                id {serial_pk},
                user_id BIGINT NOT NULL,
                subcategory_id INTEGER NOT NULL,
                subcategory_name TEXT NOT NULL,
                rate REAL NOT NULL,
                gen_name TEXT,
                gen_username TEXT NOT NULL,
                gen_password TEXT NOT NULL,
                steps_data_json TEXT NOT NULL,
                token_fingerprint TEXT,
                status TEXT DEFAULT 'pending',
                batch_id INTEGER,
                created_at TEXT
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS withdrawals (
                id {serial_pk},
                user_id BIGINT NOT NULL,
                method TEXT NOT NULL,
                account_number TEXT NOT NULL,
                requested_amount REAL NOT NULL,
                fee REAL NOT NULL,
                net_amount REAL NOT NULL,
                status TEXT DEFAULT 'pending',
                created_at TEXT,
                processed_at TEXT
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS file_submit_categories (
                id {serial_pk},
                name TEXT NOT NULL UNIQUE,
                is_active INTEGER DEFAULT 1
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS file_submissions (
                id {serial_pk},
                user_id BIGINT NOT NULL,
                category_name TEXT NOT NULL,
                file_id TEXT NOT NULL,
                file_name TEXT,
                status TEXT DEFAULT 'pending',
                created_at TEXT
            )
            """,
            f"""
            CREATE TABLE IF NOT EXISTS guide_categories (
                id {serial_pk},
                name TEXT NOT NULL UNIQUE,
                guide_text TEXT,
                video_file_id TEXT,
                is_active INTEGER DEFAULT 1
            )
            """,
        ]
        for stmt in statements:
            self.execute(stmt)

        # Seed defaults
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        self.execute(
            "INSERT INTO admins (user_id, added_by, created_at) VALUES (?, ?, ?) ON CONFLICT DO NOTHING"
            if self.use_pg
            else "INSERT OR IGNORE INTO admins (user_id, added_by, created_at) VALUES (?, ?, ?)",
            (OWNER_ID, OWNER_ID, now_str),
        )

        defaults = {
            "bot_active": "1",
            "min_withdraw": "50",
            "withdraw_fee": "5",
            "referral_bonus": "3",
            "password_prefix": "succes#",
            "password_reset_time": "00:00",
            "support_link": "https://t.me/Talha_juba098",
            "extra_buttons_json": "[]",
        }
        for k, v in defaults.items():
            self.execute(
                "INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT DO NOTHING"
                if self.use_pg
                else "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
                (k, v),
            )

        # Seed Instagram category & USA Instagram subcategory if none exist
        existing_cat = self.fetchone("SELECT id FROM categories LIMIT 1")
        if not existing_cat:
            cat_id = self.execute_insert_returning_id(
                "INSERT INTO categories (name, is_active) VALUES (?, ?)",
                ("Instagram", 1),
            )
            default_steps = [
                {
                    "name": "Hotmail Set",
                    "prompt": "সম্পূর্ণ হটমেইল ফুল টোকেনটি এখানে দিন:",
                    "validator": "hotmail_token",
                }
            ]
            self.execute(
                "INSERT INTO subcategories (category_id, name, rate, steps_json, is_active) VALUES (?, ?, ?, ?, ?)",
                (cat_id, "USA Instagram", 5.0, json.dumps(default_steps, ensure_ascii=False), 1),
            )

        # Seed default Guide categories
        existing_guide = self.fetchone("SELECT id FROM guide_categories LIMIT 1")
        if not existing_guide:
            for gname, gtext in [
                ("Instagram", "📸 Instagram কাজের নিয়ম:\n১. Work -> Instagram -> USA Instagram এ ক্লিক করুন।\n২. বটের দেওয়া নাম, ইউজারনেম ও পাসওয়ার্ড কপি করে একাউন্ট খুলুন।\n৩. Hotmail Set এ ক্লিক করে ফুল টোকেন সাবমিট করুন।"),
                ("Gmail", "📧 Gmail কাজের গাইডলাইন এখনো অ্যাড করা হয়নি।"),
                ("Facebook", "📘 Facebook কাজের গাইডলাইন এখনো অ্যাড করা হয়নি।"),
            ]:
                self.execute(
                    "INSERT INTO guide_categories (name, guide_text, video_file_id, is_active) VALUES (?, ?, ?, ?)",
                    (gname, gtext, "", 1),
                )


db = Database(DATABASE_URL)

# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================
def get_setting(key: str, default: str = "") -> str:
    row = db.fetchone("SELECT value FROM settings WHERE key = ?", (key,))
    return row["value"] if row else default


def set_setting(key: str, value: str) -> None:
    if db.use_pg:
        db.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value",
            (key, value),
        )
    else:
        db.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
            (key, value),
        )


def is_admin(user_id: int) -> bool:
    if user_id == OWNER_ID:
        return True
    row = db.fetchone("SELECT user_id FROM admins WHERE user_id = ?", (user_id,))
    return row is not None


def get_all_admin_ids() -> List[int]:
    rows = db.fetchall("SELECT user_id FROM admins")
    ids = {r["user_id"] for r in rows}
    ids.add(OWNER_ID)
    return list(ids)


def get_current_password() -> str:
    prefix = get_setting("password_prefix", "succes#")
    reset_time_str = get_setting("password_reset_time", "00:00")
    now_bd = datetime.datetime.now(BD_TZ)
    try:
        rh, rm = [int(x) for x in reset_time_str.split(":")]
    except Exception:
        rh, rm = 0, 0
    reset_dt_today = now_bd.replace(hour=rh, minute=rm, second=0, microsecond=0)
    if now_bd < reset_dt_today:
        effective_date = now_bd - datetime.timedelta(days=1)
    else:
        effective_date = now_bd
    return f"{prefix}{effective_date.day}"


def generate_random_profile() -> Tuple[str, str]:
    for _ in range(30):
        fn = random.choice(FIRST_NAMES)
        ln = random.choice(LAST_NAMES)
        full_name = f"{fn} {ln}"
        suffix_num = random.randint(100, 9999)
        suffix_char = random.choice("abcdefghijklmnopqrstuvwxyz")
        uname = f"{fn.lower()}{ln.lower()}{suffix_num}{suffix_char}"
        exists = db.fetchone("SELECT username FROM used_usernames WHERE username = ?", (uname,))
        if not exists:
            now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
            db.execute(
                "INSERT INTO used_usernames (username, created_at) VALUES (?, ?)",
                (uname, now_str),
            )
            return full_name, uname
    fallback_uname = f"user{uuid.uuid4().hex[:10]}"
    return "Alex Carter", fallback_uname


UUID_RE = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
HOTMAIL_DOMAINS = ("@hotmail.com", "@outlook.com", "@live.com", "@msn.com")


def validate_step_input(validator_type: str, text: str) -> Tuple[bool, str, str]:
    val = text.strip()
    if validator_type == "hotmail_token":
        parts = val.split("|")
        if len(parts) < 4:
            return False, "❌ ভুল ফরম্যাট! হটমেইল ফুল টোকেনে অবশ্যই `email|password|refresh_token|client_id` থাকতে হবে।", ""
        email = parts[0].strip().lower()
        pwd = parts[1].strip()
        refresh_token = parts[2].strip()
        client_id = parts[3].strip()
        if not any(email.endswith(d) for d in HOTMAIL_DOMAINS) or "@" not in email:
            return False, "❌ ভ্যালিড Hotmail/Outlook ইমেইল পাওয়া যায়নি!", ""
        if len(pwd) < 3:
            return False, "❌ পাসওয়ার্ড অংশটি সঠিক নয়!", ""
        if not refresh_token.startswith("M.") or len(refresh_token) < 40:
            return False, "❌ Refresh Token সঠিক নয়! অনুগ্রহ করে সম্পূর্ণ ফুল টোকেন দিন।", ""
        if not UUID_RE.match(client_id):
            return False, "❌ Client ID (UUID) সঠিক নয়! সম্পূর্ণ ফুল টোকেনটি কপি করে দিন।", ""
        return True, "", email
    elif validator_type == "2fa":
        cleaned = val.replace(" ", "")
        if len(cleaned) < 8:
            return False, "❌ ভ্যালিড 2FA Key দিন (কমপক্ষে ৮ অক্ষরের)।", ""
        return True, "", ""
    else:
        if len(val) < 2:
            return False, "❌ অনুগ্রহ করে সঠিক তথ্য দিন।", ""
        return True, "", ""


def ensure_user(update: Update) -> Dict[str, Any]:
    u = update.effective_user
    row = db.fetchone("SELECT * FROM users WHERE user_id = ?", (u.id,))
    if not row:
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        db.execute(
            """
            INSERT INTO users (user_id, username, full_name, joined_at)
            VALUES (?, ?, ?, ?)
            """,
            (u.id, u.username or "", u.full_name or "", now_str),
        )
        row = db.fetchone("SELECT * FROM users WHERE user_id = ?", (u.id,))
    else:
        db.execute(
            "UPDATE users SET username = ?, full_name = ? WHERE user_id = ?",
            (u.username or "", u.full_name or "", u.id),
        )
    return row


def main_keyboard(user_id: int) -> ReplyKeyboardMarkup:
    extra_raw = get_setting("extra_buttons_json", "[]")
    try:
        extra_btns = json.loads(extra_raw)
    except Exception:
        extra_btns = []

    rows = [
        [KeyboardButton("Work"), KeyboardButton("File Submit")],
        [KeyboardButton("Balance"), KeyboardButton("Withdraw")],
        [KeyboardButton("My Referral"), KeyboardButton("Support")],
        [KeyboardButton("আমি নতুন")],
    ]
    for i in range(0, len(extra_btns), 2):
        pair = [KeyboardButton(b["label"]) for b in extra_btns[i : i + 2]]
        rows.append(pair)
    if is_admin(user_id):
        rows.append([KeyboardButton("⚙️ Admin Panel")])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


def admin_panel_keyboard() -> InlineKeyboardMarkup:
    bot_active = get_setting("bot_active", "1") == "1"
    status_label = "🟢 Bot Status: ON" if bot_active else "🔴 Bot Status: OFF"
    kb = [
        [InlineKeyboardButton(status_label, callback_data="adm:toggle_bot")],
        [
            InlineKeyboardButton("📂 Pending Files", callback_data="adm:pending_files"),
            InlineKeyboardButton("📤 Upload Report (.xlsx)", callback_data="adm:upload_report"),
        ],
        [
            InlineKeyboardButton("📩 User Inbox Upload", callback_data="adm:inbox_upload"),
            InlineKeyboardButton("🛠 Work & Steps", callback_data="adm:work_mgr"),
        ],
        [
            InlineKeyboardButton("💰 Balance & Withdraw", callback_data="adm:bal_mgr"),
            InlineKeyboardButton("🔑 Password Settings", callback_data="adm:pwd_mgr"),
        ],
        [
            InlineKeyboardButton("⌨️ Main Buttons", callback_data="adm:btn_mgr"),
            InlineKeyboardButton("🎬 আমি নতুন (Guides)", callback_data="adm:guide_mgr"),
        ],
        [
            InlineKeyboardButton("📁 File Submit Cats", callback_data="adm:fsub_mgr"),
            InlineKeyboardButton("👥 User Directory", callback_data="adm:user_dir:0"),
        ],
        [
            InlineKeyboardButton("📜 File History Report", callback_data="adm:file_history"),
            InlineKeyboardButton("🗑 Delete User Data", callback_data="adm:del_user"),
        ],
        [
            InlineKeyboardButton("👮 Manage Admins", callback_data="adm:admins"),
            InlineKeyboardButton("📢 Broadcast", callback_data="adm:broadcast"),
        ],
        [InlineKeyboardButton("🆘 Edit Support Link", callback_data="adm:edit_support")],
    ]
    return InlineKeyboardMarkup(kb)


def back_to_admin_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="adm:home")]])


def classify_excel_color(cell) -> Optional[str]:
    """Detect green (approved) or red (rejected) fill color from openpyxl cell."""
    fill = cell.fill
    if not fill or fill.fill_type in (None, "none"):
        return None
    fg = fill.fgColor
    if not fg:
        return None
    rgb = None
    if fg.type == "rgb" and fg.rgb:
        rgb = str(fg.rgb).upper()
    elif fg.type == "indexed" and fg.indexed is not None:
        # Common Excel indexed colors: 10=red, 11=green, 17=green, 29=red, 50=light green, 57=green
        if fg.indexed in (11, 17, 42, 50, 57):
            return "green"
        if fg.indexed in (10, 29, 45, 53):
            return "red"
    if not rgb or len(rgb) < 6:
        return None
    hex_str = rgb[-6:]
    if hex_str in ("000000", "FFFFFF"):
        return None
    try:
        r = int(hex_str[0:2], 16)
        g = int(hex_str[2:4], 16)
        b = int(hex_str[4:6], 16)
    except ValueError:
        return None
    if g > r + 25 and g > b + 20:
        return "green"
    if r > g + 30 and r > b + 30:
        return "red"
    return None


# ==============================================================================
# START & MAIN MENU HANDLERS
# ==============================================================================
async def start_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.clear()
    u = update.effective_user
    existing = db.fetchone("SELECT user_id FROM users WHERE user_id = ?", (u.id,))
    ensure_user(update)

    # Handle referral start param
    if not existing and context.args:
        ref_arg = context.args[0].strip()
        if ref_arg.isdigit():
            ref_id = int(ref_arg)
            if ref_id != u.id:
                ref_user = db.fetchone("SELECT user_id FROM users WHERE user_id = ?", (ref_id,))
                if ref_user:
                    db.execute("UPDATE users SET referred_by = ? WHERE user_id = ?", (ref_id, u.id))
                    db.execute(
                        "UPDATE users SET referral_count = referral_count + 1 WHERE user_id = ?",
                        (ref_id,),
                    )

    if get_setting("bot_active", "1") != "1" and not is_admin(u.id):
        await update.message.reply_text("🔧 বটের আপডেটের কাজ চলতেছে। অনুগ্রহ করে কিছুক্ষণ পর আবার চেষ্টা করুন।")
        return

    await update.message.reply_text(
        f"👋 স্বাগতম, <b>{u.first_name}</b>!\nনিচের মেনু থেকে আপনার পছন্দের অপশন বেছে নিন:",
        parse_mode=ParseMode.HTML,
        reply_markup=main_keyboard(u.id),
    )


async def admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_admin(update.effective_user.id):
        return
    context.user_data.clear()
    await update.message.reply_text(
        "⚙️ <b>Admin Control Panel</b>\nনিচ থেকে যেকোনো অপশন সিলেক্ট করুন:",
        parse_mode=ParseMode.HTML,
        reply_markup=admin_panel_keyboard(),
    )


# ==============================================================================
# WORK & SUBMISSION FLOW
# ==============================================================================
async def handle_work_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    cats = db.fetchall("SELECT * FROM categories WHERE is_active = 1 ORDER BY id ASC")
    if not cats:
        await update.message.reply_text("📋 এই মুহূর্তে কোনো কাজের ক্যাটাগরি চালু নেই।")
        return
    kb = [[InlineKeyboardButton(f"{c['name']} ▸", callback_data=f"work:cat:{c['id']}")] for c in cats]
    await update.message.reply_text("📋 Select work type:", reply_markup=InlineKeyboardMarkup(kb))


async def send_next_step_prompt(
    chat_id: int, context: ContextTypes.DEFAULT_TYPE, task: Dict[str, Any]
) -> None:
    steps = task["steps"]
    idx = task["step_idx"]
    step = steps[idx]
    kb = ReplyKeyboardMarkup(
        [[KeyboardButton(step["name"])], [KeyboardButton("Cancel")]],
        resize_keyboard=True,
    )
    msg = await context.bot.send_message(
        chat_id=chat_id,
        text=f"🔹 ধাপ {idx + 1}/{len(steps)}: <b>{step['name']}</b> নিচের বাটনে ট্যাপ করুন অথবা Cancel করুন।",
        parse_mode=ParseMode.HTML,
        reply_markup=kb,
    )
    task["msg_ids"].append(msg.message_id)
    task["awaiting_button"] = True


async def handle_task_text(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str) -> bool:
    task = context.user_data.get("active_task")
    if not task:
        return False

    chat_id = update.effective_chat.id
    task["msg_ids"].append(update.message.message_id)

    if text == "Cancel":
        for mid in task["msg_ids"]:
            try:
                await context.bot.delete_message(chat_id=chat_id, message_id=mid)
            except Exception:
                pass
        context.user_data.pop("active_task", None)
        await update.message.reply_text("❌ কাজটি বাতিল করা হয়েছে।", reply_markup=main_keyboard(update.effective_user.id))
        return True

    steps = task["steps"]
    idx = task["step_idx"]
    current_step = steps[idx]

    if task.get("awaiting_button"):
        if text == current_step["name"]:
            task["awaiting_button"] = False
            pmsg = await update.message.reply_text(current_step["prompt"])
            task["msg_ids"].append(pmsg.message_id)
            return True
        else:
            wmsg = await update.message.reply_text(f"অনুগ্রহ করে নিচের <b>{current_step['name']}</b> বাটনে ক্লিক করুন অথবা Cancel চাপুন।", parse_mode=ParseMode.HTML)
            task["msg_ids"].append(wmsg.message_id)
            return True

    # Validate user input for current step
    ok, err_msg, fingerprint = validate_step_input(current_step.get("validator", "text"), text)
    if not ok:
        emsg = await update.message.reply_text(err_msg)
        task["msg_ids"].append(emsg.message_id)
        return True

    if fingerprint:
        dup = db.fetchone("SELECT id FROM submissions WHERE token_fingerprint = ?", (fingerprint,))
        if dup:
            dmsg = await update.message.reply_text("❌ এই Hotmail টি আগেই সাবমিট করা হয়েছে! নতুন একটি দিন।")
            task["msg_ids"].append(dmsg.message_id)
            return True
        task["token_fingerprint"] = fingerprint

    task["collected"][current_step["name"]] = text.strip()
    task["step_idx"] += 1

    if task["step_idx"] < len(steps):
        await send_next_step_prompt(chat_id, context, task)
        return True

    # All steps complete -> delete task messages, show verifying -> submitted
    for mid in task["msg_ids"]:
        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=mid)
        except Exception:
            pass

    ver_msg = await context.bot.send_message(
        chat_id=chat_id,
        text="⏳ Verifying submission...",
        reply_markup=main_keyboard(update.effective_user.id),
    )

    now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
    sub_id = db.execute_insert_returning_id(
        """
        INSERT INTO submissions (
            user_id, subcategory_id, subcategory_name, rate,
            gen_name, gen_username, gen_password, steps_data_json,
            token_fingerprint, status, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?)
        """,
        (
            update.effective_user.id,
            task["sub_id"],
            task["sub_name"],
            task["rate"],
            task["gen_name"],
            task["gen_username"],
            task["gen_password"],
            json.dumps(task["collected"], ensure_ascii=False),
            task.get("token_fingerprint", ""),
            now_str,
        ),
    )
    context.user_data.pop("active_task", None)

    await ver_msg.edit_text("✅ Submitted! Sent to admin review.")

    # Notify admins
    pending_cnt = db.fetchone(
        "SELECT COUNT(*) AS cnt FROM submissions WHERE user_id = ? AND status = 'pending'",
        (update.effective_user.id,),
    )["cnt"]
    u = update.effective_user
    for adm_id in get_all_admin_ids():
        try:
            await context.bot.send_message(
                chat_id=adm_id,
                text=(
                    f"📥 <b>নতুন সাবমিশন!</b> (#{sub_id})\n"
                    f"👤 ইউজার: {u.full_name} (<code>{u.id}</code>)\n"
                    f"📂 কাজ: {task['sub_name']} (৳{task['rate']:.0f})\n"
                    f"📦 এই ইউজারের মোট Pending: <b>{pending_cnt}টি</b>"
                ),
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass

    return True


# ==============================================================================
# BALANCE, REFERRAL, WITHDRAW, FILE SUBMIT, SUPPORT, GUIDE
# ==============================================================================
async def handle_balance(update: Update) -> None:
    u = ensure_user(update)
    in_review = db.fetchone(
        "SELECT COUNT(*) AS cnt FROM submissions WHERE user_id = ? AND status IN ('pending', 'received')",
        (u["user_id"],),
    )["cnt"]
    text = (
        "💼 <b>আপনার অ্যাকাউন্ট ব্যালেন্স</b>\n\n"
        f"💰 <b>মোট ব্যালেন্স:</b> ৳{u['balance']:.2f}\n"
        f"⏳ <b>পেন্ডিং উইথড্র:</b> ৳{u['pending_withdraw']:.2f}\n"
        f"📈 <b>টোটাল ইনকাম:</b> ৳{u['total_income']:.2f}\n"
        f"✅ <b>কমপ্লিট জব:</b> {u['completed_jobs']}টি\n"
        f"🔍 <b>ইন রিভিউ জব:</b> {in_review}টি"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.HTML)


async def handle_referral(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    u = ensure_user(update)
    bot_username = context.bot.username
    bonus = float(get_setting("referral_bonus", "3"))
    link = f"https://t.me/{bot_username}?start={u['user_id']}"
    text = (
        "🎁 <b>রেফার অ্যান্ড আর্ন</b>\n\n"
        f"আপনার রেফারেল লিংক শেয়ার করুন—আপনি যাকে রেফার করবেন তার <b>প্রতি কাজের জন্য আপনি ৳{bonus:.0f} করে বোনাস</b> পাবেন!\n\n"
        f"🔗 <b>আপনার রেফারেল লিংক:</b>\n<code>{link}</code>\n\n"
        f"👥 <b>মোট রেফার:</b> {u['referral_count']} জন\n"
        f"💵 <b>রেফারেল আয়:</b> ৳{u['referral_earnings']:.2f}"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.HTML)


async def handle_withdraw_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    u = ensure_user(update)
    min_w = float(get_setting("min_withdraw", "50"))
    fee = float(get_setting("withdraw_fee", "5"))
    if u["balance"] < min_w:
        await update.message.reply_text(
            f"❌ আপনার ব্যালেন্সে পর্যাপ্ত টাকা নেই।\n"
            f"সর্বনিম্ন উইথড্র: <b>৳{min_w:.0f}</b> (বর্তমান ব্যালেন্স: ৳{u['balance']:.2f})",
            parse_mode=ParseMode.HTML,
        )
        return
    kb = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("bKash", callback_data="wd:method:bKash"),
            InlineKeyboardButton("Nagad", callback_data="wd:method:Nagad"),
            InlineKeyboardButton("Binance", callback_data="wd:method:Binance"),
        ],
        [InlineKeyboardButton("❌ Cancel", callback_data="wd:cancel")],
    ])
    await update.message.reply_text(
        f"💸 <b>উইথড্র মেথড সিলেক্ট করুন</b>\n"
        f"• সর্বনিম্ন: ৳{min_w:.0f}\n"
        f"• প্রতি উইথড্র চার্জ: ৳{fee:.0f}",
        parse_mode=ParseMode.HTML,
        reply_markup=kb,
    )


async def handle_file_submit_menu(update: Update) -> None:
    cats = db.fetchall("SELECT * FROM file_submit_categories WHERE is_active = 1 ORDER BY id ASC")
    if not cats:
        await update.message.reply_text("👤 No file submit categories are active right now.")
        return
    kb = [[InlineKeyboardButton(f"📁 {c['name']}", callback_data=f"fsub:cat:{c['id']}")] for c in cats]
    await update.message.reply_text("📁 যে ক্যাটাগরিতে ফাইল জমা দেবেন তা সিলেক্ট করুন:", reply_markup=InlineKeyboardMarkup(kb))


async def handle_guide_menu(update: Update) -> None:
    guides = db.fetchall("SELECT * FROM guide_categories WHERE is_active = 1 ORDER BY id ASC")
    if not guides:
        await update.message.reply_text("ℹ️ এই মুহূর্তে কোনো গাইডলাইন যুক্ত নেই।")
        return
    kb = [[InlineKeyboardButton(f"🎬 {g['name']}", callback_data=f"guide:view:{g['id']}")] for g in guides]
    await update.message.reply_text("🆕 আপনি কোন কাজের নিয়ম ও ভিডিও দেখতে চান?", reply_markup=InlineKeyboardMarkup(kb))


# ==============================================================================
# CALLBACK QUERY ROUTER
# ==============================================================================
async def callback_router(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    data = query.data
    uid = query.from_user.id

    if get_setting("bot_active", "1") != "1" and not is_admin(uid):
        await query.message.reply_text("🔧 বটের আপডেটের কাজ চলতেছে।")
        return

    # --- WORK FLOW ---
    if data.startswith("work:cat:"):
        cat_id = int(data.split(":")[2])
        cat = db.fetchone("SELECT * FROM categories WHERE id = ?", (cat_id,))
        subs = db.fetchall(
            "SELECT * FROM subcategories WHERE category_id = ? AND is_active = 1 ORDER BY id ASC",
            (cat_id,),
        )
        if not subs:
            await query.edit_message_text("❌ এই ক্যাটাগরিতে কোনো কাজ চালু নেই।")
            return
        kb = [
            [InlineKeyboardButton(f"{s['name']} • ৳{s['rate']:.0f}", callback_data=f"work:sub:{s['id']}")]
            for s in subs
        ]
        await query.edit_message_text(
            f"📸 Select {cat['name']} work:",
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if data.startswith("work:sub:"):
        sub_id = int(data.split(":")[2])
        sub = db.fetchone("SELECT * FROM subcategories WHERE id = ? AND is_active = 1", (sub_id,))
        if not sub:
            await query.edit_message_text("❌ এই কাজটি বর্তমানে বন্ধ আছে।")
            return
        steps = json.loads(sub["steps_json"])
        if not steps:
            await query.edit_message_text("❌ এই কাজের কোনো ধাপ সেট করা হয়নি।")
            return

        gen_name, gen_username = generate_random_profile()
        gen_password = get_current_password()

        info_text = (
            f"👤 Name:- <code>{gen_name}</code>\n"
            f"🆔 Username:- <code>{gen_username}</code>\n"
            f"🔐 Password:- <code>{gen_password}</code>\n\n"
            f"📱 Tap the monospace text above to copy. Complete the account, then tap <b>{steps[0]['name']}</b> below 😊\n\n"
            f"💰 Rate: ৳{sub['rate']:.0f}"
        )
        info_msg = await query.message.reply_text(info_text, parse_mode=ParseMode.HTML)
        task = {
            "sub_id": sub["id"],
            "sub_name": sub["name"],
            "rate": float(sub["rate"]),
            "steps": steps,
            "step_idx": 0,
            "collected": {},
            "gen_name": gen_name,
            "gen_username": gen_username,
            "gen_password": gen_password,
            "msg_ids": [query.message.message_id, info_msg.message_id],
            "awaiting_button": True,
        }
        context.user_data["active_task"] = task
        await send_next_step_prompt(query.message.chat_id, context, task)
        return

    # --- GUIDE VIEW ---
    if data.startswith("guide:view:"):
        gid = int(data.split(":")[2])
        g = db.fetchone("SELECT * FROM guide_categories WHERE id = ?", (gid,))
        if not g:
            return
        if g.get("video_file_id"):
            await query.message.reply_video(
                video=g["video_file_id"],
                caption=g.get("guide_text") or g["name"],
            )
        else:
            await query.message.reply_text(g.get("guide_text") or "গাইডলাইন পাওয়া যায়নি।")
        return

    # --- FILE SUBMIT ---
    if data.startswith("fsub:cat:"):
        cid = int(data.split(":")[2])
        cat = db.fetchone("SELECT * FROM file_submit_categories WHERE id = ?", (cid,))
        if not cat:
            return
        context.user_data["state"] = f"user_fsub:{cat['name']}"
        await query.edit_message_text(f"📤 অনুগ্রহ করে <b>{cat['name']}</b>-এর জন্য আপনার ফাইলটি এখানে পাঠান:", parse_mode=ParseMode.HTML)
        return

    # --- WITHDRAW ---
    if data == "wd:cancel":
        context.user_data.pop("wd_data", None)
        context.user_data.pop("state", None)
        await query.edit_message_text("❌ উইথড্র বাতিল করা হয়েছে।")
        return

    if data.startswith("wd:method:"):
        method = data.split(":")[2]
        context.user_data["wd_data"] = {"method": method}
        context.user_data["state"] = "wd_account"
        prompt = "আপনার Binance Pay ID / UID দিন:" if method == "Binance" else f"আপনার {method} নম্বরটি দিন:"
        await query.edit_message_text(f"💳 মেথড: <b>{method}</b>\n{prompt}", parse_mode=ParseMode.HTML)
        return

    if data.startswith("wd_adm:"):
        if not is_admin(uid):
            return
        _, action, wd_id_str = data.split(":")
        wd_id = int(wd_id_str)
        wd = db.fetchone("SELECT * FROM withdrawals WHERE id = ?", (wd_id,))
        if not wd or wd["status"] != "pending":
            await query.edit_message_text("ℹ️ এই উইথড্র রিকোয়েস্ট আগেই প্রসেস করা হয়েছে।")
            return
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        if action == "paid":
            db.execute(
                "UPDATE withdrawals SET status = 'paid', processed_at = ? WHERE id = ?",
                (now_str, wd_id),
            )
            db.execute(
                "UPDATE users SET pending_withdraw = MAX(0, pending_withdraw - ?) WHERE user_id = ?",
                (wd["requested_amount"], wd["user_id"]),
            )
            await query.edit_message_text(
                f"✅ <b>Paid!</b> (#WD-{wd_id})\n"
                f"👤 UID: <code>{wd['user_id']}</code> | {wd['method']}: <code>{wd['account_number']}</code>\n"
                f"💵 পাঠানো হয়েছে: ৳{wd['net_amount']:.2f}",
                parse_mode=ParseMode.HTML,
            )
            try:
                await context.bot.send_message(
                    chat_id=wd["user_id"],
                    text=f"✅ আপনার ৳{wd['net_amount']:.2f} ({wd['method']} - {wd['account_number']}) উইথড্র পেমেন্ট সম্পন্ন হয়েছে!",
                )
            except Exception:
                pass
        elif action == "reject":
            db.execute(
                "UPDATE withdrawals SET status = 'rejected', processed_at = ? WHERE id = ?",
                (now_str, wd_id),
            )
            db.execute(
                "UPDATE users SET balance = balance + ?, pending_withdraw = MAX(0, pending_withdraw - ?) WHERE user_id = ?",
                (wd["requested_amount"], wd["requested_amount"], wd["user_id"]),
            )
            await query.edit_message_text(f"❌ <b>Rejected & Refunded</b> (#WD-{wd_id})", parse_mode=ParseMode.HTML)
            try:
                await context.bot.send_message(
                    chat_id=wd["user_id"],
                    text=f"❌ আপনার ৳{wd['requested_amount']:.2f} উইথড্র রিকোয়েস্ট রিজেক্ট করা হয়েছে এবং টাকা আপনার ব্যালেন্সে ফেরত দেওয়া হয়েছে।",
                )
            except Exception:
                pass
        return

    # --- ADMIN PANEL ---
    if data.startswith("adm:"):
        if not is_admin(uid):
            return
        await handle_admin_callback(query, context, data)


# ==============================================================================
# ADMIN PANEL CALLBACK HANDLER
# ==============================================================================
async def handle_admin_callback(query, context: ContextTypes.DEFAULT_TYPE, data: str) -> None:
    parts = data.split(":")
    action = parts[1]

    if action == "home":
        context.user_data.pop("state", None)
        await query.edit_message_text(
            "⚙️ <b>Admin Control Panel</b>\nনিচ থেকে যেকোনো অপশন সিলেক্ট করুন:",
            parse_mode=ParseMode.HTML,
            reply_markup=admin_panel_keyboard(),
        )
        return

    if action == "toggle_bot":
        cur = get_setting("bot_active", "1")
        set_setting("bot_active", "0" if cur == "1" else "1")
        await query.edit_message_reply_markup(reply_markup=admin_panel_keyboard())
        return

    if action == "pending_files":
        rows = db.fetchall(
            """
            SELECT s.user_id, s.subcategory_id, s.subcategory_name, COUNT(*) AS cnt, u.full_name
            FROM submissions s
            LEFT JOIN users u ON s.user_id = u.user_id
            WHERE s.status = 'pending'
            GROUP BY s.user_id, s.subcategory_id, s.subcategory_name, u.full_name
            ORDER BY cnt DESC
            """
        )
        if not rows:
            await query.edit_message_text("📂 কোনো পেন্ডিং ফাইল/সাবমিশন নেই।", reply_markup=back_to_admin_kb())
            return
        kb = []
        for r in rows:
            uname_lbl = r["full_name"] or str(r["user_id"])
            kb.append([
                InlineKeyboardButton(
                    f"📥 Receive: {uname_lbl} ({r['subcategory_name']} - {r['cnt']}টি)",
                    callback_data=f"adm:recv:{r['user_id']}:{r['subcategory_id']}",
                )
            ])
        kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="adm:home")])
        await query.edit_message_text(
            "📂 <b>Pending Submissions</b>\nরিসিভ করতে নিচের যেকোনো ইউজারে ক্লিক করুন:",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if action == "recv":
        target_uid = int(parts[2])
        sub_id = int(parts[3])
        subs = db.fetchall(
            "SELECT * FROM submissions WHERE user_id = ? AND subcategory_id = ? AND status = 'pending' ORDER BY id ASC",
            (target_uid, sub_id),
        )
        if not subs:
            await query.edit_message_text("❌ কোনো পেন্ডিং কাজ পাওয়া যায়নি।", reply_markup=back_to_admin_kb())
            return

        sub_meta = db.fetchone("SELECT * FROM subcategories WHERE id = ?", (sub_id,))
        step_names = [s["name"] for s in json.loads(sub_meta["steps_json"])] if sub_meta else []
        # Include any extra keys collected
        for s in subs:
            for k in json.loads(s["steps_data_json"]).keys():
                if k not in step_names:
                    step_names.append(k)

        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        batch_id = db.execute_insert_returning_id(
            """
            INSERT INTO batches (user_id, subcategory_id, subcategory_name, rate, total_count, status, received_at)
            VALUES (?, ?, ?, ?, ?, 'received', ?)
            """,
            (target_uid, sub_id, subs[0]["subcategory_name"], subs[0]["rate"], len(subs), now_str),
        )
        db.execute(
            "UPDATE submissions SET status = 'received', batch_id = ? WHERE user_id = ? AND subcategory_id = ? AND status = 'pending'",
            (batch_id, target_uid, sub_id),
        )

        # Build Excel (.xlsx)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Submissions"
        headers = ["SL", "Submission_ID", "Username", "Password"] + step_names + ["Status"]
        ws.append(headers)
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=1, column=col_idx)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center")

        for idx, s in enumerate(subs, start=1):
            sdata = json.loads(s["steps_data_json"])
            row_vals = [idx, s["id"], s["gen_username"], s["gen_password"]]
            for st_name in step_names:
                row_vals.append(sdata.get(st_name, ""))
            row_vals.append("PENDING")
            ws.append(row_vals)

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 12), 45)

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        filename = f"batch_{batch_id}_user_{target_uid}.xlsx"

        await query.message.reply_document(
            document=bio,
            filename=filename,
            caption=(
                f"📊 <b>Batch #{batch_id} Received!</b>\n"
                f"👤 UID: <code>{target_uid}</code> | কাজ: {subs[0]['subcategory_name']} ({len(subs)}টি)\n\n"
                f"💡 চেক করার পর সঠিক সারিতে <b>সবুজ (Green)</b> এবং ভুল সারিতে <b>লাল (Red)</b> কালার দিয়ে "
                f"<b>Upload Report (.xlsx)</b> অপশনে আপলোড করুন।"
            ),
            parse_mode=ParseMode.HTML,
        )

        try:
            await context.bot.send_message(
                chat_id=target_uid,
                text=f"📥 আপনার এই পর্যন্ত জমা করা <b>{len(subs)}টি</b> ({subs[0]['subcategory_name']}) কাজ অ্যাডমিন রিসিভ করেছেন।",
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass
        await query.edit_message_text(f"✅ Batch #{batch_id} রিসিভ সম্পন্ন হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if action == "upload_report":
        context.user_data["state"] = "adm_upload_report"
        await query.edit_message_text(
            "📤 সবুজ (Approve) ও লাল (Reject) কালার করা <b>.xlsx</b> রিপোর্ট ফাইলটি এখন চ্যাটে পাঠান:",
            parse_mode=ParseMode.HTML,
            reply_markup=back_to_admin_kb(),
        )
        return

    if action == "inbox_upload":
        context.user_data["state"] = "adm_inbox_uid"
        await query.edit_message_text(
            "📩 যে ইউজারের ইনবক্সে ফাইল পাঠাতে চান তার <b>User ID (UID)</b> লিখুন:",
            parse_mode=ParseMode.HTML,
            reply_markup=back_to_admin_kb(),
        )
        return

    if action == "work_mgr":
        cats = db.fetchall("SELECT * FROM categories ORDER BY id ASC")
        kb = []
        for c in cats:
            st = "🟢" if c["is_active"] else "🔴"
            kb.append([InlineKeyboardButton(f"{st} {c['name']} (Manage)", callback_data=f"adm:cat:{c['id']}")])
        kb.append([InlineKeyboardButton("➕ Add Main Category", callback_data="adm:add_cat")])
        kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="adm:home")])
        await query.edit_message_text(
            "🛠 <b>Work Categories Manager</b>",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if action == "add_cat":
        context.user_data["state"] = "adm_add_cat"
        await query.edit_message_text("নতুন মেইন ক্যাটাগরির নাম লিখুন (যেমন: Facebook):", reply_markup=back_to_admin_kb())
        return

    if action == "cat":
        cat_id = int(parts[2])
        cat = db.fetchone("SELECT * FROM categories WHERE id = ?", (cat_id,))
        subs = db.fetchall("SELECT * FROM subcategories WHERE category_id = ? ORDER BY id ASC", (cat_id,))
        kb = [
            [
                InlineKeyboardButton(
                    f"{'🟢 ON' if cat['is_active'] else '🔴 OFF'} Toggle Category",
                    callback_data=f"adm:cat_tog:{cat_id}",
                ),
                InlineKeyboardButton("🗑 Delete Category", callback_data=f"adm:cat_del:{cat_id}"),
            ]
        ]
        for s in subs:
            sst = "🟢" if s["is_active"] else "🔴"
            kb.append([
                InlineKeyboardButton(
                    f"{sst} {s['name']} (৳{s['rate']:.0f})",
                    callback_data=f"adm:sub:{s['id']}",
                )
            ])
        kb.append([InlineKeyboardButton("➕ Add Subcategory", callback_data=f"adm:add_sub:{cat_id}")])
        kb.append([InlineKeyboardButton("⬅️ Back", callback_data="adm:work_mgr")])
        await query.edit_message_text(
            f"📂 Category: <b>{cat['name']}</b>",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if action == "cat_tog":
        cat_id = int(parts[2])
        cat = db.fetchone("SELECT is_active FROM categories WHERE id = ?", (cat_id,))
        db.execute("UPDATE categories SET is_active = ? WHERE id = ?", (0 if cat["is_active"] else 1, cat_id))
        await handle_admin_callback(query, context, f"adm:cat:{cat_id}")
        return

    if action == "cat_del":
        cat_id = int(parts[2])
        db.execute("DELETE FROM subcategories WHERE category_id = ?", (cat_id,))
        db.execute("DELETE FROM categories WHERE id = ?", (cat_id,))
        await handle_admin_callback(query, context, "adm:work_mgr")
        return

    if action == "add_sub":
        cat_id = int(parts[2])
        context.user_data["state"] = f"adm_add_sub:{cat_id}"
        await query.edit_message_text(
            "নতুন সাব-ক্যাটাগরির নাম ও রেট এভাবে লিখুন:\n<code>Name | Rate</code>\n(উদাহরণ: <code>USA Facebook | 10</code>)",
            parse_mode=ParseMode.HTML,
            reply_markup=back_to_admin_kb(),
        )
        return

    if action == "sub":
        sub_id = int(parts[2])
        sub = db.fetchone("SELECT * FROM subcategories WHERE id = ?", (sub_id,))
        steps = json.loads(sub["steps_json"])
        step_txt = "\n".join([f"{i+1}. <b>{st['name']}</b> ({st.get('validator','text')})" for i, st in enumerate(steps)]) or "কোনো স্টেপ নেই"
        kb = [
            [
                InlineKeyboardButton(
                    f"{'🟢 ON' if sub['is_active'] else '🔴 OFF'} Toggle",
                    callback_data=f"adm:sub_tog:{sub_id}",
                ),
                InlineKeyboardButton("💲 Edit Rate", callback_data=f"adm:sub_rate:{sub_id}"),
            ],
            [
                InlineKeyboardButton("➕ Add Step", callback_data=f"adm:step_add:{sub_id}"),
                InlineKeyboardButton("🗑 Clear Last Step", callback_data=f"adm:step_pop:{sub_id}"),
            ],
            [InlineKeyboardButton("🗑 Delete Subcategory", callback_data=f"adm:sub_del:{sub_id}")],
            [InlineKeyboardButton("⬅️ Back", callback_data=f"adm:cat:{sub['category_id']}")],
        ]
        await query.edit_message_text(
            f"📌 <b>{sub['name']}</b> | Rate: <b>৳{sub['rate']:.0f}</b>\n\n<b>Steps:</b>\n{step_txt}",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if action == "sub_tog":
        sub_id = int(parts[2])
        sub = db.fetchone("SELECT is_active FROM subcategories WHERE id = ?", (sub_id,))
        db.execute("UPDATE subcategories SET is_active = ? WHERE id = ?", (0 if sub["is_active"] else 1, sub_id))
        await handle_admin_callback(query, context, f"adm:sub:{sub_id}")
        return

    if action == "sub_rate":
        sub_id = int(parts[2])
        context.user_data["state"] = f"adm_sub_rate:{sub_id}"
        await query.edit_message_text("নতুন রেট (সংখ্যা) লিখুন (যেমন: 6):", reply_markup=back_to_admin_kb())
        return

    if action == "step_add":
        sub_id = int(parts[2])
        context.user_data["state"] = f"adm_step_add:{sub_id}"
        await query.edit_message_text(
            "নতুন স্টেপ এই ফরম্যাটে লিখুন:\n"
            "<code>Button Name | Prompt Message | validator</code>\n\n"
            "validator হতে পারে: <code>hotmail_token</code>, <code>2fa</code>, বা <code>text</code>\n"
            "উদাহরণ: <code>2FA Set | আপনার 2FA Key দিন: | 2fa</code>",
            parse_mode=ParseMode.HTML,
            reply_markup=back_to_admin_kb(),
        )
        return

    if action == "step_pop":
        sub_id = int(parts[2])
        sub = db.fetchone("SELECT steps_json FROM subcategories WHERE id = ?", (sub_id,))
        steps = json.loads(sub["steps_json"])
        if steps:
            steps.pop()
            db.execute("UPDATE subcategories SET steps_json = ? WHERE id = ?", (json.dumps(steps, ensure_ascii=False), sub_id))
        await handle_admin_callback(query, context, f"adm:sub:{sub_id}")
        return

    if action == "sub_del":
        sub_id = int(parts[2])
        sub = db.fetchone("SELECT category_id FROM subcategories WHERE id = ?", (sub_id,))
        db.execute("DELETE FROM subcategories WHERE id = ?", (sub_id,))
        await handle_admin_callback(query, context, f"adm:cat:{sub['category_id']}")
        return

    if action == "bal_mgr":
        min_w = get_setting("min_withdraw", "50")
        fee = get_setting("withdraw_fee", "5")
        ref_b = get_setting("referral_bonus", "3")
        kb = [
            [InlineKeyboardButton(f"📉 Min Withdraw: ৳{min_w}", callback_data="adm:set_min_wd")],
            [InlineKeyboardButton(f"💸 Withdraw Fee: ৳{fee}", callback_data="adm:set_wd_fee")],
            [InlineKeyboardButton(f"🎁 Referral Bonus: ৳{ref_b}", callback_data="adm:set_ref_bonus")],
            [InlineKeyboardButton("💳 Edit User Balance (UID)", callback_data="adm:edit_user_bal")],
            [InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="adm:home")],
        ]
        await query.edit_message_text("💰 <b>Balance & Withdraw Settings</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action in ("set_min_wd", "set_wd_fee", "set_ref_bonus", "edit_user_bal"):
        context.user_data["state"] = f"adm_{action}"
        prompts = {
            "set_min_wd": "নতুন Minimum Withdraw (সংখ্যা) লিখুন (যেমন: 50 বা 100):",
            "set_wd_fee": "নতুন Withdraw Fee লিখুন (যেমন: 5):",
            "set_ref_bonus": "নতুন Referral Bonus লিখুন (যেমন: 3):",
            "edit_user_bal": "<code>UID | Amount</code> লিখুন (যোগ করতে <code>12345 | 50</code>, কাটতে <code>12345 | -20</code>):",
        }
        await query.edit_message_text(prompts[action], parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "pwd_mgr":
        cur_pwd = get_current_password()
        prefix = get_setting("password_prefix", "succes#")
        rtime = get_setting("password_reset_time", "00:00")
        kb = [
            [InlineKeyboardButton(f"✏️ Prefix: {prefix}", callback_data="adm:set_pwd_prefix")],
            [InlineKeyboardButton(f"⏰ Reset Time (BD): {rtime}", callback_data="adm:set_pwd_time")],
            [InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="adm:home")],
        ]
        await query.edit_message_text(
            f"🔑 <b>Auto Password Settings</b>\nবর্তমান পাসওয়ার্ড: <code>{cur_pwd}</code>",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if action == "set_pwd_prefix":
        context.user_data["state"] = "adm_set_pwd_prefix"
        await query.edit_message_text("নতুন Prefix লিখুন (যেমন: <code>succes#</code>):", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "set_pwd_time":
        context.user_data["state"] = "adm_set_pwd_time"
        await query.edit_message_text("বাংলাদেশ সময় (২৪ ঘণ্টা ফরম্যাটে) লিখুন (যেমন: <code>00:00</code> বা <code>06:00</code>):", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "btn_mgr":
        extra_btns = json.loads(get_setting("extra_buttons_json", "[]"))
        kb = []
        for idx, b in enumerate(extra_btns):
            kb.append([InlineKeyboardButton(f"🗑 Remove: {b['label']}", callback_data=f"adm:btn_del:{idx}")])
        kb.append([InlineKeyboardButton("➕ Add Main Button", callback_data="adm:btn_add")])
        kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="adm:home")])
        await query.edit_message_text("⌨️ <b>Main Keyboard Extra Buttons</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "btn_add":
        context.user_data["state"] = "adm_btn_add"
        await query.edit_message_text(
            "<code>Button Name | Reply Text</code> ফরম্যাটে লিখুন:\n(উদাহরণ: <code>Notice | 📢 আজকের কাজ চালু আছে!</code>)",
            parse_mode=ParseMode.HTML,
            reply_markup=back_to_admin_kb(),
        )
        return

    if action == "btn_del":
        idx = int(parts[2])
        extra_btns = json.loads(get_setting("extra_buttons_json", "[]"))
        if 0 <= idx < len(extra_btns):
            extra_btns.pop(idx)
            set_setting("extra_buttons_json", json.dumps(extra_btns, ensure_ascii=False))
        await handle_admin_callback(query, context, "adm:btn_mgr")
        return

    if action == "guide_mgr":
        guides = db.fetchall("SELECT * FROM guide_categories ORDER BY id ASC")
        kb = [[InlineKeyboardButton(f"🎬 {g['name']}", callback_data=f"adm:guide_edit:{g['id']}")] for g in guides]
        kb.append([InlineKeyboardButton("➕ Add Guide Category", callback_data="adm:guide_add")])
        kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="adm:home")])
        await query.edit_message_text("🎬 <b>আমি নতুন (Guide Manager)</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "guide_add":
        context.user_data["state"] = "adm_guide_add"
        await query.edit_message_text("নতুন গাইড ক্যাটাগরির নাম লিখুন:", reply_markup=back_to_admin_kb())
        return

    if action == "guide_edit":
        gid = int(parts[2])
        context.user_data["state"] = f"adm_guide_upload:{gid}"
        await query.edit_message_text("এই ক্যাটাগরির জন্য একটি টিউটোরিয়াল <b>Video</b> অথবা <b>Text</b> পাঠান:", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "fsub_mgr":
        cats = db.fetchall("SELECT * FROM file_submit_categories ORDER BY id ASC")
        kb = []
        for c in cats:
            st = "🟢" if c["is_active"] else "🔴"
            kb.append([InlineKeyboardButton(f"{st} {c['name']} (Toggle)", callback_data=f"adm:fsub_tog:{c['id']}")])
        kb.append([InlineKeyboardButton("➕ Add File Submit Category", callback_data="adm:fsub_add")])
        kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="adm:home")])
        await query.edit_message_text("📁 <b>File Submit Categories</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "fsub_tog":
        cid = int(parts[2])
        c = db.fetchone("SELECT is_active FROM file_submit_categories WHERE id = ?", (cid,))
        db.execute("UPDATE file_submit_categories SET is_active = ? WHERE id = ?", (0 if c["is_active"] else 1, cid))
        await handle_admin_callback(query, context, "adm:fsub_mgr")
        return

    if action == "fsub_add":
        context.user_data["state"] = "adm_fsub_add"
        await query.edit_message_text("নতুন File Submit ক্যাটাগরির নাম লিখুন:", reply_markup=back_to_admin_kb())
        return

    if action == "user_dir":
        page = int(parts[2]) if len(parts) > 2 else 0
        limit = 10
        offset = page * limit
        users = db.fetchall("SELECT * FROM users ORDER BY joined_at DESC LIMIT ? OFFSET ?", (limit, offset))
        total_u = db.fetchone("SELECT COUNT(*) AS cnt FROM users")["cnt"]
        lines = [f"👥 <b>User Directory</b> (মোট: {total_u} জন)\n"]
        for usr in users:
            lines.append(
                f"• <code>{usr['user_id']}</code> | {usr['full_name']} | ৳{usr['balance']:.1f} | ✅{usr['completed_jobs']}"
            )
        kb = []
        nav = []
        if page > 0:
            nav.append(InlineKeyboardButton("⬅️ Prev", callback_data=f"adm:user_dir:{page-1}"))
        if offset + limit < total_u:
            nav.append(InlineKeyboardButton("Next ➡️", callback_data=f"adm:user_dir:{page+1}"))
        if nav:
            kb.append(nav)
        kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="adm:home")])
        await query.edit_message_text("\n".join(lines), parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "file_history":
        batches = db.fetchall("SELECT * FROM batches ORDER BY id DESC LIMIT 15")
        if not batches:
            await query.edit_message_text("📜 কোনো ফাইল হিস্ট্রি নেই।", reply_markup=back_to_admin_kb())
            return
        lines = ["📜 <b>Recent File History</b>\n"]
        for b in batches:
            lines.append(
                f"• <b>Batch #{b['id']}</b> | UID: <code>{b['user_id']}</code> | {b['subcategory_name']} | "
                f"Total: {b['total_count']} | ✅{b['approved_count']} ❌{b['rejected_count']} ({b['status']})"
            )
        await query.edit_message_text("\n".join(lines), parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "del_user":
        context.user_data["state"] = "adm_del_user"
        await query.edit_message_text("যার ডেটা ডিলিট করতে চান তার <b>UID</b> লিখুন:", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "admins":
        adm_ids = get_all_admin_ids()
        kb = []
        for aid in adm_ids:
            if aid == OWNER_ID:
                kb.append([InlineKeyboardButton(f"👑 Owner: {aid}", callback_data="adm:noop")])
            else:
                kb.append([InlineKeyboardButton(f"🗑 Remove Admin: {aid}", callback_data=f"adm:adm_del:{aid}")])
        kb.append([InlineKeyboardButton("➕ Add Admin (UID)", callback_data="adm:adm_add")])
        kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="adm:home")])
        await query.edit_message_text("👮 <b>Manage Admins</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "adm_add":
        context.user_data["state"] = "adm_add_admin"
        await query.edit_message_text("নতুন অ্যাডমিনের Telegram User ID লিখুন:", reply_markup=back_to_admin_kb())
        return

    if action == "adm_del":
        target_aid = int(parts[2])
        if target_aid != OWNER_ID:
            db.execute("DELETE FROM admins WHERE user_id = ?", (target_aid,))
        await handle_admin_callback(query, context, "adm:admins")
        return

    if action == "broadcast":
        context.user_data["state"] = "adm_broadcast"
        await query.edit_message_text("📢 সব ইউজারকে পাঠানোর জন্য মেসেজ লিখুন:", reply_markup=back_to_admin_kb())
        return

    if action == "edit_support":
        context.user_data["state"] = "adm_edit_support"
        cur_sup = get_setting("support_link", "")
        await query.edit_message_text(f"বর্তমান সাপোর্ট লিংক: {cur_sup}\nনতুন লিংক দিন:", reply_markup=back_to_admin_kb())
        return


# ==============================================================================
# TEXT, DOCUMENT & VIDEO MESSAGE HANDLERS
# ==============================================================================
async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    uid = update.effective_user.id
    ensure_user(update)

    if get_setting("bot_active", "1") != "1" and not is_admin(uid):
        await update.message.reply_text("🔧 বটের আপডেটের কাজ চলতেছে। অনুগ্রহ করে কিছুক্ষণ পর আবার চেষ্টা করুন।")
        return

    text = (update.message.text or "").strip()

    # Check active work task first
    if context.user_data.get("active_task") and text:
        handled = await handle_task_text(update, context, text)
        if handled:
            return

    # Check stateful inputs
    state = context.user_data.get("state", "")
    if state:
        await handle_stateful_message(update, context, state)
        return

    # Main keyboard menu
    if text == "Work":
        await handle_work_menu(update, context)
    elif text == "File Submit":
        await handle_file_submit_menu(update)
    elif text == "Balance":
        await handle_balance(update)
    elif text == "Withdraw":
        await handle_withdraw_start(update, context)
    elif text == "My Referral":
        await handle_referral(update, context)
    elif text == "Support":
        sup = get_setting("support_link", "https://t.me/Talha_juba098")
        await update.message.reply_text(
            "🆘 যেকোনো সাহায্য বা সমস্যার জন্য সাপোর্টে যোগাযোগ করুন:",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("💬 Contact Support", url=sup)]]),
        )
    elif text == "আমি নতুন":
        await handle_guide_menu(update)
    elif text == "⚙️ Admin Panel" and is_admin(uid):
        await admin_cmd(update, context)
    else:
        # Check extra custom buttons
        extra_btns = json.loads(get_setting("extra_buttons_json", "[]"))
        for b in extra_btns:
            if text == b["label"]:
                await update.message.reply_text(b["reply"])
                return


async def handle_stateful_message(update: Update, context: ContextTypes.DEFAULT_TYPE, state: str) -> None:
    uid = update.effective_user.id
    msg = update.message
    text = (msg.text or "").strip()

    # --- USER WITHDRAW ---
    if state == "wd_account":
        context.user_data["wd_data"]["account"] = text
        context.user_data["state"] = "wd_amount"
        min_w = float(get_setting("min_withdraw", "50"))
        await msg.reply_text(f"💵 কত টাকা উইথড্র করতে চান লিখুন (সর্বনিম্ন ৳{min_w:.0f}):")
        return

    if state == "wd_amount":
        try:
            amount = float(text)
        except ValueError:
            await msg.reply_text("❌ সঠিক সংখ্যা লিখুন (যেমন: 50):")
            return
        min_w = float(get_setting("min_withdraw", "50"))
        fee = float(get_setting("withdraw_fee", "5"))
        u = ensure_user(update)
        if amount < min_w:
            await msg.reply_text(f"❌ সর্বনিম্ন উইথড্র ৳{min_w:.0f}।")
            return
        if amount > u["balance"]:
            await msg.reply_text(f"❌ আপনার ব্যালেন্সে পর্যাপ্ত টাকা নেই (ব্যালেন্স: ৳{u['balance']:.2f})।")
            return
        net = max(0.0, amount - fee)
        wd_data = context.user_data.pop("wd_data")
        context.user_data.pop("state", None)

        db.execute(
            "UPDATE users SET balance = balance - ?, pending_withdraw = pending_withdraw + ? WHERE user_id = ?",
            (amount, amount, uid),
        )
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        wd_id = db.execute_insert_returning_id(
            """
            INSERT INTO withdrawals (user_id, method, account_number, requested_amount, fee, net_amount, status, created_at)
            VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)
            """,
            (uid, wd_data["method"], wd_data["account"], amount, fee, net, now_str),
        )
        await msg.reply_text(
            f"✅ <b>উইথড্র রিকোয়েস্ট সম্পন্ন!</b> (#WD-{wd_id})\n"
            f"• মেথড: {wd_data['method']}\n"
            f"• অ্যাকাউন্ট: <code>{wd_data['account']}</code>\n"
            f"• রিকোয়েস্ট: ৳{amount:.2f} | ফি: ৳{fee:.2f}\n"
            f"• <b>আপনি পাবেন: ৳{net:.2f}</b>",
            parse_mode=ParseMode.HTML,
        )
        adm_kb = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("✅ Paid", callback_data=f"wd_adm:paid:{wd_id}"),
                InlineKeyboardButton("❌ Reject & Refund", callback_data=f"wd_adm:reject:{wd_id}"),
            ]
        ])
        for adm_id in get_all_admin_ids():
            try:
                await context.bot.send_message(
                    chat_id=adm_id,
                    text=(
                        f"💸 <b>নতুন উইথড্র রিকোয়েস্ট!</b> (#WD-{wd_id})\n"
                        f"👤 ইউজার: {update.effective_user.full_name} (<code>{uid}</code>)\n"
                        f"💳 মেথড: <b>{wd_data['method']}</b>\n"
                        f"📱 নম্বর/ID: <code>{wd_data['account']}</code>\n"
                        f"📊 হিসাব: ৳{amount:.2f} - ৳{fee:.2f} ফি = <b>৳{net:.2f}</b>"
                    ),
                    parse_mode=ParseMode.HTML,
                    reply_markup=adm_kb,
                )
            except Exception:
                pass
        return

    # --- USER FILE SUBMIT ---
    if state.startswith("user_fsub:"):
        cat_name = state.split(":", 1)[1]
        if not msg.document:
            await msg.reply_text("❌ অনুগ্রহ করে একটি ডকুমেন্ট/ফাইল পাঠান।")
            return
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        db.execute(
            "INSERT INTO file_submissions (user_id, category_name, file_id, file_name, created_at) VALUES (?, ?, ?, ?, ?)",
            (uid, cat_name, msg.document.file_id, msg.document.file_name or "file", now_str),
        )
        context.user_data.pop("state", None)
        await msg.reply_text("✅ আপনার ফাইলটি অ্যাডমিনের কাছে জমা দেওয়া হয়েছে!")
        for adm_id in get_all_admin_ids():
            try:
                await context.bot.send_document(
                    chat_id=adm_id,
                    document=msg.document.file_id,
                    caption=f"📁 <b>File Submit</b> ({cat_name})\n👤 UID: <code>{uid}</code>",
                    parse_mode=ParseMode.HTML,
                )
            except Exception:
                pass
        return

    # --- ADMIN STATES ---
    if not is_admin(uid):
        context.user_data.pop("state", None)
        return

    if state == "adm_upload_report":
        if not msg.document or not (msg.document.file_name or "").lower().endswith(".xlsx"):
            await msg.reply_text("❌ অনুগ্রহ করে একটি ভ্যালিড `.xlsx` এক্সেল ফাইল পাঠান।")
            return
        tg_file = await msg.document.get_file()
        bio = io.BytesIO()
        await tg_file.download_to_memory(out=bio)
        bio.seek(0)
        await process_admin_report_excel(msg, context, bio)
        context.user_data.pop("state", None)
        return

    if state == "adm_inbox_uid":
        if not text.isdigit():
            await msg.reply_text("❌ সঠিক সংখ্যার UID দিন।")
            return
        context.user_data["inbox_target_uid"] = int(text)
        context.user_data["state"] = "adm_inbox_file"
        await msg.reply_text("এবার ইউজারের ইনবক্সে পাঠানোর জন্য ফাইল, ছবি বা মেসেজ দিন:")
        return

    if state == "adm_inbox_file":
        target_uid = context.user_data.pop("inbox_target_uid")
        context.user_data.pop("state", None)
        try:
            await msg.copy(chat_id=target_uid)
            await msg.reply_text(f"✅ ইউজার <code>{target_uid}</code>-এর ইনবক্সে সফলভাবে পাঠানো হয়েছে!", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        except Exception as e:
            await msg.reply_text(f"❌ পাঠানো যায়নি: {e}", reply_markup=back_to_admin_kb())
        return

    if state == "adm_add_cat":
        db.execute("INSERT INTO categories (name, is_active) VALUES (?, 1)", (text,))
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ ক্যাটাগরি <b>{text}</b> যুক্ত হয়েছে।", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_add_sub:"):
        cat_id = int(state.split(":")[1])
        if "|" not in text:
            await msg.reply_text("❌ ফরম্যাট: `Name | Rate`")
            return
        sname, srate_str = [x.strip() for x in text.split("|", 1)]
        srate = float(srate_str)
        default_steps = [{"name": "Submit Details", "prompt": "কাজের তথ্য দিন:", "validator": "text"}]
        db.execute(
            "INSERT INTO subcategories (category_id, name, rate, steps_json, is_active) VALUES (?, ?, ?, ?, 1)",
            (cat_id, sname, srate, json.dumps(default_steps, ensure_ascii=False)),
        )
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ সাব-ক্যাটাগরি <b>{sname}</b> (৳{srate:.0f}) যুক্ত হয়েছে।", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_sub_rate:"):
        sub_id = int(state.split(":")[1])
        db.execute("UPDATE subcategories SET rate = ? WHERE id = ?", (float(text), sub_id))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ রেট আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_step_add:"):
        sub_id = int(state.split(":")[1])
        parts = [x.strip() for x in text.split("|")]
        if len(parts) < 2:
            await msg.reply_text("❌ ফরম্যাট: `Button Name | Prompt Message | validator`")
            return
        bname = parts[0]
        bprompt = parts[1]
        bval = parts[2] if len(parts) >= 3 else "text"
        sub = db.fetchone("SELECT steps_json FROM subcategories WHERE id = ?", (sub_id,))
        steps = json.loads(sub["steps_json"])
        steps.append({"name": bname, "prompt": bprompt, "validator": bval})
        db.execute("UPDATE subcategories SET steps_json = ? WHERE id = ?", (json.dumps(steps, ensure_ascii=False), sub_id))
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ নতুন স্টেপ <b>{bname}</b> যুক্ত হয়েছে।", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if state == "adm_set_min_wd":
        set_setting("min_withdraw", str(float(text)))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ Minimum Withdraw আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_set_wd_fee":
        set_setting("withdraw_fee", str(float(text)))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ Withdraw Fee আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_set_ref_bonus":
        set_setting("referral_bonus", str(float(text)))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ Referral Bonus আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_edit_user_bal":
        p = [x.strip() for x in text.split("|")]
        tuid, amt = int(p[0]), float(p[1])
        db.execute("UPDATE users SET balance = balance + ? WHERE user_id = ?", (amt, tuid))
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ UID {tuid}-এর ব্যালেন্সে {amt:+.2f} আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_set_pwd_prefix":
        set_setting("password_prefix", text)
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ নতুন পাসওয়ার্ড: <code>{get_current_password()}</code>", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if state == "adm_set_pwd_time":
        set_setting("password_reset_time", text)
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ রিসেট টাইম আপডেট হয়েছে ({text})।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_btn_add":
        p = [x.strip() for x in text.split("|", 1)]
        extra_btns = json.loads(get_setting("extra_buttons_json", "[]"))
        extra_btns.append({"label": p[0], "reply": p[1]})
        set_setting("extra_buttons_json", json.dumps(extra_btns, ensure_ascii=False))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ নতুন বাটন যুক্ত হয়েছে!", reply_markup=back_to_admin_kb())
        return

    if state == "adm_guide_add":
        db.execute("INSERT INTO guide_categories (name, guide_text, video_file_id, is_active) VALUES (?, ?, '', 1)", (text, "গাইডলাইন"))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ গাইড ক্যাটাগরি যুক্ত হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_guide_upload:"):
        gid = int(state.split(":")[1])
        if msg.video:
            db.execute(
                "UPDATE guide_categories SET video_file_id = ?, guide_text = ? WHERE id = ?",
                (msg.video.file_id, msg.caption or "", gid),
            )
        else:
            db.execute("UPDATE guide_categories SET guide_text = ? WHERE id = ?", (text, gid))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ গাইড আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_fsub_add":
        db.execute("INSERT INTO file_submit_categories (name, is_active) VALUES (?, 1)", (text,))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ File Submit ক্যাটাগরি যুক্ত হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_del_user":
        tuid = int(text)
        db.execute("DELETE FROM submissions WHERE user_id = ?", (tuid,))
        db.execute("DELETE FROM withdrawals WHERE user_id = ?", (tuid,))
        db.execute("DELETE FROM users WHERE user_id = ?", (tuid,))
        context.user_data.pop("state", None)
        await msg.reply_text(f"🗑 UID {tuid}-এর সব ডেটা মুছে ফেলা হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_add_admin":
        aid = int(text)
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        db.execute(
            "INSERT INTO admins (user_id, added_by, created_at) VALUES (?, ?, ?) ON CONFLICT DO NOTHING"
            if db.use_pg
            else "INSERT OR IGNORE INTO admins (user_id, added_by, created_at) VALUES (?, ?, ?)",
            (aid, uid, now_str),
        )
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ {aid} কে অ্যাডমিন করা হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_broadcast":
        users = db.fetchall("SELECT user_id FROM users")
        sent = 0
        for usr in users:
            try:
                await context.bot.send_message(chat_id=usr["user_id"], text=text)
                sent += 1
            except Exception:
                pass
        context.user_data.pop("state", None)
        await msg.reply_text(f"📢 ব্রডকাস্ট সম্পন্ন! পাঠানো হয়েছে: {sent} জনকে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_edit_support":
        set_setting("support_link", text)
        context.user_data.pop("state", None)
        await msg.reply_text("✅ সাপোর্ট লিংক আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return


# ==============================================================================
# EXCEL GREEN/RED REPORT PROCESSOR
# ==============================================================================
async def process_admin_report_excel(msg, context: ContextTypes.DEFAULT_TYPE, bio: io.BytesIO) -> None:
    wb = openpyxl.load_workbook(bio)
    ws = wb.active

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    if "Submission_ID" not in headers:
        await msg.reply_text("❌ এই এক্সেল ফাইলে `Submission_ID` কলাম পাওয়া যায়নি!", reply_markup=back_to_admin_kb())
        return

    sub_id_col = headers.index("Submission_ID") + 1
    status_col = headers.index("Status") + 1 if "Status" in headers else len(headers)

    user_summary: Dict[int, Dict[str, Any]] = {}
    batch_stats: Dict[int, Dict[str, int]] = {}
    ref_bonus = float(get_setting("referral_bonus", "3"))

    for row_idx in range(2, ws.max_row + 1):
        sub_val = ws.cell(row=row_idx, column=sub_id_col).value
        if not sub_val or not str(sub_val).isdigit():
            continue
        sub_id = int(sub_val)
        sub = db.fetchone("SELECT * FROM submissions WHERE id = ?", (sub_id,))
        if not sub or sub["status"] in ("approved", "rejected"):
            continue

        # Detect green or red across Status cell or entire row
        color = classify_excel_color(ws.cell(row=row_idx, column=status_col))
        if not color:
            for c_idx in range(1, len(headers) + 1):
                color = classify_excel_color(ws.cell(row=row_idx, column=c_idx))
                if color:
                    break
        if not color:
            continue

        tuid = sub["user_id"]
        if tuid not in user_summary:
            user_summary[tuid] = {"approved": 0, "rejected": 0, "earned": 0.0}
        bid = sub.get("batch_id")
        if bid and bid not in batch_stats:
            batch_stats[bid] = {"approved": 0, "rejected": 0}

        if color == "green":
            rate = float(sub["rate"])
            db.execute("UPDATE submissions SET status = 'approved' WHERE id = ?", (sub_id,))
            db.execute(
                "UPDATE users SET balance = balance + ?, total_income = total_income + ?, completed_jobs = completed_jobs + 1 WHERE user_id = ?",
                (rate, rate, tuid),
            )
            user_summary[tuid]["approved"] += 1
            user_summary[tuid]["earned"] += rate
            if bid:
                batch_stats[bid]["approved"] += 1

            # Check first-time referral bonus
            usr = db.fetchone("SELECT referred_by, referral_paid FROM users WHERE user_id = ?", (tuid,))
            if usr and usr.get("referred_by") and not usr.get("referral_paid"):
                ref_id = usr["referred_by"]
                db.execute("UPDATE users SET referral_paid = 1 WHERE user_id = ?", (tuid,))
                db.execute(
                    "UPDATE users SET balance = balance + ?, total_income = total_income + ?, referral_earnings = referral_earnings + ? WHERE user_id = ?",
                    (ref_bonus, ref_bonus, ref_bonus, ref_id),
                )
                try:
                    await context.bot.send_message(
                        chat_id=ref_id,
                        text=f"🎉 অভিনন্দন! আপনার রেফার করা ইউজারের প্রথম কাজ অ্যাপ্রুভ হওয়ায় আপনার ব্যালেন্সে <b>৳{ref_bonus:.0f}</b> বোনাস যোগ হয়েছে!",
                        parse_mode=ParseMode.HTML,
                    )
                except Exception:
                    pass

        elif color == "red":
            db.execute("UPDATE submissions SET status = 'rejected' WHERE id = ?", (sub_id,))
            db.execute("UPDATE users SET rejected_jobs = rejected_jobs + 1 WHERE user_id = ?", (tuid,))
            user_summary[tuid]["rejected"] += 1
            if bid:
                batch_stats[bid]["rejected"] += 1

    now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
    for bid, st in batch_stats.items():
        db.execute(
            "UPDATE batches SET approved_count = approved_count + ?, rejected_count = rejected_count + ?, status = 'reported', reported_at = ? WHERE id = ?",
            (st["approved"], st["rejected"], now_str, bid),
        )

    total_app = sum(v["approved"] for v in user_summary.values())
    total_rej = sum(v["rejected"] for v in user_summary.values())
    for tuid, v in user_summary.items():
        try:
            await context.bot.send_message(
                chat_id=tuid,
                text=(
                    "📊 <b>কাজের রিপোর্ট প্রকাশিত হয়েছে!</b>\n\n"
                    f"✅ অ্যাপ্রুভ: <b>{v['approved']}টি</b>\n"
                    f"❌ রিজেক্ট: <b>{v['rejected']}টি</b>\n"
                    f"💰 ব্যালেন্সে যোগ হয়েছে: <b>৳{v['earned']:.2f}</b>"
                ),
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass

    await msg.reply_text(
        f"✅ <b>রিপোর্ট প্রসেস সম্পন্ন!</b>\n"
        f"• মোট অ্যাপ্রুভ (সবুজ): {total_app}টি\n"
        f"• মোট রিজেক্ট (লাল): {total_rej}টি",
        parse_mode=ParseMode.HTML,
        reply_markup=back_to_admin_kb(),
    )


# ==============================================================================
# RENDER KEEP-ALIVE WEB SERVER & BOT RUNNER
# ==============================================================================
async def health_handler(request):
    return web.Response(text="Bot is running!", status=200)


async def run_bot() -> None:
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN environment variable is missing!")

    application = Application.builder().token(BOT_TOKEN).build()
    application.add_handler(CommandHandler("start", start_cmd))
    application.add_handler(CommandHandler("admin", admin_cmd))
    application.add_handler(CallbackQueryHandler(callback_router))
    application.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, message_handler))

    # Start aiohttp web server for Render & UptimeRobot
    web_app = web.Application()
    web_app.router.add_get("/", health_handler)
    web_app.router.add_get("/health", health_handler)

    runner = web.AppRunner(web_app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    logger.info(f"Keep-alive web server started on port {PORT}")

    async with application:
        await application.initialize()
        await application.bot.delete_webhook(drop_pending_updates=True)
        await application.start()
        await application.updater.start_polling(drop_pending_updates=True)
        logger.info("Telegram bot polling started...")
        try:
            while True:
                await asyncio.sleep(3600)
        finally:
            await application.updater.stop()
            await application.stop()
            await runner.cleanup()


if __name__ == "__main__":
    asyncio.run(run_bot())
