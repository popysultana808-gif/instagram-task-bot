import asyncio
import datetime
import io
import json
import logging
import os
import random
import re
import sqlite3
import time
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
    import asyncpg
except ImportError:
    asyncpg = None

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
WEBHOOK_BASE = (os.getenv("WEBHOOK_URL") or os.getenv("RENDER_EXTERNAL_URL") or "").strip().rstrip("/")
OWNER_ID = 8919985167
BD_TZ = datetime.timezone(datetime.timedelta(hours=6))


def kb_btn(text: str, style: str = "primary") -> KeyboardButton:
    try:
        return KeyboardButton(text, style=style)
    except TypeError:
        return KeyboardButton(text)


def ikb_btn(
    text: str,
    callback_data: Optional[str] = None,
    url: Optional[str] = None,
    switch_inline_query: Optional[str] = None,
    style: str = "primary",
) -> InlineKeyboardButton:
    kwargs: Dict[str, Any] = {}
    if callback_data is not None:
        kwargs["callback_data"] = callback_data
    if url is not None:
        kwargs["url"] = url
    if switch_inline_query is not None:
        kwargs["switch_inline_query"] = switch_inline_query
    try:
        return InlineKeyboardButton(text, style=style, **kwargs)
    except TypeError:
        return InlineKeyboardButton(text, **kwargs)


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


# ==============================================================================
# ASYNC DATABASE LAYER (asyncpg for Supabase/Postgres + SQLite fallback)
# ==============================================================================
class AsyncDatabase:
    def __init__(self, db_url: str = ""):
        self.db_url = db_url
        self.use_pg = bool(db_url and asyncpg is not None)
        self.sqlite_path = os.getenv("SQLITE_PATH", "bot_data.db")
        self.pool = None
        self.sqlite_conn = None

    def _ph(self, query: str) -> str:
        if not self.use_pg:
            return query
        idx = 0
        out = []
        for ch in query:
            if ch == "?":
                idx += 1
                out.append(f"${idx}")
            else:
                out.append(ch)
        return "".join(out)

    async def init(self) -> None:
        if self.use_pg:
            self.pool = await asyncpg.create_pool(
                dsn=self.db_url,
                min_size=2,
                max_size=10,
                statement_cache_size=0,
                ssl="require",
            )
        else:
            self.sqlite_conn = sqlite3.connect(self.sqlite_path, check_same_thread=False)
            self.sqlite_conn.row_factory = sqlite3.Row
        await self._create_tables_and_seed()

    async def close(self) -> None:
        if self.pool:
            await self.pool.close()
        if self.sqlite_conn:
            self.sqlite_conn.close()

    async def execute(self, query: str, params: Tuple = ()) -> None:
        if self.use_pg:
            async with self.pool.acquire() as conn:
                await conn.execute(self._ph(query), *params)
        else:
            cur = self.sqlite_conn.cursor()
            cur.execute(query, params)
            self.sqlite_conn.commit()

    async def execute_insert_returning_id(self, query: str, params: Tuple = ()) -> int:
        if self.use_pg:
            async with self.pool.acquire() as conn:
                row = await conn.fetchrow(self._ph(query) + " RETURNING id", *params)
                return int(row["id"])
        else:
            cur = self.sqlite_conn.cursor()
            cur.execute(query, params)
            self.sqlite_conn.commit()
            return int(cur.lastrowid)

    async def fetchone(self, query: str, params: Tuple = ()) -> Optional[Dict[str, Any]]:
        if self.use_pg:
            async with self.pool.acquire() as conn:
                row = await conn.fetchrow(self._ph(query), *params)
                return dict(row) if row else None
        else:
            cur = self.sqlite_conn.cursor()
            cur.execute(query, params)
            row = cur.fetchone()
            return dict(row) if row else None

    async def fetchall(self, query: str, params: Tuple = ()) -> List[Dict[str, Any]]:
        if self.use_pg:
            async with self.pool.acquire() as conn:
                rows = await conn.fetch(self._ph(query), *params)
                return [dict(r) for r in rows]
        else:
            cur = self.sqlite_conn.cursor()
            cur.execute(query, params)
            rows = cur.fetchall()
            return [dict(r) for r in rows]

    async def _create_tables_and_seed(self) -> None:
        serial_pk = "SERIAL PRIMARY KEY" if self.use_pg else "INTEGER PRIMARY KEY AUTOINCREMENT"
        statements = [
            "CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL)",
            "CREATE TABLE IF NOT EXISTS admins (user_id BIGINT PRIMARY KEY, added_by BIGINT, created_at TEXT)",
            """CREATE TABLE IF NOT EXISTS users (
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
            )""",
            f"CREATE TABLE IF NOT EXISTS categories (id {serial_pk}, name TEXT NOT NULL UNIQUE, emoji TEXT DEFAULT '📸', is_active INTEGER DEFAULT 1)",
            f"""CREATE TABLE IF NOT EXISTS subcategories (
                id {serial_pk},
                category_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                rate REAL NOT NULL DEFAULT 5.0,
                steps_json TEXT NOT NULL,
                is_active INTEGER DEFAULT 1
            )""",
            "CREATE TABLE IF NOT EXISTS used_usernames (username TEXT PRIMARY KEY, created_at TEXT)",
            f"""CREATE TABLE IF NOT EXISTS recycled_usernames (
                id {serial_pk},
                username TEXT NOT NULL UNIQUE,
                full_name TEXT NOT NULL,
                last_user_id BIGINT NOT NULL,
                created_at TEXT
            )""",
            f"""CREATE TABLE IF NOT EXISTS activity_logs (
                id {serial_pk},
                user_id BIGINT NOT NULL,
                action_type TEXT NOT NULL,
                detail TEXT,
                created_at TEXT
            )""",
            "CREATE TABLE IF NOT EXISTS bot_groups (chat_id BIGINT PRIMARY KEY, title TEXT, added_at TEXT)",
            f"""CREATE TABLE IF NOT EXISTS batches (
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
            )""",
            f"""CREATE TABLE IF NOT EXISTS submissions (
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
            )""",
            f"""CREATE TABLE IF NOT EXISTS withdrawals (
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
            )""",
            f"""CREATE TABLE IF NOT EXISTS file_submit_categories (
                id {serial_pk},
                name TEXT NOT NULL UNIQUE,
                cutoff_time TEXT DEFAULT '10:00 PM',
                rate REAL DEFAULT 5.0,
                usdt_rate REAL DEFAULT 0.0394,
                required_format TEXT DEFAULT 'Uid pas cookies',
                report_col TEXT DEFAULT 'UID',
                is_active INTEGER DEFAULT 1
            )""",
            f"""CREATE TABLE IF NOT EXISTS file_submissions (
                id {serial_pk},
                user_id BIGINT NOT NULL,
                category_id INTEGER,
                category_name TEXT NOT NULL,
                rate REAL DEFAULT 5.0,
                total_rows INTEGER DEFAULT 0,
                file_id TEXT NOT NULL,
                file_name TEXT,
                status TEXT DEFAULT 'pending',
                created_at TEXT
            )""",
            f"""CREATE TABLE IF NOT EXISTS file_submit_rows (
                id {serial_pk},
                fsub_id INTEGER NOT NULL,
                user_id BIGINT NOT NULL,
                category_id INTEGER NOT NULL,
                category_name TEXT NOT NULL,
                rate REAL NOT NULL,
                identifier TEXT NOT NULL,
                raw_line TEXT,
                status TEXT DEFAULT 'pending',
                created_at TEXT
            )""",
            f"""CREATE TABLE IF NOT EXISTS work_admin_notifs (
                id {serial_pk},
                user_id BIGINT NOT NULL,
                sub_id INTEGER NOT NULL,
                admin_id BIGINT NOT NULL,
                message_id BIGINT NOT NULL
            )""",
            f"""CREATE TABLE IF NOT EXISTS guide_categories (
                id {serial_pk},
                name TEXT NOT NULL UNIQUE,
                guide_text TEXT,
                video_file_id TEXT,
                is_active INTEGER DEFAULT 1
            )""",
        ]
        for stmt in statements:
            await self.execute(stmt)

        # Safe column migrations for existing databases
        migrations = [
            "ALTER TABLE file_submit_categories ADD COLUMN cutoff_time TEXT DEFAULT '10:00 PM'",
            "ALTER TABLE file_submit_categories ADD COLUMN rate REAL DEFAULT 5.0",
            "ALTER TABLE file_submit_categories ADD COLUMN usdt_rate REAL DEFAULT 0.0394",
            "ALTER TABLE file_submit_categories ADD COLUMN required_format TEXT DEFAULT 'Uid pas cookies'",
            "ALTER TABLE file_submit_categories ADD COLUMN report_col TEXT DEFAULT 'UID'",
            "ALTER TABLE file_submissions ADD COLUMN category_id INTEGER",
            "ALTER TABLE file_submissions ADD COLUMN rate REAL DEFAULT 5.0",
            "ALTER TABLE file_submissions ADD COLUMN total_rows INTEGER DEFAULT 0",
        ]
        for m in migrations:
            try:
                await self.execute(m)
            except Exception:
                pass

        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        await self.execute(
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
            "password_prefix": "succes@",
            "password_reset_time": "00:00",
            "support_link": "https://t.me/Talha_juba098",
            "extra_buttons_json": "[]",
            "main_btns_json": json.dumps({
                "work": "💼 Work",
                "fsub": "📁 File Submit",
                "bal": "💰 Balance",
                "wd": "💸 Withdraw",
                "ref": "🎁 My Referral",
                "sup": "🆘 Support",
                "new": "🆕 I'm New"
            }, ensure_ascii=False),
            "msg_welcome": "⭐ Welcome to {bot_name}!\n\n⭐ এই Bot-এর মাধ্যমে আপনি খুব সহজেই আপনার Facebook / Instagram / অন্যান্য File Submit করতে পারবেন।\n\n⭐ File Submit করার আগে অবশ্যই সকল Rules অনুসরণ করুন।\n\n⭐ নিচের Menu থেকে আপনার প্রয়োজনীয় Option নির্বাচন করে File Submit করুন।",
            "msg_support": "💬 <b>Customer Support</b>\n\nDear member,\nContact our support team for any problem or question. We will try to solve it quickly.",
            "msg_ref_rules": "⚠️ <b>Referral Rules</b>\n\n• আপনার রেফারেল লিংক দিয়ে কেউ জয়েন করে কাজ করলে প্রতি কাজের জন্য আপনি <b>৳{bonus}</b> করে বোনাস পাবেন!\n• Fake/spam accounts may have commission cancelled.\n• You cannot refer yourself.",
            "force_join_active": "0",
            "force_join_channel": "",
            "force_join_link": "",
            "force_join_channels_json": "[]",
        }
        for k, v in defaults.items():
            await self.execute(
                "INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT DO NOTHING"
                if self.use_pg
                else "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
                (k, v),
            )

        existing_cat = await self.fetchone("SELECT id FROM categories LIMIT 1")
        if not existing_cat:
            cat_id = await self.execute_insert_returning_id(
                "INSERT INTO categories (name, emoji, is_active) VALUES (?, ?, ?)",
                ("Instagram", "📸", 1),
            )
            default_steps = [
                {
                    "name": "Hotmail Set",
                    "prompt": "সম্পূর্ণ হটমেইল ফুল টোকেনটি এখানে দিন:",
                    "validator": "hotmail_token",
                }
            ]
            await self.execute(
                "INSERT INTO subcategories (category_id, name, rate, steps_json, is_active) VALUES (?, ?, ?, ?, ?)",
                (cat_id, "USA Instagram", 5.0, json.dumps(default_steps, ensure_ascii=False), 1),
            )

        existing_fsub = await self.fetchone("SELECT id FROM file_submit_categories LIMIT 1")
        if not existing_fsub:
            await self.execute(
                "INSERT INTO file_submit_categories (name, cutoff_time, rate, usdt_rate, required_format, report_col, is_active) VALUES (?, ?, ?, ?, ?, ?, 1)",
                ("Fb cookies", "10:00 PM", 5.0, 0.0394, "Uid pas cookies", "UID"),
            )

        existing_guide = await self.fetchone("SELECT id FROM guide_categories LIMIT 1")
        if not existing_guide:
            for gname, gtext in [
                ("Instagram", "📸 Instagram কাজের নিয়ম:\n১. 💼 Work -> 📸 Instagram Work ▸ -> USA Instagram • ৳5 এ ক্লিক করুন।\n২. বটের দেওয়া নাম, ইউজারনেম ও পাসওয়ার্ড কপি করে একাউন্ট খুলুন।\n৩. Hotmail Set এ ক্লিক করে ফুল টোকেন সাবমিট করুন।"),
                ("Gmail", "📧 Gmail কাজের গাইডলাইন এখনো অ্যাড করা হয়নি।"),
                ("Facebook", "📘 Facebook কাজের গাইডলাইন এখনো অ্যাড করা হয়নি।"),
            ]:
                await self.execute(
                    "INSERT INTO guide_categories (name, guide_text, video_file_id, is_active) VALUES (?, ?, ?, ?)",
                    (gname, gtext, "", 1),
                )


db = AsyncDatabase(DATABASE_URL)

# ==============================================================================
# IN-MEMORY CACHE FOR ULTRA-FAST RESPONSES
# ==============================================================================
_SETTINGS_CACHE: Dict[str, str] = {}
_ADMINS_CACHE: Optional[set] = None
_USERS_CACHE: Dict[int, Dict[str, Any]] = {}
_CATS_CACHE: Optional[List[Dict[str, Any]]] = None
_SUBS_CACHE: Dict[int, List[Dict[str, Any]]] = {}
_FJOIN_OK_CACHE: Dict[int, float] = {}


async def warm_caches() -> None:
    global _ADMINS_CACHE, _CATS_CACHE
    rows = await db.fetchall("SELECT key, value FROM settings")
    for r in rows:
        _SETTINGS_CACHE[r["key"]] = r["value"]
    adm_rows = await db.fetchall("SELECT user_id FROM admins")
    _ADMINS_CACHE = {r["user_id"] for r in adm_rows} | {OWNER_ID}
    _CATS_CACHE = await db.fetchall("SELECT * FROM categories WHERE is_active = 1 ORDER BY id ASC")


async def get_setting(key: str, default: str = "") -> str:
    if key in _SETTINGS_CACHE:
        return _SETTINGS_CACHE[key]
    row = await db.fetchone("SELECT value FROM settings WHERE key = ?", (key,))
    val = row["value"] if row else default
    _SETTINGS_CACHE[key] = val
    return val


async def set_setting(key: str, value: str) -> None:
    _SETTINGS_CACHE[key] = value
    if db.use_pg:
        await db.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value",
            (key, value),
        )
    else:
        await db.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))


async def is_admin(user_id: int) -> bool:
    global _ADMINS_CACHE
    if user_id == OWNER_ID:
        return True
    if _ADMINS_CACHE is None:
        adm_rows = await db.fetchall("SELECT user_id FROM admins")
        _ADMINS_CACHE = {r["user_id"] for r in adm_rows} | {OWNER_ID}
    return user_id in _ADMINS_CACHE


async def get_all_admin_ids() -> List[int]:
    global _ADMINS_CACHE
    if _ADMINS_CACHE is None:
        adm_rows = await db.fetchall("SELECT user_id FROM admins")
        _ADMINS_CACHE = {r["user_id"] for r in adm_rows} | {OWNER_ID}
    return list(_ADMINS_CACHE)


async def ensure_user(update: Update) -> Dict[str, Any]:
    u = update.effective_user
    if u.id in _USERS_CACHE:
        return _USERS_CACHE[u.id]
    row = await db.fetchone("SELECT * FROM users WHERE user_id = ?", (u.id,))
    if not row:
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        await db.execute(
            "INSERT INTO users (user_id, username, full_name, joined_at) VALUES (?, ?, ?, ?)",
            (u.id, u.username or "", u.full_name or "", now_str),
        )
        row = await db.fetchone("SELECT * FROM users WHERE user_id = ?", (u.id,))
    _USERS_CACHE[u.id] = row
    return row


def invalidate_user_cache(user_id: int) -> None:
    _USERS_CACHE.pop(user_id, None)


def log_activity_bg(user_id: int, action_type: str, detail: str = "") -> None:
    now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
    asyncio.create_task(
        db.execute(
            "INSERT INTO activity_logs (user_id, action_type, detail, created_at) VALUES (?, ?, ?, ?)",
            (user_id, action_type, detail[:300], now_str),
        )
    )


async def get_active_categories() -> List[Dict[str, Any]]:
    global _CATS_CACHE
    if _CATS_CACHE is None:
        _CATS_CACHE = await db.fetchall("SELECT * FROM categories WHERE is_active = 1 ORDER BY id ASC")
    return _CATS_CACHE


async def get_active_subcategories(cat_id: int) -> List[Dict[str, Any]]:
    if cat_id not in _SUBS_CACHE:
        _SUBS_CACHE[cat_id] = await db.fetchall(
            "SELECT * FROM subcategories WHERE category_id = ? AND is_active = 1 ORDER BY id ASC",
            (cat_id,),
        )
    return _SUBS_CACHE[cat_id]


# ==============================================================================
# HELPERS: PASSWORD, SMART USERNAME POOL, VALIDATION, FORCE JOIN, EXCEL COLOR
# ==============================================================================
async def get_current_password() -> str:
    prefix = await get_setting("password_prefix", "succes@")
    reset_time_str = await get_setting("password_reset_time", "00:00")
    now_bd = datetime.datetime.now(BD_TZ)
    try:
        rh, rm = [int(x) for x in reset_time_str.split(":")]
    except Exception:
        rh, rm = 0, 0
    reset_dt_today = now_bd.replace(hour=rh, minute=rm, second=0, microsecond=0)
    effective_date = now_bd - datetime.timedelta(days=1) if now_bd < reset_dt_today else now_bd
    return f"{prefix}{effective_date.day}"


async def acquire_username_profile(user_id: int) -> Tuple[str, str]:
    rec = await db.fetchone(
        "SELECT * FROM recycled_usernames WHERE last_user_id != ? ORDER BY id ASC LIMIT 1",
        (user_id,),
    )
    if rec:
        await db.execute("DELETE FROM recycled_usernames WHERE id = ?", (rec["id"],))
        return rec["full_name"], rec["username"]

    for _ in range(15):
        fn = random.choice(FIRST_NAMES)
        ln = random.choice(LAST_NAMES)
        full_name = f"{fn} {ln}"
        uname = f"{fn.lower()}{ln.lower()}{random.randint(100, 9999)}{random.choice('abcdefghijklmnopqrstuvwxyz')}"
        exists = await db.fetchone("SELECT username FROM used_usernames WHERE username = ?", (uname,))
        if not exists:
            now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
            await db.execute("INSERT INTO used_usernames (username, created_at) VALUES (?, ?)", (uname, now_str))
            return full_name, uname
    return "Alex Carter", f"user{uuid.uuid4().hex[:10]}"


async def recycle_username_profile(full_name: str, username: str, user_id: int) -> None:
    now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
    await db.execute(
        "INSERT INTO recycled_usernames (username, full_name, last_user_id, created_at) VALUES (?, ?, ?, ?) ON CONFLICT DO NOTHING"
        if db.use_pg
        else "INSERT OR IGNORE INTO recycled_usernames (username, full_name, last_user_id, created_at) VALUES (?, ?, ?, ?)",
        (username, full_name, user_id, now_str),
    )


UUID_RE = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
HOTMAIL_DOMAINS = ("@hotmail.com", "@outlook.com", "@live.com", "@msn.com")


def validate_step_input(validator_type: str, text: str) -> Tuple[bool, str, str]:
    val = text.strip()
    if validator_type == "hotmail_token":
        parts = val.split("|")
        if len(parts) < 4:
            return False, "❌ ভুল ফরম্যাট! হটমেইল ফুল টোকেনটি অবশ্যই `email|password|refresh_token|client_id` থাকতে হবে।", ""
        email = parts[0].strip().lower()
        pwd = parts[1].strip()
        refresh_token = parts[2].strip()
        client_id = parts[3].strip()
        if not any(email.endswith(d) for d in HOTMAIL_DOMAINS) or "@" not in email:
            return False, "❌ ভ্যালিড Hotmail/Outlook ইমেইল পাওয়া যায়নি!", ""
        if len(pwd) < 3:
            return False, "❌ পাসওয়ার্ড অংশটি সঠিক নয়!", ""
        if not refresh_token.startswith("M.") or len(refresh_token) < 40:
            return False, "❌ Refresh Token সঠিক নয়! অনুগ্রহ করে সম্পূর্ণ ফুল টোকেন দিন।", ""
        if not UUID_RE.match(client_id):
            return False, "❌ Client ID (UUID) সঠিক নয়! সম্পূর্ণ ফুল টোকেনটি কপি করে দিন।", ""
        return True, "", email
    elif validator_type == "2fa":
        if len(val.replace(" ", "")) < 8:
            return False, "❌ ভ্যালিড 2FA Key দিন (কমপক্ষে ৮ অক্ষরের)।", ""
        return True, "", ""
    else:
        if len(val) < 2:
            return False, "❌ অনুগ্রহ করে সঠিক তথ্য দিন।", ""
        return True, "", ""


async def get_force_join_channels() -> List[Dict[str, str]]:
    raw = await get_setting("force_join_channels_json", "[]")
    try:
        chs = json.loads(raw)
        if chs:
            return chs
    except Exception:
        pass
    legacy_ch = await get_setting("force_join_channel", "")
    legacy_lk = await get_setting("force_join_link", "")
    if legacy_ch:
        return [{"channel": legacy_ch, "link": legacy_lk or f"https://t.me/{legacy_ch.lstrip('@')}", "label": "Join Channel"}]
    return []


async def check_force_join(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    uid = update.effective_user.id
    if await is_admin(uid):
        return True
    if await get_setting("force_join_active", "0") != "1":
        return True
    channels = await get_force_join_channels()
    if not channels:
        return True
    now_ts = time.time()
    if now_ts - _FJOIN_OK_CACHE.get(uid, 0) < 60:
        return True

    unjoined = []
    for ch_obj in channels:
        ch = ch_obj["channel"]
        try:
            member = await context.bot.get_chat_member(chat_id=ch, user_id=uid)
            if member.status not in ("member", "administrator", "creator"):
                unjoined.append(ch_obj)
        except Exception:
            continue

    if not unjoined:
        _FJOIN_OK_CACHE[uid] = now_ts
        return True

    btns = [[ikb_btn(f"📢 {c.get('label', 'Join Channel')}", url=c["link"], style="primary")] for c in unjoined]
    btns.append([ikb_btn("✅ Joined / Check Again", callback_data="fjoin:check", style="primary")])
    kb = InlineKeyboardMarkup(btns)
    if update.effective_message:
        await update.effective_message.reply_text(
            "📢 বটটি ব্যবহার করতে প্রথমে আমাদের অফিসিয়াল চ্যানেলে জয়েন করুন, তারপর নিচের <b>Joined</b> বাটনে চাপ দিন:",
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
    return False


async def get_main_btn_labels() -> Dict[str, str]:
    raw = await get_setting("main_btns_json", "{}")
    defaults = {
        "work": "💼 Work", "fsub": "📁 File Submit", "bal": "💰 Balance",
        "wd": "💸 Withdraw", "ref": "🎁 My Referral", "sup": "🆘 Support", "new": "🆕 I'm New"
    }
    try:
        loaded = json.loads(raw)
        defaults.update(loaded)
    except Exception:
        pass
    return defaults


async def main_keyboard(user_id: int) -> ReplyKeyboardMarkup:
    lbl = await get_main_btn_labels()
    extra_raw = await get_setting("extra_buttons_json", "[]")
    try:
        extra_btns = json.loads(extra_raw)
    except Exception:
        extra_btns = []

    rows = [
        [kb_btn(lbl["work"], style="success"), kb_btn(lbl["fsub"], style="primary")],
        [kb_btn(lbl["bal"], style="primary"), kb_btn(lbl["wd"], style="primary")],
        [kb_btn(lbl["ref"], style="primary"), kb_btn(lbl["sup"], style="primary")],
        [kb_btn(lbl["new"], style="primary")],
    ]
    for i in range(0, len(extra_btns), 2):
        rows.append([kb_btn(b["label"], style="primary") for b in extra_btns[i : i + 2]])
    if await is_admin(user_id):
        rows.append([kb_btn("⚙️ Admin Panel", style="primary")])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


async def work_categories_keyboard() -> ReplyKeyboardMarkup:
    cats = await get_active_categories()
    rows = [[KeyboardButton(f"{c.get('emoji', '📌')} {c['name']} Work ▸")] for c in cats]
    rows.append([kb_btn("❌ Cancel", style="danger")])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


async def work_subcategories_keyboard(cat_id: int) -> ReplyKeyboardMarkup:
    subs = await get_active_subcategories(cat_id)
    rows = [[kb_btn(f"{s['name']} • ৳{s['rate']:.0f}", style="primary")] for s in subs]
    rows.append([kb_btn("◀️ Work Menu ▸", style="danger"), kb_btn("❌ Cancel", style="danger")])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


async def admin_panel_keyboard() -> InlineKeyboardMarkup:
    bot_active = (await get_setting("bot_active", "1")) == "1"
    fj_active = (await get_setting("force_join_active", "0")) == "1"
    status_label = "🟢 Bot Status: ON" if bot_active else "🔴 Bot Status: OFF"
    fj_label = "📢 Force Join: ON" if fj_active else "📢 Force Join: OFF"
    kb = [
        [
            ikb_btn(status_label, callback_data="adm:toggle_bot", style="primary"),
            ikb_btn(fj_label, callback_data="adm:fjoin_mgr", style="primary"),
        ],
        [
            ikb_btn("📂 Pending Work", callback_data="adm:pending_files", style="success"),
            ikb_btn("📤 Work Report (.xlsx)", callback_data="adm:upload_report", style="success"),
        ],
        [
            ikb_btn("📊 File Submit Report (.xlsx)", callback_data="adm:fsub_report_pick", style="primary"),
        ],
        [
            ikb_btn("👁 View User Inbox", callback_data="adm:view_inbox_ask", style="primary"),
            ikb_btn("📩 Send to User Inbox", callback_data="adm:inbox_upload", style="primary"),
        ],
        [
            ikb_btn("🛠 Work & Steps", callback_data="adm:work_mgr", style="success"),
            ikb_btn("💰 Balance & Withdraw", callback_data="adm:bal_mgr", style="success"),
        ],
        [
            ikb_btn("🔑 Password Settings", callback_data="adm:pwd_mgr", style="primary"),
            ikb_btn("⌨️ Main Buttons", callback_data="adm:btn_mgr", style="primary"),
        ],
        [
            ikb_btn("🎬 আমি নতুন (Guides)", callback_data="adm:guide_mgr", style="primary"),
            ikb_btn("📁 File Submit Cats", callback_data="adm:fsub_mgr", style="success"),
        ],
        [
            ikb_btn("👥 User Directory", callback_data="adm:user_dir:0", style="primary"),
            ikb_btn("📜 File History Report", callback_data="adm:file_history", style="primary"),
        ],
        [
            ikb_btn("🗑 Delete User Data", callback_data="adm:del_user", style="danger"),
            ikb_btn("👮 Manage Admins", callback_data="adm:admins", style="primary"),
        ],
        [
            ikb_btn("📢 Broadcast", callback_data="adm:broadcast", style="primary"),
            ikb_btn("📝 Edit Bot Messages", callback_data="adm:edit_msgs", style="primary"),
        ],
        [
            ikb_btn("🆘 Edit Support Link", callback_data="adm:edit_support", style="primary"),
            ikb_btn("❌ Close Panel", callback_data="adm:close", style="primary"),
        ],
    ]
    return InlineKeyboardMarkup(kb)


def back_to_admin_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")]])


def classify_excel_color(cell) -> Optional[str]:
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
        r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
    except ValueError:
        return None
    if g > r + 25 and g > b + 20:
        return "green"
    if r > g + 30 and r > b + 30:
        return "red"
    return None


def _validate_token_part(field: str, val: str) -> bool:
    f = field.lower()
    v = val.strip()
    if not v:
        return False
    if f in ("uid", "username", "id"):
        return len(v) >= 3
    if f in ("pas", "pass", "password"):
        return len(v) >= 2
    if f in ("cookies", "cookie"):
        return ("c_user=" in v and "xs=" in v) or ("sessionid=" in v) or len(v) >= 20
    if f in ("2fa", "twofa"):
        return len(v.replace(" ", "")) >= 8
    if f in ("hotmail", "outlook", "fr_outlook", "fr", "recovery", "email", "mail"):
        return "@" in v and "." in v
    return len(v) >= 2


def parse_and_validate_fsub_excel(bio: io.BytesIO, required_format: str) -> Tuple[List[Tuple[str, str]], str]:
    try:
        wb = openpyxl.load_workbook(bio, data_only=True)
    except Exception:
        return [], "❌ ফাইলটি পড়া যায়নি! সঠিক <b>.xlsx</b> ফাইল দিন।"
    ws = wb.active
    fields = [x.strip().lower() for x in re.split(r"[\s|,]+", required_format.strip()) if x.strip()]
    if not fields:
        fields = ["uid"]

    raw_rows = []
    for r in ws.iter_rows(values_only=True):
        vals = [str(c).strip() if c is not None else "" for c in r]
        if any(vals):
            raw_rows.append(vals)

    if not raw_rows:
        return [], "❌ এক্সেল ফাইলটি খালি!"

    first_cell_low = raw_rows[0][0].lower()
    if first_cell_low in ("uid", "id", "username", "sl", "account", "email"):
        raw_rows = raw_rows[1:]
    if not raw_rows:
        return [], "❌ ফাইলে কোনো অ্যাকাউন্টের সারি পাওয়া যায়নি!"

    parsed: List[Tuple[str, str]] = []
    for row_idx, vals in enumerate(raw_rows):
        non_empty = [v for v in vals if v != ""]
        if len(non_empty) == 1 and ("|" in non_empty[0] or " " in non_empty[0]) and len(fields) > 1:
            sep = "|" if "|" in non_empty[0] else None
            parts = [x.strip() for x in (non_empty[0].split("|") if sep else non_empty[0].split(None, len(fields) - 1))]
        else:
            parts = vals[:len(fields)]
            while len(parts) < len(fields):
                parts.append("")

        for fi, fname in enumerate(fields):
            val = parts[fi] if fi < len(parts) else ""
            if fname in ("pas", "pass", "password"):
                if row_idx == 0 and not _validate_token_part(fname, val):
                    return [], f"❌ প্রথম সারিতে <b>Password</b> বাধ্যতামূলক! অনুগ্রহ করে ফরম্যাট <code>{required_format}</code> অনুযায়ী দিন।"
                continue
            if not _validate_token_part(fname, val):
                return [], f"❌ সারি {row_idx + 1}-এ <b>{fname}</b> অংশটি সঠিক নয়! ফরম্যাট: <code>{required_format}</code>"

        identifier = parts[0].strip()
        raw_line = " | ".join(parts)
        parsed.append((identifier, raw_line))

    return parsed, ""


# ==============================================================================
# START & ADMIN COMMANDS
# ==============================================================================
async def start_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_chat.type in ("group", "supergroup"):
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        await db.execute(
            "INSERT INTO bot_groups (chat_id, title, added_at) VALUES (?, ?, ?) ON CONFLICT (chat_id) DO UPDATE SET title = EXCLUDED.title"
            if db.use_pg
            else "INSERT OR REPLACE INTO bot_groups (chat_id, title, added_at) VALUES (?, ?, ?)",
            (update.effective_chat.id, update.effective_chat.title or "Group", now_str),
        )
        return

    context.user_data.clear()
    u = update.effective_user
    existing = await db.fetchone("SELECT user_id FROM users WHERE user_id = ?", (u.id,))
    await ensure_user(update)

    if not existing and context.args:
        ref_arg = context.args[0].strip()
        if ref_arg.isdigit():
            ref_id = int(ref_arg)
            if ref_id != u.id:
                ref_user = await db.fetchone("SELECT user_id FROM users WHERE user_id = ?", (ref_id,))
                if ref_user:
                    await db.execute("UPDATE users SET referred_by = ? WHERE user_id = ?", (ref_id, u.id))
                    await db.execute("UPDATE users SET referral_count = referral_count + 1 WHERE user_id = ?", (ref_id,))
                    invalidate_user_cache(u.id)
                    invalidate_user_cache(ref_id)

    if await get_setting("bot_active", "1") != "1" and not await is_admin(u.id):
        await update.message.reply_text("🔧 বটের আপডেটের কাজ চলতেছে। অনুগ্রহ করে কিছুক্ষণ পর আবার চেষ্টা করুন।")
        return

    if not await check_force_join(update, context):
        return

    log_activity_bg(u.id, "START", "/start")
    bot_name = context.bot.first_name or "Fast Account Bot"
    tpl = await get_setting("msg_welcome", "⭐ Welcome to {bot_name}!")
    welcome_text = tpl.replace("{bot_name}", f"<b>{bot_name}</b>")
    await update.message.reply_text(
        welcome_text,
        parse_mode=ParseMode.HTML,
        reply_markup=await main_keyboard(u.id),
    )


async def admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await is_admin(update.effective_user.id):
        return
    context.user_data.clear()
    await update.message.reply_text(
        "⚙️ <b>Admin Control Panel</b>\nনিচ থেকে যেকোনো অপশন সিলেক্ট করুন:",
        parse_mode=ParseMode.HTML,
        reply_markup=await admin_panel_keyboard(),
    )


# ==============================================================================
# WORK & TASK FLOW (REPLY KEYBOARD BASED)
# ==============================================================================
async def handle_work_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    cats = await get_active_categories()
    if not cats:
        await update.message.reply_text("📋 এই মুহূর্তে কোনো কাজের ক্যাটাগরি চালু নেই।")
        return
    context.user_data["nav"] = "work_cats"
    await update.message.reply_text("📝 Select work type:", reply_markup=await work_categories_keyboard())


async def send_next_step_prompt(chat_id: int, context: ContextTypes.DEFAULT_TYPE, task: Dict[str, Any]) -> None:
    steps = task["steps"]
    idx = task["step_idx"]
    step = steps[idx]
    kb = ReplyKeyboardMarkup([[kb_btn(f"🔑 {step['name']}", style="success")], [kb_btn("❌ Cancel", style="danger")]], resize_keyboard=True)
    msg = await context.bot.send_message(
        chat_id=chat_id,
        text=f"🔹 ধাপ {idx + 1}/{len(steps)}: নিচের <b>{step['name']}</b> বাটনে ট্যাপ করুন অথবা Cancel করুন।",
        parse_mode=ParseMode.HTML,
        reply_markup=kb,
    )
    task["msg_ids"].append(msg.message_id)
    task["awaiting_button"] = True


async def handle_task_text(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str) -> bool:
    task = context.user_data.get("active_task")
    if not task:
        return False

    uid = update.effective_user.id
    chat_id = update.effective_chat.id
    task["msg_ids"].append(update.message.message_id)

    if text in ("❌ Cancel", "Cancel"):
        await recycle_username_profile(task["gen_name"], task["gen_username"], uid)
        log_activity_bg(uid, "CANCEL_TASK", f"Cancelled @{task['gen_username']} (recycled)")
        for mid in task["msg_ids"]:
            try:
                await context.bot.delete_message(chat_id=chat_id, message_id=mid)
            except Exception:
                pass
        context.user_data.pop("active_task", None)
        context.user_data.pop("nav", None)
        await update.message.reply_text(
            f"🗑️ Cancelled. Username <code>{task['gen_username']}</code> returned to pool.",
            parse_mode=ParseMode.HTML,
            reply_markup=await main_keyboard(uid),
        )
        return True

    steps = task["steps"]
    idx = task["step_idx"]
    current_step = steps[idx]
    btn_label = f"🔑 {current_step['name']}"

    if task.get("awaiting_button"):
        if text in (btn_label, current_step["name"]):
            task["awaiting_button"] = False
            pmsg = await update.message.reply_text(current_step["prompt"])
            task["msg_ids"].append(pmsg.message_id)
            return True
        wmsg = await update.message.reply_text("⚠️ Finish or cancel your current task first.")
        task["msg_ids"].append(wmsg.message_id)
        return True

    ok, err_msg, fingerprint = validate_step_input(current_step.get("validator", "text"), text)
    if not ok:
        emsg = await update.message.reply_text(err_msg)
        task["msg_ids"].append(emsg.message_id)
        return True

    if fingerprint:
        task["token_fingerprint"] = fingerprint

    task["collected"][current_step["name"]] = text.strip()
    task["step_idx"] += 1

    if task["step_idx"] < len(steps):
        await send_next_step_prompt(chat_id, context, task)
        return True

    for mid in task["msg_ids"]:
        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=mid)
        except Exception:
            pass

    cat_id = task.get("cat_id")
    ret_kb = await work_subcategories_keyboard(cat_id) if cat_id else await main_keyboard(uid)
    await context.bot.send_message(chat_id=chat_id, text="⏳ Verifying submission...")

    now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
    sub_id = await db.execute_insert_returning_id(
        """
        INSERT INTO submissions (
            user_id, subcategory_id, subcategory_name, rate,
            gen_name, gen_username, gen_password, steps_data_json,
            token_fingerprint, status, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?)
        """,
        (
            uid,
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
    log_activity_bg(uid, "SUBMIT_TASK", f"#{sub_id} {task['sub_name']} (@{task['gen_username']})")
    sub_id_for_btn = task["sub_id"]
    context.user_data.pop("active_task", None)
    await context.bot.send_message(
        chat_id=chat_id,
        text="✅ Submitted! Sent to admin review.",
        reply_markup=ret_kb,
    )

    pending_cnt = (
        await db.fetchone("SELECT COUNT(*) AS cnt FROM submissions WHERE user_id = ? AND status = 'pending'", (uid,))
    )["cnt"]
    u = update.effective_user
    recv_kb = InlineKeyboardMarkup([
        [ikb_btn(f"📥 Receive ({pending_cnt}টি)", callback_data=f"adm:recv:{u.id}:{sub_id_for_btn}", style="success")]
    ])
    notif_text = (
        f"📥 <b>নতুন সাবমিশন!</b>\n"
        f"👤 ইউজার: {u.full_name} (<code>{u.id}</code>)\n"
        f"📂 কাজ: {task['sub_name']} (৳{task['rate']:.0f})\n"
        f"📦 মোট Pending: <b>{pending_cnt}টি</b>"
    )
    existing_notifs = await db.fetchall(
        "SELECT * FROM work_admin_notifs WHERE user_id = ? AND sub_id = ?",
        (uid, sub_id_for_btn),
    )
    if existing_notifs:
        for n in existing_notifs:
            try:
                await context.bot.edit_message_text(
                    chat_id=n["admin_id"],
                    message_id=n["message_id"],
                    text=notif_text,
                    parse_mode=ParseMode.HTML,
                    reply_markup=recv_kb,
                )
            except Exception:
                pass
    else:
        for adm_id in await get_all_admin_ids():
            try:
                sent_m = await context.bot.send_message(
                    chat_id=adm_id,
                    text=notif_text,
                    parse_mode=ParseMode.HTML,
                    reply_markup=recv_kb,
                )
                await db.execute(
                    "INSERT INTO work_admin_notifs (user_id, sub_id, admin_id, message_id) VALUES (?, ?, ?, ?)",
                    (uid, sub_id_for_btn, adm_id, sent_m.message_id),
                )
            except Exception:
                pass
    return True


# ==============================================================================
# MENU HANDLERS
# ==============================================================================
async def handle_balance(update: Update) -> None:
    uid = update.effective_user.id
    invalidate_user_cache(uid)
    u = await ensure_user(update)
    in_review = (
        await db.fetchone("SELECT COUNT(*) AS cnt FROM submissions WHERE user_id = ? AND status IN ('pending', 'received')", (uid,))
    )["cnt"]
    text = (
        "💰 <b>Your Balance</b>\n\n"
        f"💵 <b>Total Balance:</b> ৳{u['balance']:.2f}\n"
        f"⏳ <b>Pending Withdraw:</b> ৳{u['pending_withdraw']:.2f}\n"
        f"📈 <b>Total Income:</b> ৳{u['total_income']:.2f}\n"
        f"✅ <b>Completed Jobs:</b> {u['completed_jobs']}\n"
        f"🔍 <b>In Review Jobs:</b> {in_review}"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.HTML)


async def handle_referral(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    uid = update.effective_user.id
    invalidate_user_cache(uid)
    u = await ensure_user(update)
    link = f"https://t.me/{context.bot.username}?start={uid}"
    text = (
        "🎁 <b>Refer & Earn</b>\n\n"
        f"👥 <b>Total Refer:</b> {u['referral_count']}\n"
        f"💵 <b>Total Refer Income:</b> ৳{u['referral_earnings']:.2f}\n\n"
        f"🔗 <b>Your referral link:</b>\n<code>{link}</code>"
    )
    kb = InlineKeyboardMarkup([
        [ikb_btn("📤 Forward This Message", switch_inline_query=f"Join & earn easily! {link}", style="primary")],
        [
            ikb_btn("🏆 Team Leaderboard", callback_data="ref:leaderboard", style="primary"),
            ikb_btn("⚠️ Rules", callback_data="ref:rules", style="primary"),
        ],
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)


async def handle_withdraw_start(update: Update) -> None:
    uid = update.effective_user.id
    invalidate_user_cache(uid)
    u = await ensure_user(update)
    min_w = float(await get_setting("min_withdraw", "50"))
    fee = float(await get_setting("withdraw_fee", "5"))
    if u["balance"] < min_w:
        await update.message.reply_text(
            f"❌ আপনার ব্যালেন্সে পর্যাপ্ত টাকা নেই।\nসর্বনিম্ন উইথড্র: <b>৳{min_w:.0f}</b> (বর্তমান ব্যালেন্স: ৳{u['balance']:.2f})",
            parse_mode=ParseMode.HTML,
        )
        return
    kb = InlineKeyboardMarkup([
        [
            ikb_btn("bKash", callback_data="wd:method:bKash", style="primary"),
            ikb_btn("Nagad", callback_data="wd:method:Nagad", style="primary"),
            ikb_btn("Binance", callback_data="wd:method:Binance", style="primary"),
        ],
        [ikb_btn("❌ Cancel", callback_data="wd:cancel", style="danger")],
    ])
    await update.message.reply_text(
        f"💸 <b>Choose a withdraw method (min ৳{min_w:.0f})</b>\n"
        f"• bKash — ৳{fee:.0f} gateway fee\n"
        f"• Nagad — ৳{fee:.0f} gateway fee\n"
        f"• Binance — ৳{fee:.0f} gateway fee",
        parse_mode=ParseMode.HTML,
        reply_markup=kb,
    )


async def handle_file_submit_menu(update: Update) -> None:
    cats = await db.fetchall("SELECT * FROM file_submit_categories WHERE is_active = 1 ORDER BY id ASC")
    if not cats:
        await update.message.reply_text("👤 No file submit categories are active right now.")
        return
    rows = [[kb_btn(f"📁 {c['name']} Submit ▸", style="success")] for c in cats]
    rows.append([kb_btn("❌ Cancel", style="danger")])
    await update.message.reply_text(
        "📁 আপনি কোন ক্যাটাগরির একাউন্ট সাবমিট করতে চান তা নির্বাচন করুন:",
        reply_markup=ReplyKeyboardMarkup(rows, resize_keyboard=True),
    )


async def handle_guide_menu(update: Update) -> None:
    guides = await db.fetchall("SELECT * FROM guide_categories WHERE is_active = 1 ORDER BY id ASC")
    if not guides:
        await update.message.reply_text("ℹ️ এই মুহূর্তে কোনো গাইডলাইন যুক্ত নেই।")
        return
    kb = [[ikb_btn(f"🎬 {g['name']}", callback_data=f"guide:view:{g['id']}", style="primary")] for g in guides]
    await update.message.reply_text("🆕 আপনি কোন কাজের নিয়ম ও ভিডিও দেখতে চান?", reply_markup=InlineKeyboardMarkup(kb))


# ==============================================================================
# CALLBACK ROUTER
# ==============================================================================
async def callback_router(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    data = query.data
    uid = query.from_user.id

    if data == "fjoin:check":
        _FJOIN_OK_CACHE.pop(uid, None)
        if await check_force_join(update, context):
            await query.edit_message_text("✅ ধন্যবাদ! চ্যানেল জয়েন ভেরিফাই হয়েছে। এখন /start দিন।")
        return

    if await get_setting("bot_active", "1") != "1" and not await is_admin(uid):
        await query.message.reply_text("🔧 বটের আপডেটের কাজ চলতেছে।")
        return

    if data == "ref:rules":
        bonus = float(await get_setting("referral_bonus", "3"))
        rtpl = await get_setting("msg_ref_rules", "⚠️ <b>Referral Rules</b>\n\n• প্রতি কাজের জন্য আপনি ৳{bonus} করে বোনাস পাবেন!")
        await query.message.reply_text(rtpl.replace("{bonus}", f"{bonus:.0f}"), parse_mode=ParseMode.HTML)
        return

    if data == "ref:leaderboard":
        top = await db.fetchall("SELECT full_name, referral_count, referral_earnings FROM users WHERE referral_count > 0 ORDER BY referral_count DESC LIMIT 10")
        if not top:
            await query.message.reply_text("🏆 এখনো কোনো রেফারেল লিডারবোর্ড তৈরি হয়নি।")
            return
        lines = ["🏆 <b>Team Leaderboard</b>\n"] + [
            f"{i+1}. {r['full_name']} — {r['referral_count']} refers (৳{r['referral_earnings']:.0f})"
            for i, r in enumerate(top)
        ]
        await query.message.reply_text("\n".join(lines), parse_mode=ParseMode.HTML)
        return

    if data.startswith("guide:view:"):
        gid = int(data.split(":")[2])
        g = await db.fetchone("SELECT * FROM guide_categories WHERE id = ?", (gid,))
        if not g:
            return
        if g.get("video_file_id"):
            await query.message.reply_video(video=g["video_file_id"], caption=g.get("guide_text") or g["name"])
        else:
            await query.message.reply_text(g.get("guide_text") or "গাইডলাইন পাওয়া যায়নি।")
        return

    if data.startswith("fsub:cat:"):
        cid = int(data.split(":")[2])
        cat = await db.fetchone("SELECT * FROM file_submit_categories WHERE id = ?", (cid,))
        if not cat:
            return
        context.user_data["state"] = f"user_fsub:{cat['id']}"
        await query.edit_message_text(f"📤 অনুগ্রহ করে <b>{cat['name']}</b>-এর জন্য আপনার ফাইলটি এখানে পাঠান:", parse_mode=ParseMode.HTML)
        return

    if data == "wd:cancel":
        context.user_data.pop("wd_data", None)
        context.user_data.pop("state", None)
        await query.edit_message_text("❌ উইথড্র বাতিল করা হয়েছে।")
        return

    if data.startswith("wd:method:"):
        method = data.split(":")[2]
        context.user_data["wd_data"] = {"method": method}
        context.user_data["state"] = "wd_account"
        prompt = "আপনার Binance Pay ID / UID দিন:" if method == "Binance" else f"আপনার {method} নম্বরটি দিন:"
        await query.edit_message_text(f"💳 মেথড: <b>{method}</b>\n{prompt}", parse_mode=ParseMode.HTML)
        return

    if data.startswith("wd_adm:"):
        if not await is_admin(uid):
            return
        _, action, wd_id_str = data.split(":")
        wd_id = int(wd_id_str)
        wd = await db.fetchone("SELECT * FROM withdrawals WHERE id = ?", (wd_id,))
        if not wd or wd["status"] != "pending":
            await query.edit_message_text("ℹ️ এই উইথড্র রিকোয়েস্ট আগেই প্রসেস করা হয়েছে।")
            return
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        if action == "paid":
            await db.execute("UPDATE withdrawals SET status = 'paid', processed_at = ? WHERE id = ?", (now_str, wd_id))
            await db.execute("UPDATE users SET pending_withdraw = MAX(0, pending_withdraw - ?) WHERE user_id = ?", (wd["requested_amount"], wd["user_id"]))
            invalidate_user_cache(wd["user_id"])
            await query.edit_message_text(
                f"✅ <b>Paid!</b> \n👤 UID: <code>{wd['user_id']}</code> | {wd['method']}: <code>{wd['account_number']}</code>\n💵 পাঠানো হয়েছে: ৳{wd['net_amount']:.2f}",
                parse_mode=ParseMode.HTML,
            )
            try:
                await context.bot.send_message(
                    chat_id=wd["user_id"],
                    text=f"✅ আপনার ৳{wd['net_amount']:.2f} ({wd['method']} - {wd['account_number']}) উইথড্র পেমেন্ট সম্পন্ন হয়েছে!",
                )
            except Exception:
                pass
        elif action == "reject":
            await db.execute("UPDATE withdrawals SET status = 'rejected', processed_at = ? WHERE id = ?", (now_str, wd_id))
            await db.execute(
                "UPDATE users SET balance = balance + ?, pending_withdraw = MAX(0, pending_withdraw - ?) WHERE user_id = ?",
                (wd["requested_amount"], wd["requested_amount"], wd["user_id"]),
            )
            invalidate_user_cache(wd["user_id"])
            await query.edit_message_text("❌ <b>Rejected & Refunded</b>", parse_mode=ParseMode.HTML)
            try:
                await context.bot.send_message(
                    chat_id=wd["user_id"],
                    text=f"❌ আপনার ৳{wd['requested_amount']:.2f} উইথড্র রিকোয়েস্ট রিজেক্ট করা হয়েছে এবং টাকা আপনার ব্যালেন্সে ফেরত দেওয়া হয়েছে।",
                )
            except Exception:
                pass
        return

    if data.startswith("adm:"):
        if not await is_admin(uid):
            return
        await handle_admin_callback(query, context, data)


# ==============================================================================
# ADMIN CALLBACK HANDLER
# ==============================================================================
async def handle_admin_callback(query, context: ContextTypes.DEFAULT_TYPE, data: str) -> None:
    global _CATS_CACHE, _ADMINS_CACHE
    parts = data.split(":")
    action = parts[1]

    if action == "close":
        context.user_data.pop("state", None)
        try:
            await query.message.delete()
        except Exception:
            await query.edit_message_text("✅ Admin Panel closed.")
        return

    if action == "edit_msgs":
        kb = [
            [ikb_btn("👋 Edit Welcome Message", callback_data="adm:emsg:msg_welcome", style="primary")],
            [ikb_btn("💬 Edit Support Message", callback_data="adm:emsg:msg_support", style="primary")],
            [ikb_btn("⚠️ Edit Referral Rules", callback_data="adm:emsg:msg_ref_rules", style="primary")],
            [ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")],
        ]
        await query.edit_message_text("📝 <b>Edit Bot Messages</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "emsg":
        mkey = parts[2]
        context.user_data["state"] = f"adm_emsg:{mkey}"
        cur_val = await get_setting(mkey, "")
        await query.edit_message_text(
            f"বর্তমান মেসেজ:\n<code>{cur_val}</code>\n\nনতুন মেসেজটি লিখে পাঠান:",
            parse_mode=ParseMode.HTML,
            reply_markup=back_to_admin_kb(),
        )
        return

    if action == "home":
        context.user_data.pop("state", None)
        await query.edit_message_text(
            "⚙️ <b>Admin Control Panel</b>\nনিচ থেকে যেকোনো অপশন সিলেক্ট করুন:",
            parse_mode=ParseMode.HTML,
            reply_markup=await admin_panel_keyboard(),
        )
        return

    if action == "toggle_bot":
        cur = await get_setting("bot_active", "1")
        await set_setting("bot_active", "0" if cur == "1" else "1")
        await query.edit_message_reply_markup(reply_markup=await admin_panel_keyboard())
        return

    if action == "fjoin_mgr":
        fj_active = (await get_setting("force_join_active", "0")) == "1"
        channels = await get_force_join_channels()
        kb = [
            [ikb_btn(f"{'🟢 Turn OFF' if fj_active else '🔴 Turn ON'} Force Join", callback_data="adm:fjoin_tog", style="primary")],
        ]
        for idx, c in enumerate(channels):
            kb.append([
                ikb_btn(f"✏️ {c.get('label', c['channel'])}", callback_data=f"adm:fjoin_edit:{idx}", style="primary"),
                ikb_btn("🗑 Delete", callback_data=f"adm:fjoin_del:{idx}", style="danger"),
            ])
        kb.append([ikb_btn("➕ Add Channel Link", callback_data="adm:fjoin_add", style="primary")])
        kb.append([ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")])
        ch_txt = "\n".join([f"{i+1}. <b>{c.get('label','Channel')}</b>: <code>{c['channel']}</code>" for i, c in enumerate(channels)]) or "কোনো চ্যানেল যুক্ত নেই।"
        await query.edit_message_text(
            f"📢 <b>Force Join Channels ({len(channels)}টি)</b>\n\n{ch_txt}",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if action == "fjoin_tog":
        cur = await get_setting("force_join_active", "0")
        await set_setting("force_join_active", "0" if cur == "1" else "1")
        await handle_admin_callback(query, context, "adm:fjoin_mgr")
        return

    if action == "fjoin_add":
        context.user_data["state"] = "adm_fjoin_add"
        await query.edit_message_text(
            "নতুন চ্যানেল এই ফরম্যাটে লিখুন:\n<code>@channel_username | https://t.me/link | Button Name</code>",
            parse_mode=ParseMode.HTML,
            reply_markup=back_to_admin_kb(),
        )
        return

    if action == "fjoin_edit":
        idx = int(parts[2])
        context.user_data["state"] = f"adm_fjoin_edit:{idx}"
        await query.edit_message_text(
            "নতুন তথ্য এই ফরম্যাটে লিখুন:\n<code>@channel_username | https://t.me/link | Button Name</code>",
            parse_mode=ParseMode.HTML,
            reply_markup=back_to_admin_kb(),
        )
        return

    if action == "fjoin_del":
        idx = int(parts[2])
        channels = await get_force_join_channels()
        if 0 <= idx < len(channels):
            channels.pop(idx)
            await set_setting("force_join_channels_json", json.dumps(channels, ensure_ascii=False))
        await handle_admin_callback(query, context, "adm:fjoin_mgr")
        return

    if action == "view_inbox_ask":
        context.user_data["state"] = "adm_view_inbox_uid"
        await query.edit_message_text("👁 যে ইউজারের ইনবক্স ও লাইভ অ্যাক্টিভিটি দেখতে চান তার <b>UID</b> লিখুন:", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "view_inbox":
        tuid = int(parts[2])
        usr = await db.fetchone("SELECT * FROM users WHERE user_id = ?", (tuid,))
        logs = await db.fetchall("SELECT * FROM activity_logs WHERE user_id = ? ORDER BY id DESC LIMIT 20", (tuid,))
        header = f"👁 <b>User Inbox & Activity</b>\n👤 {usr['full_name'] if usr else tuid} (<code>{tuid}</code>)\n\n"
        log_lines = [f"• [{l['created_at'][11:16]}] <b>{l['action_type']}</b>: {l['detail']}" for l in reversed(logs)] or ["কোনো অ্যাক্টিভিটি নেই।"]
        kb = [
            [ikb_btn("👤 Open Telegram Chat", url=f"tg://user?id={tuid}", style="primary")],
            [ikb_btn("📩 Send Message / File to Inbox", callback_data=f"adm:inbox_direct:{tuid}", style="primary")],
            [ikb_btn("🔄 Refresh", callback_data=f"adm:view_inbox:{tuid}", style="primary")],
            [ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")],
        ]
        await query.edit_message_text(header + "\n".join(log_lines), parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "inbox_direct":
        context.user_data["inbox_target_uid"] = int(parts[2])
        context.user_data["state"] = "adm_inbox_file"
        await query.edit_message_text("এবার ইউজারের ইনবক্সে পাঠানোর জন্য ফাইল, ছবি বা মেসেজ দিন:", reply_markup=back_to_admin_kb())
        return

    if action == "pending_files":
        rows = await db.fetchall(
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
        kb = [
            [ikb_btn(f"📥 Receive: {r['full_name'] or r['user_id']} ({r['subcategory_name']} - {r['cnt']}টি)", callback_data=f"adm:recv:{r['user_id']}:{r['subcategory_id']}", style="success")]
            for r in rows
        ]
        kb.append([ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")])
        await query.edit_message_text("📂 <b>Pending Submissions</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "recv":
        target_uid, sub_id = int(parts[2]), int(parts[3])
        subs = await db.fetchall(
            "SELECT * FROM submissions WHERE user_id = ? AND subcategory_id = ? AND status = 'pending' ORDER BY id ASC",
            (target_uid, sub_id),
        )
        if not subs:
            await query.edit_message_text("❌ কোনো পেন্ডিং কাজ পাওয়া যায়নি।", reply_markup=back_to_admin_kb())
            return

        sub_meta = await db.fetchone("SELECT * FROM subcategories WHERE id = ?", (sub_id,))
        step_names = [s["name"] for s in json.loads(sub_meta["steps_json"])] if sub_meta else []
        for s in subs:
            for k in json.loads(s["steps_data_json"]).keys():
                if k not in step_names:
                    step_names.append(k)

        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        batch_id = await db.execute_insert_returning_id(
            "INSERT INTO batches (user_id, subcategory_id, subcategory_name, rate, total_count, status, received_at) VALUES (?, ?, ?, ?, ?, 'received', ?)",
            (target_uid, sub_id, subs[0]["subcategory_name"], subs[0]["rate"], len(subs), now_str),
        )
        await db.execute(
            "UPDATE submissions SET status = 'received', batch_id = ? WHERE user_id = ? AND subcategory_id = ? AND status = 'pending'",
            (batch_id, target_uid, sub_id),
        )
        notifs = await db.fetchall("SELECT * FROM work_admin_notifs WHERE user_id = ? AND sub_id = ?", (target_uid, sub_id))
        for n in notifs:
            if n["message_id"] != query.message.message_id:
                try:
                    await context.bot.delete_message(chat_id=n["admin_id"], message_id=n["message_id"])
                except Exception:
                    pass
        await db.execute("DELETE FROM work_admin_notifs WHERE user_id = ? AND sub_id = ?", (target_uid, sub_id))

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
            row_vals = [idx, s["id"], s["gen_username"], s["gen_password"]] + [sdata.get(st, "") for st in step_names] + ["PENDING"]
            ws.append(row_vals)

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 12), 45)

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        await query.message.reply_document(
            document=bio,
            filename=f"batch_{batch_id}_user_{target_uid}.xlsx",
            caption=f"📊 <b>Received!</b> ({len(subs)}টি)\nসঠিক সারিতে <b>সবুজ</b> ও ভুলে <b>লাল</b> কালার দিয়ে <b>Upload Report (.xlsx)</b> করুন।",
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
        try:
            await query.message.delete()
        except Exception:
            await query.edit_message_text("✅ রিসিভ সম্পন্ন হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if action == "upload_report":
        context.user_data["state"] = "adm_upload_report"
        await query.edit_message_text("📤 সবুজ ও লাল কালার করা <b>.xlsx</b> ফাইলটি এখানে পাঠান:", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "inbox_upload":
        context.user_data["state"] = "adm_inbox_uid"
        await query.edit_message_text("📩 ইউজারের <b>User ID (UID)</b> লিখুন:", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "work_mgr":
        cats = await db.fetchall("SELECT * FROM categories ORDER BY id ASC")
        kb = [[ikb_btn(f"{'🟢' if c['is_active'] else '🔴'} {c.get('emoji','📌')} {c['name']}", callback_data=f"adm:cat:{c['id']}", style="danger")] for c in cats]
        kb.append([ikb_btn("➕ Add Main Category", callback_data="adm:add_cat", style="primary")])
        kb.append([ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")])
        await query.edit_message_text("🛠 <b>Work Categories Manager</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "add_cat":
        context.user_data["state"] = "adm_add_cat"
        await query.edit_message_text("নতুন ক্যাটাগরি লিখুন (যেমন: <code>📘 | Facebook</code> বা শুধু <code>Facebook</code>):", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "cat":
        cat_id = int(parts[2])
        cat = await db.fetchone("SELECT * FROM categories WHERE id = ?", (cat_id,))
        subs = await db.fetchall("SELECT * FROM subcategories WHERE category_id = ? ORDER BY id ASC", (cat_id,))
        kb = [[
            ikb_btn(f"{'🟢 ON' if cat['is_active'] else '🔴 OFF'} Toggle", callback_data=f"adm:cat_tog:{cat_id}", style="primary"),
            ikb_btn("🗑 Delete", callback_data=f"adm:cat_del:{cat_id}", style="danger"),
        ]]
        for s in subs:
            kb.append([ikb_btn(f"{'🟢' if s['is_active'] else '🔴'} {s['name']} (৳{s['rate']:.0f})", callback_data=f"adm:sub:{s['id']}", style="danger")])
        kb.append([ikb_btn("➕ Add Subcategory", callback_data=f"adm:add_sub:{cat_id}", style="primary")])
        kb.append([ikb_btn("⬅️ Back", callback_data="adm:work_mgr", style="success")])
        await query.edit_message_text(f"📂 Category: <b>{cat['name']}</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "cat_tog":
        cat_id = int(parts[2])
        cat = await db.fetchone("SELECT is_active FROM categories WHERE id = ?", (cat_id,))
        await db.execute("UPDATE categories SET is_active = ? WHERE id = ?", (0 if cat["is_active"] else 1, cat_id))
        _CATS_CACHE = None
        await handle_admin_callback(query, context, f"adm:cat:{cat_id}")
        return

    if action == "cat_del":
        cat_id = int(parts[2])
        await db.execute("DELETE FROM subcategories WHERE category_id = ?", (cat_id,))
        await db.execute("DELETE FROM categories WHERE id = ?", (cat_id,))
        _CATS_CACHE = None
        _SUBS_CACHE.pop(cat_id, None)
        await handle_admin_callback(query, context, "adm:work_mgr")
        return

    if action == "add_sub":
        context.user_data["state"] = f"adm_add_sub:{parts[2]}"
        await query.edit_message_text("<code>Name | Rate</code> ফরম্যাটে লিখুন (যেমন: <code>USA Facebook | 10</code>):", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "sub":
        sub_id = int(parts[2])
        sub = await db.fetchone("SELECT * FROM subcategories WHERE id = ?", (sub_id,))
        steps = json.loads(sub["steps_json"])
        step_txt = "\n".join([f"{i+1}. <b>{st['name']}</b> ({st.get('validator','text')})" for i, st in enumerate(steps)]) or "কোনো স্টেপ নেই"
        kb = [
            [
                ikb_btn(f"{'🟢 ON' if sub['is_active'] else '🔴 OFF'} Toggle", callback_data=f"adm:sub_tog:{sub_id}", style="primary"),
                ikb_btn("💲 Edit Rate", callback_data=f"adm:sub_rate:{sub_id}", style="primary"),
            ],
            [
                ikb_btn("➕ Add Step", callback_data=f"adm:step_add:{sub_id}", style="success"),
                ikb_btn("🗑 Clear Last Step", callback_data=f"adm:step_pop:{sub_id}", style="success"),
            ],
            [ikb_btn("🗑 Delete Subcategory", callback_data=f"adm:sub_del:{sub_id}", style="danger")],
            [ikb_btn("⬅️ Back", callback_data=f"adm:cat:{sub['category_id']}", style="primary")],
        ]
        await query.edit_message_text(f"📌 <b>{sub['name']}</b> | Rate: <b>৳{sub['rate']:.0f}</b>\n\n<b>Steps:</b>\n{step_txt}", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "sub_tog":
        sub_id = int(parts[2])
        sub = await db.fetchone("SELECT category_id, is_active FROM subcategories WHERE id = ?", (sub_id,))
        await db.execute("UPDATE subcategories SET is_active = ? WHERE id = ?", (0 if sub["is_active"] else 1, sub_id))
        _SUBS_CACHE.pop(sub["category_id"], None)
        await handle_admin_callback(query, context, f"adm:sub:{sub_id}")
        return

    if action == "sub_rate":
        context.user_data["state"] = f"adm_sub_rate:{parts[2]}"
        await query.edit_message_text("নতুন রেট লিখুন:", reply_markup=back_to_admin_kb())
        return

    if action == "step_add":
        context.user_data["state"] = f"adm_step_add:{parts[2]}"
        await query.edit_message_text(
            "<code>Button Name | Prompt Message | validator</code>\n(validator: <code>hotmail_token</code>, <code>2fa</code>, বা <code>text</code>)",
            parse_mode=ParseMode.HTML,
            reply_markup=back_to_admin_kb(),
        )
        return

    if action == "step_pop":
        sub_id = int(parts[2])
        sub = await db.fetchone("SELECT category_id, steps_json FROM subcategories WHERE id = ?", (sub_id,))
        steps = json.loads(sub["steps_json"])
        if steps:
            steps.pop()
            await db.execute("UPDATE subcategories SET steps_json = ? WHERE id = ?", (json.dumps(steps, ensure_ascii=False), sub_id))
            _SUBS_CACHE.pop(sub["category_id"], None)
        await handle_admin_callback(query, context, f"adm:sub:{sub_id}")
        return

    if action == "sub_del":
        sub_id = int(parts[2])
        sub = await db.fetchone("SELECT category_id FROM subcategories WHERE id = ?", (sub_id,))
        await db.execute("DELETE FROM subcategories WHERE id = ?", (sub_id,))
        _SUBS_CACHE.pop(sub["category_id"], None)
        await handle_admin_callback(query, context, f"adm:cat:{sub['category_id']}")
        return

    if action == "bal_mgr":
        min_w = await get_setting("min_withdraw", "50")
        fee = await get_setting("withdraw_fee", "5")
        ref_b = await get_setting("referral_bonus", "3")
        kb = [
            [ikb_btn(f"📉 Min Withdraw: ৳{min_w}", callback_data="adm:set_min_wd", style="primary")],
            [ikb_btn(f"💸 Withdraw Fee: ৳{fee}", callback_data="adm:set_wd_fee", style="primary")],
            [ikb_btn(f"🎁 Referral Bonus: ৳{ref_b}", callback_data="adm:set_ref_bonus", style="primary")],
            [ikb_btn("💳 Edit User Balance (UID)", callback_data="adm:edit_user_bal", style="primary")],
            [ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")],
        ]
        await query.edit_message_text("💰 <b>Balance & Withdraw Settings</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action in ("set_min_wd", "set_wd_fee", "set_ref_bonus", "edit_user_bal"):
        context.user_data["state"] = f"adm_{action}"
        prompts = {
            "set_min_wd": "নতুন Minimum Withdraw লিখুন:",
            "set_wd_fee": "নতুন Withdraw Fee লিখুন:",
            "set_ref_bonus": "নতুন Referral Bonus লিখুন:",
            "edit_user_bal": "<code>UID | Amount</code> লিখুন (যেমন: <code>12345 | 50</code> বা <code>12345 | -20</code>):",
        }
        await query.edit_message_text(prompts[action], parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "pwd_mgr":
        cur_pwd = await get_current_password()
        prefix = await get_setting("password_prefix", "succes@")
        rtime = await get_setting("password_reset_time", "00:00")
        kb = [
            [ikb_btn(f"✏️ Prefix: {prefix}", callback_data="adm:set_pwd_prefix", style="primary")],
            [ikb_btn(f"⏰ Reset Time (BD): {rtime}", callback_data="adm:set_pwd_time", style="primary")],
            [ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")],
        ]
        await query.edit_message_text(f"🔑 <b>Auto Password Settings</b>\nবর্তমান: <code>{cur_pwd}</code>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "set_pwd_prefix":
        context.user_data["state"] = "adm_set_pwd_prefix"
        await query.edit_message_text("নতুন Prefix লিখুন (যেমন: <code>succes@</code>):", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "set_pwd_time":
        context.user_data["state"] = "adm_set_pwd_time"
        await query.edit_message_text("বাংলাদেশ সময় লিখুন (যেমন: <code>00:00</code>):", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "btn_mgr":
        lbls = await get_main_btn_labels()
        extra_btns = json.loads(await get_setting("extra_buttons_json", "[]"))
        kb = [[ikb_btn(f"✏️ Rename: {v}", callback_data=f"adm:mbtn_edit:{k}", style="primary")] for k, v in lbls.items()]
        for i, b in enumerate(extra_btns):
            kb.append([ikb_btn(f"🗑 Remove Extra: {b['label']}", callback_data=f"adm:btn_del:{i}", style="danger")])
        kb.append([ikb_btn("➕ Add Extra Main Button", callback_data="adm:btn_add", style="primary")])
        kb.append([ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")])
        await query.edit_message_text("⌨️ <b>Main Keyboard Buttons Manager</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "mbtn_edit":
        bkey = parts[2]
        context.user_data["state"] = f"adm_mbtn_edit:{bkey}"
        await query.edit_message_text("বাটনের নতুন নাম (ইমোজিসহ) লিখুন:", reply_markup=back_to_admin_kb())
        return

    if action == "btn_add":
        context.user_data["state"] = "adm_btn_add"
        await query.edit_message_text("<code>Button Name | Reply Text</code> ফরম্যাটে লিখুন:", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "btn_del":
        idx = int(parts[2])
        extra_btns = json.loads(await get_setting("extra_buttons_json", "[]"))
        if 0 <= idx < len(extra_btns):
            extra_btns.pop(idx)
            await set_setting("extra_buttons_json", json.dumps(extra_btns, ensure_ascii=False))
        await handle_admin_callback(query, context, "adm:btn_mgr")
        return

    if action == "guide_mgr":
        guides = await db.fetchall("SELECT * FROM guide_categories ORDER BY id ASC")
        kb = [[ikb_btn(f"🎬 {g['name']}", callback_data=f"adm:guide_edit:{g['id']}", style="primary")] for g in guides]
        kb.append([ikb_btn("➕ Add Guide Category", callback_data="adm:guide_add", style="primary")])
        kb.append([ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")])
        await query.edit_message_text("🎬 <b>আমি নতুন (Guide Manager)</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "guide_add":
        context.user_data["state"] = "adm_guide_add"
        await query.edit_message_text("নতুন গাইড ক্যাটাগরির নাম লিখুন:", reply_markup=back_to_admin_kb())
        return

    if action == "guide_edit":
        context.user_data["state"] = f"adm_guide_upload:{parts[2]}"
        await query.edit_message_text("টিউটোরিয়াল <b>Video</b> অথবা <b>Text</b> পাঠান:", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "fsub_mgr":
        cats = await db.fetchall("SELECT * FROM file_submit_categories ORDER BY id ASC")
        kb = [[ikb_btn(f"{'🟢' if c['is_active'] else '🔴'} {c['name']} (৳{c['rate']:.0f})", callback_data=f"adm:fsub_cat:{c['id']}", style="danger")] for c in cats]
        kb.append([ikb_btn("➕ Add File Submit Category", callback_data="adm:fsub_add", style="success")])
        kb.append([ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")])
        await query.edit_message_text("📁 <b>File Submit Categories</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "fsub_cat":
        cid = int(parts[2])
        c = await db.fetchone("SELECT * FROM file_submit_categories WHERE id = ?", (cid,))
        kb = [
            [
                ikb_btn(f"{'🟢 ON' if c['is_active'] else '🔴 OFF'} Toggle", callback_data=f"adm:fsub_tog:{cid}", style="primary"),
                ikb_btn("🗑 Delete", callback_data=f"adm:fsub_del:{cid}", style="danger"),
            ],
            [
                ikb_btn(f"🕒 Time: {c['cutoff_time']}", callback_data=f"adm:fsub_edit:{cid}:cutoff_time", style="primary"),
                ikb_btn(f"৳ Rate: {c['rate']:.0f}", callback_data=f"adm:fsub_edit:{cid}:rate", style="primary"),
            ],
            [
                ikb_btn(f"💲 USDT: {c['usdt_rate']:.4f}", callback_data=f"adm:fsub_edit:{cid}:usdt_rate", style="primary"),
                ikb_btn(f"📋 Report Col: {c['report_col']}", callback_data=f"adm:fsub_edit:{cid}:report_col", style="primary"),
            ],
            [ikb_btn(f"🖍 Format: {c['required_format']}", callback_data=f"adm:fsub_edit:{cid}:required_format", style="primary")],
            [ikb_btn("⬅️ Back", callback_data="adm:fsub_mgr", style="primary")],
        ]
        await query.edit_message_text(
            f"📁 <b>{c['name']}</b>\n"
            f"• Time: {c['cutoff_time']}\n"
            f"• Rate: {c['rate']:.0f} ৳ ({c['usdt_rate']:.4f} USDT)\n"
            f"• Format: <code>{c['required_format']}</code>\n"
            f"• Report Column: <b>{c['report_col']}</b>",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if action == "fsub_tog":
        cid = int(parts[2])
        c = await db.fetchone("SELECT is_active FROM file_submit_categories WHERE id = ?", (cid,))
        await db.execute("UPDATE file_submit_categories SET is_active = ? WHERE id = ?", (0 if c["is_active"] else 1, cid))
        await handle_admin_callback(query, context, f"adm:fsub_cat:{cid}")
        return

    if action == "fsub_del":
        cid = int(parts[2])
        await db.execute("DELETE FROM file_submit_categories WHERE id = ?", (cid,))
        await handle_admin_callback(query, context, "adm:fsub_mgr")
        return

    if action == "fsub_edit":
        cid, field = parts[2], parts[3]
        context.user_data["state"] = f"adm_fsub_edit:{cid}:{field}"
        hints = {
            "cutoff_time": "নতুন শেষ সময় লিখুন (যেমন: <code>10:00 PM</code>):",
            "rate": "নতুন রেট (টাকায়) লিখুন (যেমন: <code>5</code>):",
            "usdt_rate": "নতুন ডলার/USDT রেট লিখুন (যেমন: <code>0.0394</code>):",
            "required_format": "নতুন ফরম্যাট লিখুন (যেমন: <code>Uid pas cookies</code> বা <code>Uid pas 2fa hotmail</code>):",
            "report_col": "রিপোর্ট কলামের নাম লিখুন (<code>UID</code> অথবা <code>Username</code>):",
        }
        await query.edit_message_text(hints.get(field, "নতুন মান লিখুন:"), parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "fsub_add":
        context.user_data["state"] = "adm_fsub_add"
        await query.edit_message_text("নতুন File Submit ক্যাটাগরির নাম লিখুন (যেমন: <code>Fb cookies</code>):", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "fsub_recv":
        fsub_id = int(parts[2])
        fs = await db.fetchone("SELECT * FROM file_submissions WHERE id = ?", (fsub_id,))
        if not fs:
            await query.answer("পাওয়া যায়নি", show_alert=True)
            return
        await db.execute("UPDATE file_submissions SET status = 'received' WHERE id = ?", (fsub_id,))
        await db.execute("UPDATE file_submit_rows SET status = 'received' WHERE fsub_id = ? AND status = 'pending'", (fsub_id,))
        try:
            await query.edit_message_reply_markup(reply_markup=None)
        except Exception:
            pass
        try:
            await context.bot.send_message(
                chat_id=fs["user_id"],
                text=f"📥 আপনার <b>{fs['category_name']}</b> ফাইলটি ({fs['total_rows']}টি অ্যাকাউন্ট) অ্যাডমিন রিসিভ করেছেন!",
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass
        return

    if action == "fsub_report_pick":
        cats = await db.fetchall("SELECT * FROM file_submit_categories ORDER BY id ASC")
        if not cats:
            await query.edit_message_text("❌ কোনো File Submit ক্যাটাগরি নেই।", reply_markup=back_to_admin_kb())
            return
        kb = [[ikb_btn(f"📊 {c['name']}", callback_data=f"adm:fsub_rep_cat:{c['id']}", style="primary")] for c in cats]
        kb.append([ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")])
        await query.edit_message_text("📊 কোন ক্যাটাগরির <b>.xlsx</b> রিপোর্ট আপলোড করবেন?", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "fsub_rep_cat":
        cid = int(parts[2])
        context.user_data["state"] = f"adm_fsub_report:{cid}"
        await query.edit_message_text(
            "📤 <code>UID</code> বা <code>Username</code> কলামে <b>সবুজ (Ok)</b> ও <b>লাল (Back)</b> কালার করা <b>.xlsx</b> রিপোর্ট ফাইলটি এখন পাঠান:",
            parse_mode=ParseMode.HTML,
            reply_markup=back_to_admin_kb(),
        )
        return

    if action == "user_dir":
        page = int(parts[2]) if len(parts) > 2 else 0
        limit = 10
        offset = page * limit
        users = await db.fetchall("SELECT * FROM users ORDER BY joined_at DESC LIMIT ? OFFSET ?", (limit, offset))
        total_u = (await db.fetchone("SELECT COUNT(*) AS cnt FROM users"))["cnt"]
        lines = [f"👥 <b>User Directory</b> (মোট: {total_u} জন)\n"]
        kb = []
        for u in users:
            lines.append(f"• <code>{u['user_id']}</code> | {u['full_name']} | ৳{u['balance']:.1f} | ✅{u['completed_jobs']}")
            kb.append([ikb_btn(f"👁 {u['full_name']} ({u['user_id']})", callback_data=f"adm:view_inbox:{u['user_id']}", style="primary")])
        nav = []
        if page > 0:
            nav.append(ikb_btn("⬅️ Prev", callback_data=f"adm:user_dir:{page-1}", style="primary"))
        if offset + limit < total_u:
            nav.append(ikb_btn("Next ➡️", callback_data=f"adm:user_dir:{page+1}", style="primary"))
        if nav:
            kb.append(nav)
        kb.append([ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")])
        await query.edit_message_text("\n".join(lines), parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "file_history":
        batches = await db.fetchall("SELECT * FROM batches ORDER BY id DESC LIMIT 15")
        if not batches:
            await query.edit_message_text("📜 কোনো ফাইল হিস্ট্রি নেই।", reply_markup=back_to_admin_kb())
            return
        lines = ["📜 <b>Recent File History</b>\n"] + [
            f"• UID: <code>{b['user_id']}</code> | {b['subcategory_name']} | ✅{b['approved_count']} ❌{b['rejected_count']}"
            for b in batches
        ]
        await query.edit_message_text("\n".join(lines), parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "del_user":
        context.user_data["state"] = "adm_del_user"
        await query.edit_message_text("যার ডেটা ডিলিট করতে চান তার <b>UID</b> লিখুন:", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if action == "admins":
        adm_ids = await get_all_admin_ids()
        kb = []
        for aid in adm_ids:
            if aid == OWNER_ID:
                kb.append([ikb_btn(f"👑 Owner: {aid}", callback_data="adm:noop", style="primary")])
            else:
                kb.append([ikb_btn(f"🗑 Remove Admin: {aid}", callback_data=f"adm:adm_del:{aid}", style="danger")])
        kb.append([ikb_btn("➕ Add Admin (UID)", callback_data="adm:adm_add", style="success")])
        kb.append([ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")])
        await query.edit_message_text("👮 <b>Manage Admins</b>", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if action == "adm_add":
        context.user_data["state"] = "adm_add_admin"
        await query.edit_message_text("নতুন অ্যাডমিনের UID লিখুন:", reply_markup=back_to_admin_kb())
        return

    if action == "adm_del":
        target_aid = int(parts[2])
        if target_aid != OWNER_ID:
            await db.execute("DELETE FROM admins WHERE user_id = ?", (target_aid,))
            _ADMINS_CACHE = None
        await handle_admin_callback(query, context, "adm:admins")
        return

    if action == "broadcast":
        context.user_data["state"] = "adm_bc_msg"
        await query.edit_message_text("📢 সব ইউজার ও গ্রুপে পাঠানোর জন্য মেসেজ (বা ছবি/ভিডিও) দিন:", reply_markup=back_to_admin_kb())
        return

    if action == "edit_support":
        context.user_data["state"] = "adm_edit_support"
        cur_sup = await get_setting("support_link", "")
        await query.edit_message_text(f"বর্তমান লিংক: {cur_sup}\nনতুন লিংক দিন:", reply_markup=back_to_admin_kb())
        return


# ==============================================================================
# MESSAGE & STATE HANDLERS
# ==============================================================================
async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if update.effective_chat.type in ("group", "supergroup"):
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        await db.execute(
            "INSERT INTO bot_groups (chat_id, title, added_at) VALUES (?, ?, ?) ON CONFLICT (chat_id) DO UPDATE SET title = EXCLUDED.title"
            if db.use_pg
            else "INSERT OR REPLACE INTO bot_groups (chat_id, title, added_at) VALUES (?, ?, ?)",
            (update.effective_chat.id, update.effective_chat.title or "Group", now_str),
        )
        return

    uid = update.effective_user.id
    await ensure_user(update)

    if await get_setting("bot_active", "1") != "1" and not await is_admin(uid):
        await update.message.reply_text("🔧 বটের আপডেটের কাজ চলতেছে। অনুগ্রহ করে কিছুক্ষণ পর আবার চেষ্টা করুন।")
        return

    if not await check_force_join(update, context):
        return

    text = (update.message.text or "").strip()

    if context.user_data.get("active_task") and text:
        if await handle_task_text(update, context, text):
            return

    state = context.user_data.get("state", "")
    if state:
        await handle_stateful_message(update, context, state)
        return

    lbl = await get_main_btn_labels()
    if text in ("◀️ Work Menu ▸", lbl["work"], "💼 Work", "Work"):
        log_activity_bg(uid, "MENU", text)
        await handle_work_menu(update, context)
        return

    if text in ("❌ Cancel", "Cancel"):
        context.user_data.pop("nav", None)
        await update.message.reply_text("Back to main menu.", reply_markup=await main_keyboard(uid))
        return

    cats = await get_active_categories()
    for c in cats:
        btn_lbl = f"{c.get('emoji', '📌')} {c['name']} Work ▸"
        if text == btn_lbl:
            log_activity_bg(uid, "SELECT_CAT", c["name"])
            subs = await get_active_subcategories(c["id"])
            if not subs:
                await update.message.reply_text("❌ এই ক্যাটাগরিতে কোনো কাজ চালু নেই।")
                return
            context.user_data["nav"] = f"cat:{c['id']}"
            await update.message.reply_text(
                f"{c.get('emoji', '📸')} Select {c['name']} work:",
                reply_markup=await work_subcategories_keyboard(c["id"]),
            )
            return

    for c in cats:
        subs = await get_active_subcategories(c["id"])
        for s in subs:
            sub_lbl = f"{s['name']} • ৳{s['rate']:.0f}"
            if text == sub_lbl:
                steps = json.loads(s["steps_json"])
                if not steps:
                    await update.message.reply_text("❌ এই কাজের কোনো ধাপ সেট করা হয়নি।")
                    return
                gen_name, gen_username = await acquire_username_profile(uid)
                gen_password = await get_current_password()
                log_activity_bg(uid, "START_TASK", f"{s['name']} -> @{gen_username}")
                info_text = (
                    f"👤 Name:- <code>{gen_name}</code>\n"
                    f"🆔 Username:- <code>{gen_username}</code>\n"
                    f"🔐 Password:- <code>{gen_password}</code>\n\n"
                    f"📱 Tap the monospace text above to copy. Open {c['name']}, then tap <b>{steps[0]['name']}</b> below 😁\n\n"
                    f"💰 Rate: ৳{s['rate']:.0f}"
                )
                info_msg = await update.message.reply_text(info_text, parse_mode=ParseMode.HTML)
                task = {
                    "cat_id": c["id"],
                    "sub_id": s["id"],
                    "sub_name": s["name"],
                    "rate": float(s["rate"]),
                    "steps": steps,
                    "step_idx": 0,
                    "collected": {},
                    "gen_name": gen_name,
                    "gen_username": gen_username,
                    "gen_password": gen_password,
                    "msg_ids": [update.message.message_id, info_msg.message_id],
                    "awaiting_button": True,
                }
                context.user_data["active_task"] = task
                await send_next_step_prompt(update.effective_chat.id, context, task)
                return

    if text.startswith("📁 ") and text.endswith(" Submit ▸"):
        cat_name = text[2:-9].strip()
        cat = await db.fetchone("SELECT * FROM file_submit_categories WHERE name = ? AND is_active = 1", (cat_name,))
        if not cat:
            await update.message.reply_text("❌ এই ক্যাটাগরিটি বর্তমানে বন্ধ আছে।", reply_markup=await main_keyboard(uid))
            return
        context.user_data["state"] = f"user_fsub:{cat['id']}"
        await update.message.reply_text(
            f"📸 <b>Category:</b> {cat['name']}\n"
            f"🕒 <b>আজকের শেষ সময়:</b> {cat['cutoff_time']}\n"
            f"💲 <b>Per Account Rate:</b> {cat['rate']:.0f} ৳ ({cat['usdt_rate']:.4f} USDT)\n\n"
            f"🖍 <b>Required Format:</b>\n<code>{cat['required_format']}</code>\n\n"
            f"<i>অনুগ্রহ করে উপরের ফরম্যাট অনুযায়ী আপনার .xlsx ফাইলটি এখানে সাবমিট করুন।</i>",
            parse_mode=ParseMode.HTML,
            reply_markup=ReplyKeyboardMarkup([[kb_btn("❌ Cancel", style="danger")]], resize_keyboard=True),
        )
        return

    if text in (lbl["fsub"], "📁 File Submit", "File Submit"):
        log_activity_bg(uid, "MENU", "File Submit")
        await handle_file_submit_menu(update)
    elif text in (lbl["bal"], "💰 Balance", "Balance"):
        log_activity_bg(uid, "MENU", "Balance")
        await handle_balance(update)
    elif text in (lbl["wd"], "💸 Withdraw", "Withdraw"):
        log_activity_bg(uid, "MENU", "Withdraw")
        await handle_withdraw_start(update)
    elif text in (lbl["ref"], "🎁 My Referral", "My Referral"):
        log_activity_bg(uid, "MENU", "My Referral")
        await handle_referral(update, context)
    elif text in (lbl["sup"], "🆘 Support", "Support"):
        log_activity_bg(uid, "MENU", "Support")
        sup = await get_setting("support_link", "https://t.me/Talha_juba098")
        sup_msg = await get_setting("msg_support", "💬 <b>Customer Support</b>\n\nDear member,\nContact our support team for any problem or question. We will try to solve it quickly.")
        await update.message.reply_text(
            sup_msg,
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([[ikb_btn("📞 Contact Support", url=sup, style="primary")]]),
        )
    elif text in (lbl["new"], "🆕 I'm New", "🆕 আমি নতুন", "আমি নতুন"):
        log_activity_bg(uid, "MENU", "আমি নতুন")
        await handle_guide_menu(update)
    elif text == "⚙️ Admin Panel" and await is_admin(uid):
        await admin_cmd(update, context)
    else:
        extra_btns = json.loads(await get_setting("extra_buttons_json", "[]"))
        for b in extra_btns:
            if text == b["label"]:
                log_activity_bg(uid, "CUSTOM_BTN", text)
                await update.message.reply_text(b["reply"])
                return
        log_activity_bg(uid, "MESSAGE", text)


async def handle_stateful_message(update: Update, context: ContextTypes.DEFAULT_TYPE, state: str) -> None:
    global _CATS_CACHE, _ADMINS_CACHE
    uid = update.effective_user.id
    msg = update.message
    text = (msg.text or "").strip()

    if state == "wd_account":
        context.user_data["wd_data"]["account"] = text
        context.user_data["state"] = "wd_amount"
        min_w = float(await get_setting("min_withdraw", "50"))
        await msg.reply_text(f"💵 কত টাকা উইথড্র করতে চান লিখুন (সর্বনিম্ন ৳{min_w:.0f}):")
        return

    if state == "wd_amount":
        try:
            amount = float(text)
        except ValueError:
            await msg.reply_text("❌ সঠিক সংখ্যা লিখুন (যেমন: 50):")
            return
        min_w = float(await get_setting("min_withdraw", "50"))
        fee = float(await get_setting("withdraw_fee", "5"))
        invalidate_user_cache(uid)
        u = await ensure_user(update)
        if amount < min_w:
            await msg.reply_text(f"❌ সর্বনিম্ন উইথড্র ৳{min_w:.0f}।")
            return
        if amount > u["balance"]:
            await msg.reply_text(f"❌ আপনার ব্যালেন্সে পর্যাপ্ত টাকা নেই (ব্যালেন্স: ৳{u['balance']:.2f})।")
            return
        net = max(0.0, amount - fee)
        wd_data = context.user_data.pop("wd_data")
        context.user_data.pop("state", None)

        await db.execute(
            "UPDATE users SET balance = balance - ?, pending_withdraw = pending_withdraw + ? WHERE user_id = ?",
            (amount, amount, uid),
        )
        invalidate_user_cache(uid)
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        wd_id = await db.execute_insert_returning_id(
            "INSERT INTO withdrawals (user_id, method, account_number, requested_amount, fee, net_amount, status, created_at) VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)",
            (uid, wd_data["method"], wd_data["account"], amount, fee, net, now_str),
        )
        log_activity_bg(uid, "WITHDRAW", f"#WD-{wd_id} ৳{amount} via {wd_data['method']}")
        await msg.reply_text(
            f"✅ <b>উইথড্র রিকোয়েস্ট সম্পন্ন!</b> \n"
            f"• মেথড: {wd_data['method']}\n• অ্যাকাউন্ট: <code>{wd_data['account']}</code>\n"
            f"• রিকোয়েস্ট: ৳{amount:.2f} | ফি: ৳{fee:.2f}\n• <b>আপনি পাবেন: ৳{net:.2f}</b>",
            parse_mode=ParseMode.HTML,
        )
        adm_kb = InlineKeyboardMarkup([[
            ikb_btn("✅ Paid", callback_data=f"wd_adm:paid:{wd_id}", style="primary"),
            ikb_btn("❌ Reject & Refund", callback_data=f"wd_adm:reject:{wd_id}", style="primary"),
        ]])
        for adm_id in await get_all_admin_ids():
            try:
                await context.bot.send_message(
                    chat_id=adm_id,
                    text=(
                        f"💸 <b>নতুন উইথড্র রিকোয়েস্ট!</b> \n"
                        f"👤 {update.effective_user.full_name} (<code>{uid}</code>)\n"
                        f"💳 {wd_data['method']}: <code>{wd_data['account']}</code>\n"
                        f"📊 ৳{amount:.2f} - ৳{fee:.2f} = <b>৳{net:.2f}</b>"
                    ),
                    parse_mode=ParseMode.HTML,
                    reply_markup=adm_kb,
                )
            except Exception:
                pass
        return

    if state.startswith("user_fsub:"):
        if text in ("❌ Cancel", "Cancel"):
            context.user_data.pop("state", None)
            await msg.reply_text("Back to main menu.", reply_markup=await main_keyboard(uid))
            return
        cat_id = int(state.split(":", 1)[1])
        cat = await db.fetchone("SELECT * FROM file_submit_categories WHERE id = ?", (cat_id,))
        if not cat:
            context.user_data.pop("state", None)
            await msg.reply_text("❌ ক্যাটাগরি পাওয়া যায়নি।", reply_markup=await main_keyboard(uid))
            return
        fname = (msg.document.file_name or "").lower() if msg.document else ""
        if not msg.document or not fname.endswith(".xlsx"):
            await msg.reply_text("❌ অনুগ্রহ করে সঠিক ফরম্যাটের একটি <b>.xlsx</b> এক্সেল ফাইল পাঠান অথবা ❌ Cancel চাপুন।", parse_mode=ParseMode.HTML)
            return

        scan_msg = await msg.reply_text(f"🔄 Scanning your file for Category: <code>{cat['name']}</code>...", parse_mode=ParseMode.HTML)
        tg_file = await msg.document.get_file()
        bio = io.BytesIO()
        await tg_file.download_to_memory(out=bio)
        bio.seek(0)

        parsed_rows, err_reason = parse_and_validate_fsub_excel(bio, cat["required_format"])
        if err_reason:
            await scan_msg.edit_text(err_reason, parse_mode=ParseMode.HTML)
            return

        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        fsub_id = await db.execute_insert_returning_id(
            "INSERT INTO file_submissions (user_id, category_id, category_name, rate, total_rows, file_id, file_name, status, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, 'pending', ?)",
            (uid, cat["id"], cat["name"], float(cat["rate"]), len(parsed_rows), msg.document.file_id, msg.document.file_name or "file.xlsx", now_str),
        )
        for ident, raw_line in parsed_rows:
            await db.execute(
                "INSERT INTO file_submit_rows (fsub_id, user_id, category_id, category_name, rate, identifier, raw_line, status, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, 'pending', ?)",
                (fsub_id, uid, cat["id"], cat["name"], float(cat["rate"]), ident, raw_line, now_str),
            )

        log_activity_bg(uid, "FILE_SUBMIT", f"{cat['name']}: {len(parsed_rows)} rows")
        context.user_data.pop("state", None)

        await scan_msg.edit_text(
            f"✅ <b>File Submission Report</b> ✅\n\n"
            f"📁 <b>Category:</b> {cat['name']}\n"
            f"--------------------------------\n"
            f"📊 <b>Total Accounts Found:</b> {len(parsed_rows)}\n"
            f"--------------------------------\n\n"
            f"💰 রিপোর্ট আসলে শীঘ্রই ব্যালেন্স অটোমেটিক আপনার একাউন্টে যোগ হয়ে যাবে",
            parse_mode=ParseMode.HTML,
        )
        await msg.reply_text("📋 ফাইলটি অ্যাডমিন রিভিউয়ের জন্য পাঠানো হয়েছে।", reply_markup=await main_keyboard(uid))

        recv_kb = InlineKeyboardMarkup([[ikb_btn("📥 Receive File", callback_data=f"adm:fsub_recv:{fsub_id}", style="primary")]])
        for adm_id in await get_all_admin_ids():
            try:
                await context.bot.send_document(
                    chat_id=adm_id,
                    document=msg.document.file_id,
                    caption=(
                        f"📁 <b>New File Submit</b>\n"
                        f"👤 ইউজার: {update.effective_user.full_name} (<code>{uid}</code>)\n"
                        f"📂 ক্যাটাগরি: <b>{cat['name']}</b>\n"
                        f"📊 মোট অ্যাকাউন্ট: <b>{len(parsed_rows)}টি</b>"
                    ),
                    parse_mode=ParseMode.HTML,
                    reply_markup=recv_kb,
                )
            except Exception:
                pass
        return

    if not await is_admin(uid):
        context.user_data.pop("state", None)
        return

    if state == "adm_view_inbox_uid":
        if not text.isdigit():
            await msg.reply_text("❌ সঠিক UID দিন।")
            return
        tuid = int(text)
        context.user_data.pop("state", None)
        usr = await db.fetchone("SELECT * FROM users WHERE user_id = ?", (tuid,))
        logs = await db.fetchall("SELECT * FROM activity_logs WHERE user_id = ? ORDER BY id DESC LIMIT 20", (tuid,))
        header = f"👁 <b>User Inbox & Activity</b>\n👤 {usr['full_name'] if usr else tuid} (<code>{tuid}</code>)\n\n"
        log_lines = [f"• [{l['created_at'][11:16]}] <b>{l['action_type']}</b>: {l['detail']}" for l in reversed(logs)] or ["কোনো অ্যাক্টিভিটি নেই।"]
        kb = InlineKeyboardMarkup([
            [ikb_btn("👤 Open Telegram Chat", url=f"tg://user?id={tuid}", style="primary")],
            [ikb_btn("📩 Send Message / File to Inbox", callback_data=f"adm:inbox_direct:{tuid}", style="primary")],
            [ikb_btn("⬅️ Back to Admin Panel", callback_data="adm:home", style="primary")],
        ])
        await msg.reply_text(header + "\n".join(log_lines), parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if state == "adm_bc_msg":
        context.user_data["bc_msg_id"] = msg.message_id
        context.user_data["bc_chat_id"] = msg.chat_id
        context.user_data["state"] = "adm_bc_btn"
        await msg.reply_text(
            "🔗 মেসেজের নিচে কোনো লিংকযুক্ত বাটন দিতে চাইলে লিখুন:\n"
            "<code>Button Text | https://example.com</code>\n\n"
            "আর বাটন ছাড়া পাঠাতে চাইলে <code>skip</code> লিখুন:",
            parse_mode=ParseMode.HTML,
        )
        return

    if state == "adm_bc_btn":
        bc_msg_id = context.user_data.pop("bc_msg_id")
        bc_chat_id = context.user_data.pop("bc_chat_id")
        context.user_data.pop("state", None)
        kb = None
        if text.lower() != "skip" and "|" in text:
            b_lbl, b_url = [x.strip() for x in text.split("|", 1)]
            kb = InlineKeyboardMarkup([[ikb_btn(b_lbl, url=b_url, style="primary")]])

        users = await db.fetchall("SELECT user_id FROM users")
        groups = await db.fetchall("SELECT chat_id FROM bot_groups")
        targets = [u["user_id"] for u in users] + [g["chat_id"] for g in groups]
        sent = 0
        for tid in targets:
            try:
                await context.bot.copy_message(chat_id=tid, from_chat_id=bc_chat_id, message_id=bc_msg_id, reply_markup=kb)
                sent += 1
            except Exception:
                pass
        await msg.reply_text(f"📢 ব্রডকাস্ট সম্পন্ন! পাঠানো হয়েছে: <b>{sent}টি</b> চ্যাট/গ্রুপে।", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if state == "adm_fjoin_add" or state.startswith("adm_fjoin_edit:"):
        p = [x.strip() for x in text.split("|")]
        if len(p) < 2:
            await msg.reply_text("❌ ফরম্যাট: <code>@channel | https://t.me/link | Button Name</code>", parse_mode=ParseMode.HTML)
            return
        ch, lk = p[0], p[1]
        lbl = p[2] if len(p) >= 3 else f"Join {ch}"
        channels = await get_force_join_channels()
        entry = {"channel": ch, "link": lk, "label": lbl}
        if state == "adm_fjoin_add":
            channels.append(entry)
        else:
            idx = int(state.split(":")[1])
            if 0 <= idx < len(channels):
                channels[idx] = entry
        await set_setting("force_join_channels_json", json.dumps(channels, ensure_ascii=False))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ Force Join চ্যানেল আপডেট হয়েছে!", reply_markup=back_to_admin_kb())
        return

    if state == "adm_upload_report":
        if not msg.document or not (msg.document.file_name or "").lower().endswith(".xlsx"):
            await msg.reply_text("❌ ভ্যালিড `.xlsx` ফাইল পাঠান।")
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
            await msg.reply_text("❌ সঠিক UID দিন।")
            return
        context.user_data["inbox_target_uid"] = int(text)
        context.user_data["state"] = "adm_inbox_file"
        await msg.reply_text("এবার পাঠানোর জন্য ফাইল, ছবি বা মেসেজ দিন:")
        return

    if state == "adm_inbox_file":
        target_uid = context.user_data.pop("inbox_target_uid")
        context.user_data.pop("state", None)
        try:
            await msg.copy(chat_id=target_uid)
            await msg.reply_text(f"✅ UID <code>{target_uid}</code>-এ পাঠানো হয়েছে!", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        except Exception as e:
            await msg.reply_text(f"❌ পাঠানো যায়নি: {e}", reply_markup=back_to_admin_kb())
        return

    if state == "adm_add_cat":
        if "|" in text:
            emoji, cname = [x.strip() for x in text.split("|", 1)]
        else:
            emoji, cname = "📌", text
        await db.execute("INSERT INTO categories (name, emoji, is_active) VALUES (?, ?, 1)", (cname, emoji))
        _CATS_CACHE = None
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ ক্যাটাগরি <b>{emoji} {cname}</b> যুক্ত হয়েছে।", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_add_sub:"):
        cat_id = int(state.split(":")[1])
        sname, srate_str = [x.strip() for x in text.split("|", 1)]
        srate = float(srate_str)
        default_steps = [{"name": "Submit Details", "prompt": "কাজের তথ্য দিন:", "validator": "text"}]
        await db.execute(
            "INSERT INTO subcategories (category_id, name, rate, steps_json, is_active) VALUES (?, ?, ?, ?, 1)",
            (cat_id, sname, srate, json.dumps(default_steps, ensure_ascii=False)),
        )
        _SUBS_CACHE.pop(cat_id, None)
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ সাব-ক্যাটাগরি <b>{sname}</b> (৳{srate:.0f}) যুক্ত হয়েছে।", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_sub_rate:"):
        sub_id = int(state.split(":")[1])
        sub = await db.fetchone("SELECT category_id FROM subcategories WHERE id = ?", (sub_id,))
        await db.execute("UPDATE subcategories SET rate = ? WHERE id = ?", (float(text), sub_id))
        if sub:
            _SUBS_CACHE.pop(sub["category_id"], None)
        context.user_data.pop("state", None)
        await msg.reply_text("✅ রেট আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_step_add:"):
        sub_id = int(state.split(":")[1])
        parts = [x.strip() for x in text.split("|")]
        bname, bprompt = parts[0], parts[1]
        bval = parts[2] if len(parts) >= 3 else "text"
        sub = await db.fetchone("SELECT category_id, steps_json FROM subcategories WHERE id = ?", (sub_id,))
        steps = json.loads(sub["steps_json"])
        steps.append({"name": bname, "prompt": bprompt, "validator": bval})
        await db.execute("UPDATE subcategories SET steps_json = ? WHERE id = ?", (json.dumps(steps, ensure_ascii=False), sub_id))
        _SUBS_CACHE.pop(sub["category_id"], None)
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ নতুন স্টেপ <b>{bname}</b> যুক্ত হয়েছে।", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if state == "adm_set_min_wd":
        await set_setting("min_withdraw", str(float(text)))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ Minimum Withdraw আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_set_wd_fee":
        await set_setting("withdraw_fee", str(float(text)))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ Withdraw Fee আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_set_ref_bonus":
        await set_setting("referral_bonus", str(float(text)))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ Referral Bonus আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_edit_user_bal":
        p = [x.strip() for x in text.split("|")]
        tuid, amt = int(p[0]), float(p[1])
        await db.execute("UPDATE users SET balance = balance + ? WHERE user_id = ?", (amt, tuid))
        invalidate_user_cache(tuid)
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ UID {tuid}-এর ব্যালেন্সে {amt:+.2f} আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_set_pwd_prefix":
        await set_setting("password_prefix", text)
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ নতুন পাসওয়ার্ড: <code>{await get_current_password()}</code>", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if state == "adm_set_pwd_time":
        await set_setting("password_reset_time", text)
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ রিসেট টাইম আপডেট হয়েছে ({text})।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_btn_add":
        p = [x.strip() for x in text.split("|", 1)]
        extra_btns = json.loads(await get_setting("extra_buttons_json", "[]"))
        extra_btns.append({"label": p[0], "reply": p[1]})
        await set_setting("extra_buttons_json", json.dumps(extra_btns, ensure_ascii=False))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ নতুন বাটন যুক্ত হয়েছে!", reply_markup=back_to_admin_kb())
        return

    if state == "adm_guide_add":
        await db.execute("INSERT INTO guide_categories (name, guide_text, video_file_id, is_active) VALUES (?, ?, '', 1)", (text, "গাইডলাইন"))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ গাইড ক্যাটাগরি যুক্ত হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_guide_upload:"):
        gid = int(state.split(":")[1])
        if msg.video:
            await db.execute("UPDATE guide_categories SET video_file_id = ?, guide_text = ? WHERE id = ?", (msg.video.file_id, msg.caption or "", gid))
        else:
            await db.execute("UPDATE guide_categories SET guide_text = ? WHERE id = ?", (text, gid))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ গাইড আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_fsub_add":
        await db.execute(
            "INSERT INTO file_submit_categories (name, cutoff_time, rate, usdt_rate, required_format, report_col, is_active) VALUES (?, '10:00 PM', 5.0, 0.0394, 'Uid pas cookies', 'UID', 1)",
            (text,),
        )
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ File Submit ক্যাটাগরি <b>{text}</b> যুক্ত হয়েছে।", parse_mode=ParseMode.HTML, reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_fsub_edit:"):
        _, cid_str, field = state.split(":", 2)
        cid = int(cid_str)
        val: Any = float(text) if field in ("rate", "usdt_rate") else text
        await db.execute(f"UPDATE file_submit_categories SET {field} = ? WHERE id = ?", (val, cid))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ আপডেট হয়েছে!", reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_fsub_report:"):
        cid = int(state.split(":", 1)[1])
        if not msg.document or not (msg.document.file_name or "").lower().endswith(".xlsx"):
            await msg.reply_text("❌ ভ্যালিড `.xlsx` ফাইল পাঠান।")
            return
        tg_file = await msg.document.get_file()
        bio = io.BytesIO()
        await tg_file.download_to_memory(out=bio)
        bio.seek(0)
        await process_admin_fsub_report_excel(msg, context, bio, cid)
        context.user_data.pop("state", None)
        return

    if state == "adm_del_user":
        tuid = int(text)
        await db.execute("DELETE FROM activity_logs WHERE user_id = ?", (tuid,))
        await db.execute("DELETE FROM submissions WHERE user_id = ?", (tuid,))
        await db.execute("DELETE FROM withdrawals WHERE user_id = ?", (tuid,))
        await db.execute("DELETE FROM users WHERE user_id = ?", (tuid,))
        invalidate_user_cache(tuid)
        context.user_data.pop("state", None)
        await msg.reply_text(f"🗑 UID {tuid}-এর সব ডেটা মুছে ফেলা হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state == "adm_add_admin":
        aid = int(text)
        now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
        await db.execute(
            "INSERT INTO admins (user_id, added_by, created_at) VALUES (?, ?, ?) ON CONFLICT DO NOTHING"
            if db.use_pg
            else "INSERT OR IGNORE INTO admins (user_id, added_by, created_at) VALUES (?, ?, ?)",
            (aid, uid, now_str),
        )
        _ADMINS_CACHE = None
        context.user_data.pop("state", None)
        await msg.reply_text(f"✅ {aid} কে অ্যাডমিন করা হয়েছে।", reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_emsg:"):
        mkey = state.split(":", 1)[1]
        await set_setting(mkey, text)
        context.user_data.pop("state", None)
        await msg.reply_text("✅ মেসেজ আপডেট হয়েছে!", reply_markup=back_to_admin_kb())
        return

    if state.startswith("adm_mbtn_edit:"):
        bkey = state.split(":", 1)[1]
        lbls = await get_main_btn_labels()
        lbls[bkey] = text
        await set_setting("main_btns_json", json.dumps(lbls, ensure_ascii=False))
        context.user_data.pop("state", None)
        await msg.reply_text("✅ বাটনের নাম আপডেট হয়েছে!", reply_markup=back_to_admin_kb())
        return

    if state == "adm_edit_support":
        await set_setting("support_link", text)
        context.user_data.pop("state", None)
        await msg.reply_text("✅ সাপোর্ট লিংক আপডেট হয়েছে।", reply_markup=back_to_admin_kb())
        return


# ==============================================================================
# EXCEL GREEN/RED REPORT PROCESSOR
# ==============================================================================
async def process_admin_report_excel(msg, context: ContextTypes.DEFAULT_TYPE, bio: io.BytesIO) -> None:
    wb = openpyxl.load_workbook(bio)
    ws = wb.active

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    if "Submission_ID" not in headers:
        await msg.reply_text("❌ এই এক্সেল ফাইলে `Submission_ID` কলাম পাওয়া যায়নি!", reply_markup=back_to_admin_kb())
        return

    sub_id_col = headers.index("Submission_ID") + 1
    status_col = headers.index("Status") + 1 if "Status" in headers else len(headers)

    user_summary: Dict[int, Dict[str, Any]] = {}
    batch_stats: Dict[int, Dict[str, int]] = {}
    ref_bonus = float(await get_setting("referral_bonus", "3"))

    for row_idx in range(2, ws.max_row + 1):
        sub_val = ws.cell(row=row_idx, column=sub_id_col).value
        if not sub_val or not str(sub_val).isdigit():
            continue
        sub_id = int(sub_val)
        sub = await db.fetchone("SELECT * FROM submissions WHERE id = ?", (sub_id,))
        if not sub or sub["status"] in ("approved", "rejected"):
            continue

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
            await db.execute("UPDATE submissions SET status = 'approved' WHERE id = ?", (sub_id,))
            await db.execute(
                "UPDATE users SET balance = balance + ?, total_income = total_income + ?, completed_jobs = completed_jobs + 1 WHERE user_id = ?",
                (rate, rate, tuid),
            )
            invalidate_user_cache(tuid)
            user_summary[tuid]["approved"] += 1
            user_summary[tuid]["earned"] += rate
            if bid:
                batch_stats[bid]["approved"] += 1

            usr = await db.fetchone("SELECT referred_by, referral_paid FROM users WHERE user_id = ?", (tuid,))
            if usr and usr.get("referred_by") and not usr.get("referral_paid"):
                ref_id = usr["referred_by"]
                await db.execute("UPDATE users SET referral_paid = 1 WHERE user_id = ?", (tuid,))
                await db.execute(
                    "UPDATE users SET balance = balance + ?, total_income = total_income + ?, referral_earnings = referral_earnings + ? WHERE user_id = ?",
                    (ref_bonus, ref_bonus, ref_bonus, ref_id),
                )
                invalidate_user_cache(ref_id)
                try:
                    await context.bot.send_message(
                        chat_id=ref_id,
                        text=f"🎉 অভিনন্দন! আপনার রেফার করা ইউজারের প্রথম কাজ অ্যাপ্রুভ হওয়ায় আপনার ব্যালেন্সে <b>৳{ref_bonus:.0f}</b> বোনাস যোগ হয়েছে!",
                        parse_mode=ParseMode.HTML,
                    )
                except Exception:
                    pass

        elif color == "red":
            await db.execute("UPDATE submissions SET status = 'rejected' WHERE id = ?", (sub_id,))
            await db.execute("UPDATE users SET rejected_jobs = rejected_jobs + 1 WHERE user_id = ?", (tuid,))
            invalidate_user_cache(tuid)
            user_summary[tuid]["rejected"] += 1
            if bid:
                batch_stats[bid]["rejected"] += 1

    now_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d %H:%M:%S")
    for bid, st in batch_stats.items():
        await db.execute(
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
                    "📊 <b>কাজের রিপোর্ট প্রকাশিত হয়েছে!</b>\n\n"
                    f"✅ অ্যাপ্রুভ: <b>{v['approved']}টি</b>\n"
                    f"❌ রিজেক্ট: <b>{v['rejected']}টি</b>\n"
                    f"💰 ব্যালেন্সে যোগ হয়েছে: <b>৳{v['earned']:.2f}</b>"
                ),
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass

    await msg.reply_text(
        f"✅ <b>রিপোর্ট প্রসেস সম্পন্ন!</b>\n• মোট অ্যাপ্রুভ (সবুজ): {total_app}টি\n• মোট রিজেক্ট (লাল): {total_rej}টি",
        parse_mode=ParseMode.HTML,
        reply_markup=back_to_admin_kb(),
    )


async def process_admin_fsub_report_excel(msg, context: ContextTypes.DEFAULT_TYPE, bio: io.BytesIO, cat_id: int) -> None:
    cat = await db.fetchone("SELECT * FROM file_submit_categories WHERE id = ?", (cat_id,))
    if not cat:
        await msg.reply_text("❌ ক্যাটাগরি পাওয়া যায়নি।", reply_markup=back_to_admin_kb())
        return

    wb = openpyxl.load_workbook(bio)
    ws = wb.active
    col_header = cat["report_col"] or "UID"
    first_val = str(ws.cell(row=1, column=1).value or "").strip().lower()
    start_row = 2 if first_val in ("uid", "username", "id", "sl") else 1

    user_results: Dict[int, Dict[str, Any]] = {}
    ref_bonus = float(await get_setting("referral_bonus", "3"))
    today_str = datetime.datetime.now(BD_TZ).strftime("%Y-%m-%d")

    for r_idx in range(start_row, ws.max_row + 1):
        cell = ws.cell(row=r_idx, column=1)
        ident = str(cell.value or "").strip()
        if not ident:
            continue
        color = classify_excel_color(cell)
        if not color:
            continue

        row_rec = await db.fetchone(
            "SELECT * FROM file_submit_rows WHERE category_id = ? AND identifier = ? AND status IN ('pending', 'received') ORDER BY id DESC LIMIT 1",
            (cat_id, ident),
        )
        if not row_rec:
            continue

        tuid = row_rec["user_id"]
        if tuid not in user_results:
            user_results[tuid] = {"ok": 0, "back": 0, "earned": 0.0, "items": []}

        if color == "green":
            rate = float(row_rec["rate"])
            await db.execute("UPDATE file_submit_rows SET status = 'approved' WHERE id = ?", (row_rec["id"],))
            await db.execute(
                "UPDATE users SET balance = balance + ?, total_income = total_income + ?, completed_jobs = completed_jobs + 1 WHERE user_id = ?",
                (rate, rate, tuid),
            )
            invalidate_user_cache(tuid)
            user_results[tuid]["ok"] += 1
            user_results[tuid]["earned"] += rate
            user_results[tuid]["items"].append((ident, "green"))

            usr = await db.fetchone("SELECT referred_by, referral_paid FROM users WHERE user_id = ?", (tuid,))
            if usr and usr.get("referred_by") and not usr.get("referral_paid"):
                ref_id = usr["referred_by"]
                await db.execute("UPDATE users SET referral_paid = 1 WHERE user_id = ?", (tuid,))
                await db.execute(
                    "UPDATE users SET balance = balance + ?, total_income = total_income + ?, referral_earnings = referral_earnings + ? WHERE user_id = ?",
                    (ref_bonus, ref_bonus, ref_bonus, ref_id),
                )
                invalidate_user_cache(ref_id)
                try:
                    await context.bot.send_message(
                        chat_id=ref_id,
                        text=f"🎉 অভিনন্দন! আপনার রেফার করা ইউজারের প্রথম কাজ অ্যাপ্রুভ হওয়ায় আপনার ব্যালেন্সে <b>৳{ref_bonus:.0f}</b> বোনাস যোগ হয়েছে!",
                        parse_mode=ParseMode.HTML,
                    )
                except Exception:
                    pass
        elif color == "red":
            await db.execute("UPDATE file_submit_rows SET status = 'rejected' WHERE id = ?", (row_rec["id"],))
            await db.execute("UPDATE users SET rejected_jobs = rejected_jobs + 1 WHERE user_id = ?", (tuid,))
            invalidate_user_cache(tuid)
            user_results[tuid]["back"] += 1
            user_results[tuid]["items"].append((ident, "red"))

    green_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    total_ok, total_back = 0, 0
    for tuid, res in user_results.items():
        total_ok += res["ok"]
        total_back += res["back"]
        u_wb = openpyxl.Workbook()
        u_ws = u_wb.active
        u_ws.title = f"Report {today_str}"
        u_ws.append([col_header])
        for idx, (ident, clr) in enumerate(res["items"], start=2):
            c = u_ws.cell(row=idx, column=1, value=ident)
            c.fill = green_fill if clr == "green" else red_fill
        u_ws.column_dimensions["A"].width = 25
        u_bio = io.BytesIO()
        u_wb.save(u_bio)
        u_bio.seek(0)

        try:
            await context.bot.send_message(
                chat_id=tuid,
                text=(
                    f"🗓 <b>আপনার {cat['name']} ক্যাটাগরির ({today_str}) রিপোর্ট এসেছে!</b>\n\n"
                    f"✅ <b>Ok ID:</b> {res['ok']} টি\n"
                    f"❌ <b>Back ID:</b> {res['back']} টি\n"
                    f"💲 <b>Balance অ্যাড হয়েছে:</b> {res['earned']:.0f} ৳"
                ),
                parse_mode=ParseMode.HTML,
            )
            await context.bot.send_document(
                chat_id=tuid,
                document=u_bio,
                filename=f"Report_{cat['name']}_{today_str}.xlsx",
                caption="📁 আপনার বিস্তারিত কালারফুল রিপোর্ট ফাইলটি চেক করুন।",
            )
        except Exception:
            pass

    await msg.reply_text(
        f"✅ <b>File Submit রিপোর্ট প্রসেস সম্পন্ন!</b>\n• মোট Ok ID (সবুজ): {total_ok} টি\n• মোট Back ID (লাল): {total_back} টি",
        parse_mode=ParseMode.HTML,
        reply_markup=back_to_admin_kb(),
    )


# ==============================================================================
# WEBHOOK / POLLING + KEEP-ALIVE SERVER
# ==============================================================================
async def health_handler(request):
    return web.Response(text="Bot is running!", status=200)


async def run_bot() -> None:
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN environment variable is missing!")

    await db.init()
    await warm_caches()

    application = Application.builder().token(BOT_TOKEN).build()
    application.add_handler(CommandHandler("start", start_cmd))
    application.add_handler(CommandHandler("admin", admin_cmd))
    application.add_handler(CallbackQueryHandler(callback_router))
    application.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, message_handler))

    web_app = web.Application()
    web_app.router.add_get("/", health_handler)
    web_app.router.add_get("/health", health_handler)

    async with application:
        await application.initialize()

        if WEBHOOK_BASE:
            async def telegram_webhook(request):
                try:
                    data = await request.json()
                    upd = Update.de_json(data, application.bot)
                    await application.update_queue.put(upd)
                except Exception as e:
                    logger.error(f"Webhook parse error: {e}")
                return web.Response(text="OK", status=200)

            web_app.router.add_post("/webhook", telegram_webhook)

        runner = web.AppRunner(web_app)
        await runner.setup()
        site = web.TCPSite(runner, "0.0.0.0", PORT)
        await site.start()
        logger.info(f"Web server started on port {PORT}")

        await application.start()
        if WEBHOOK_BASE:
            wh_url = f"{WEBHOOK_BASE}/webhook"
            await application.bot.set_webhook(url=wh_url, drop_pending_updates=True)
            logger.info(f"Webhook active at {wh_url}")
        else:
            await application.bot.delete_webhook(drop_pending_updates=True)
            await application.updater.start_polling(drop_pending_updates=True)
            logger.info("Polling active...")

        try:
            while True:
                await asyncio.sleep(3600)
        finally:
            if not WEBHOOK_BASE and application.updater:
                await application.updater.stop()
            await application.stop()
            await runner.cleanup()
            await db.close()


if __name__ == "__main__":
    asyncio.run(run_bot())
