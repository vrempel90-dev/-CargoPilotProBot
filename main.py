from __future__ import annotations

import asyncio
from contextlib import asynccontextmanager
import html
import logging
import os
import re
from io import BytesIO
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Iterable, Optional

import aiosqlite
import qrcode
from openpyxl import Workbook, load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from aiohttp import web
from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.exceptions import TelegramBadRequest
from aiogram.filters import Command, CommandStart
from aiogram.filters.command import CommandObject
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    FSInputFile,
    BufferedInputFile,
)
from dotenv import load_dotenv

load_dotenv()

# =========================
# CONFIG
# =========================
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is empty. Add it to .env or hosting environment variables.")

COMPANY_NAME = os.getenv("COMPANY_NAME", "CargoAI CRM").strip()
DATA_DIR = Path(os.getenv("DATA_DIR", ".")).resolve()
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = os.getenv("DB_PATH", str(DATA_DIR / "cargo_ai_crm.sqlite3"))
EXPORT_DIR = DATA_DIR / "exports"
EXPORT_DIR.mkdir(parents=True, exist_ok=True)
RECEIPT_DIR = DATA_DIR / "receipts"
RECEIPT_DIR.mkdir(parents=True, exist_ok=True)
PORT = int(os.getenv("PORT", "8080"))
CURRENCY = os.getenv("CURRENCY", "$").strip() or "$"

# Бесплатные автостатусы: бот сам меняет статус по срокам маршрута.
# Не требует WhatsApp API, OpenAI, Gemini, Kaspi API или других платных сервисов.
AUTO_STATUS_ENABLED = os.getenv("AUTO_STATUS_ENABLED", "true").strip().lower() in {"1", "true", "yes", "on"}
AUTO_STATUS_INTERVAL_SECONDS = int(os.getenv("AUTO_STATUS_INTERVAL_SECONDS", "300"))
# days = реальная работа, minutes = демо-режим для быстрых показов клиенту. Оба режима бесплатные.
AUTO_STATUS_TIME_UNIT = os.getenv("AUTO_STATUS_TIME_UNIT", "days").strip().lower()
# Если true, новые заявки сразу получают бесплатный авто-маршрут без ручного включения.
AUTO_STATUS_ON_NEW_ORDERS = os.getenv("AUTO_STATUS_ON_NEW_ORDERS", "true").strip().lower() in {"1", "true", "yes", "on"}
REGION_MODE = os.getenv("REGION_MODE", "CIS").strip() or "CIS"

ADMIN_IDS = {
    int(x.strip())
    for x in os.getenv("ADMIN_IDS", "").split(",")
    if x.strip().lstrip("-").isdigit()
}
WAREHOUSE_IDS = {
    int(x.strip())
    for x in os.getenv("WAREHOUSE_IDS", "").split(",")
    if x.strip().lstrip("-").isdigit()
}
COURIER_IDS = {
    int(x.strip())
    for x in os.getenv("COURIER_IDS", "").split(",")
    if x.strip().lstrip("-").isdigit()
}

DEFAULT_RATES = {
    # Основные направления карго
    "китай": 3.5,
    "china": 3.5,
    "турция": 4.2,
    "turkey": 4.2,
    "дубай": 5.0,
    "uae": 5.0,
    "оаэ": 5.0,

    # СНГ / локальная логистика
    "казахстан": 2.0,
    "kazakhstan": 2.0,
    "россия": 2.5,
    "russia": 2.5,
    "узбекистан": 2.8,
    "uzbekistan": 2.8,
    "кыргызстан": 2.6,
    "киргизия": 2.6,
    "kyrgyzstan": 2.6,
    "беларусь": 3.0,
    "belarus": 3.0,
    "армения": 3.2,
    "armenia": 3.2,
    "азербайджан": 3.2,
    "azerbaijan": 3.2,
    "таджикистан": 3.0,
    "tajikistan": 3.0,
    "молдова": 3.1,
    "moldova": 3.1,
}
DEFAULT_COMMISSION_PERCENT = float(os.getenv("PARTNER_COMMISSION_PERCENT", "5"))

STATUSES = [
    "новая заявка",
    "принят на склад",
    "ожидает отправки",
    "отправлен",
    "в пути",
    "на таможне",
    "прибыл в страну",
    "прибыл в Казахстан",
    "прибыл в город",
    "готов к выдаче",
    "передан курьеру",
    "доставлен",
    "проблема",
    "отменён",
]

# Бесплатные шаблоны автостатусов. День считается от даты создания заявки.
# Это не интеграция с реальным складом/API, а автоматизация по типовым срокам маршрута.
AUTO_STATUS_ROUTES = {
    "china_cis": {
        "name": "Китай → СНГ",
        "stages": [
            (0, "новая заявка", "Заявка создана"),
            (1, "принят на склад", "Груз принят на склад отправителя"),
            (3, "отправлен", "Груз отправлен по маршруту"),
            (5, "в пути", "Груз находится в пути"),
            (8, "на таможне", "Груз проходит таможенный этап"),
            (12, "прибыл в страну", "Груз прибыл в страну назначения"),
            (14, "прибыл в город", "Груз прибыл в город получателя"),
            (15, "готов к выдаче", "Груз готов к выдаче"),
        ],
    },
    "turkey_cis": {
        "name": "Турция → СНГ",
        "stages": [
            (0, "новая заявка", "Заявка создана"),
            (1, "принят на склад", "Груз принят на склад отправителя"),
            (2, "отправлен", "Груз отправлен по маршруту"),
            (4, "в пути", "Груз находится в пути"),
            (7, "на таможне", "Груз проходит таможенный этап"),
            (10, "прибыл в страну", "Груз прибыл в страну назначения"),
            (12, "прибыл в город", "Груз прибыл в город получателя"),
            (13, "готов к выдаче", "Груз готов к выдаче"),
        ],
    },
    "uae_cis": {
        "name": "Дубай/ОАЭ → СНГ",
        "stages": [
            (0, "новая заявка", "Заявка создана"),
            (1, "принят на склад", "Груз принят на склад отправителя"),
            (2, "отправлен", "Груз отправлен по маршруту"),
            (4, "в пути", "Груз находится в пути"),
            (6, "на таможне", "Груз проходит таможенный этап"),
            (9, "прибыл в страну", "Груз прибыл в страну назначения"),
            (11, "прибыл в город", "Груз прибыл в город получателя"),
            (12, "готов к выдаче", "Груз готов к выдаче"),
        ],
    },
    "cis_local": {
        "name": "СНГ → СНГ",
        "stages": [
            (0, "новая заявка", "Заявка создана"),
            (1, "принят на склад", "Груз принят на склад"),
            (2, "отправлен", "Груз отправлен"),
            (3, "в пути", "Груз находится в пути"),
            (5, "прибыл в город", "Груз прибыл в город получателя"),
            (6, "готов к выдаче", "Груз готов к выдаче"),
        ],
    },
}

TERMINAL_STATUSES = {"доставлен", "отменён", "проблема", "передан курьеру"}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("cargo_ai_crm")
router = Router()


# =========================
# HELPERS
# =========================
def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def safe(text: object) -> str:
    return html.escape(str(text or ""))


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").lower().strip())


def short_username(message: Message) -> str:
    user = message.from_user
    if not user:
        return ""
    return user.username or ""


def full_name(message: Message) -> str:
    user = message.from_user
    if not user:
        return ""
    return " ".join(filter(None, [user.first_name, user.last_name])).strip()


def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


def is_warehouse(user_id: int) -> bool:
    return user_id in WAREHOUSE_IDS or is_admin(user_id)


def is_courier_static(user_id: int) -> bool:
    return user_id in COURIER_IDS or is_admin(user_id)


async def has_role(user_id: int, *roles: str) -> bool:
    if is_admin(user_id):
        return True
    user = await get_user(user_id)
    return bool(user and user["role"] in roles)


def partner_code(user_id: int) -> str:
    return f"P{abs(user_id)}"


def parse_tracking_code(text: str) -> Optional[str]:
    match = re.search(r"\bCG\d{6}\d{4,}\b", (text or "").upper())
    if match:
        return match.group(0)
    text = (text or "").strip().upper()
    if text.startswith("CG"):
        return text
    return None



def client_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📦 Оформить доставку"), KeyboardButton(text="🔎 Где мой груз?")],
            [KeyboardButton(text="🧮 Рассчитать доставку"), KeyboardButton(text="🛒 Выкуп товара")],
            [KeyboardButton(text="🏢 Оптовая доставка"), KeyboardButton(text="⚠️ Жалоба / проблема")],
            [KeyboardButton(text="🛠️ Техподдержка")],
            [KeyboardButton(text="👤 Мой кабинет"), KeyboardButton(text="📋 Мои заказы")],
            [KeyboardButton(text="🕓 История статусов"), KeyboardButton(text="📷 Фото груза")],
            [KeyboardButton(text="🔳 QR груза"), KeyboardButton(text="❓ FAQ")],
            [KeyboardButton(text="🤝 Партнёрка")],
        ],
        resize_keyboard=True,
        input_field_placeholder="Выберите действие или напишите номер груза CG...",
    )


def admin_keyboard() -> ReplyKeyboardMarkup:
    # Важные кнопки держим в первых строках: в Telegram Desktop нижние строки
    # reply-клавиатуры иногда скрываются, если окно маленькое.
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📥 Новые заявки"), KeyboardButton(text="📦 Все грузы")],
            [KeyboardButton(text="🤖 Автостатусы"), KeyboardButton(text="🔁 Изменить статус")],
            [KeyboardButton(text="🔎 Найти груз"), KeyboardButton(text="💸 Долги/оплаты")],
            [KeyboardButton(text="💰 Финансы"), KeyboardButton(text="🛠️ Техподдержка")],
            [KeyboardButton(text="💬 Жалобы"), KeyboardButton(text="🚚 Курьеры")],
            [KeyboardButton(text="📤 Excel экспорт"), KeyboardButton(text="📥 Обновить из Excel")],
            [KeyboardButton(text="⚙️ Тарифы"), KeyboardButton(text="📄 PDF квитанция")],
            [KeyboardButton(text="👥 Роли"), KeyboardButton(text="🤝 Партнёры")],
            [KeyboardButton(text="📊 Отчёт за день"), KeyboardButton(text="📢 Рассылка")],
            [KeyboardButton(text="🏭 Меню склада"), KeyboardButton(text="👤 Клиентское меню")],
        ],
        resize_keyboard=True,
        input_field_placeholder="Админ-панель: выберите действие",
    )


def warehouse_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="➕ Принять груз"), KeyboardButton(text="⚖️ Указать вес")],
            [KeyboardButton(text="📸 Добавить фото"), KeyboardButton(text="🔁 Изменить статус")],
            [KeyboardButton(text="📦 Грузы на складе"), KeyboardButton(text="🔎 Найти груз")],
            [KeyboardButton(text="🚚 Курьерское меню"), KeyboardButton(text="👤 Клиентское меню")],
        ],
        resize_keyboard=True,
        input_field_placeholder="Склад: введите CG... или выберите действие",
    )


def courier_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🚚 Мои доставки"), KeyboardButton(text="✅ Отметить доставлено")],
            [KeyboardButton(text="🔁 Изменить статус"), KeyboardButton(text="🔎 Найти груз")],
            [KeyboardButton(text="👤 Клиентское меню")],
        ],
        resize_keyboard=True,
        input_field_placeholder="Курьер: выберите действие",
    )



def client_menu_text(name: str = "") -> str:
    greeting = f", {safe(name)}" if name else ""
    return (
        f"<b>🚚 {safe(COMPANY_NAME)}</b>\n"
        f"Здравствуйте{greeting}!\n\n"
        "<b>Что можно сделать:</b>\n"
        "📦 оформить доставку или оптовую партию\n"
        "🔎 проверить статус груза по номеру\n"
        "🧮 рассчитать примерную стоимость\n"
        "🛒 оставить заявку на выкуп товара\n"
        "⚠️ отправить жалобу или проблему\n"
        "🛠️ создать обращение в техподдержку\n\n"
        "Напишите номер груза вида <code>CG...</code> или выберите действие ниже."
    )


def admin_menu_text() -> str:
    return (
        f"<b>👑 Админ-меню {safe(COMPANY_NAME)}</b>\n\n"
        "<b>Работаем только через Telegram:</b>\n"
        "📥 заявки и грузы — обработка клиентов\n"
        "🤖 автостатусы — бот сам ведёт груз по маршруту и уведомляет клиента\n"
        "🔎 поиск груза — быстро найти заказ по номеру\n"
        "🔁 статусы — ручная смена этапов, если нужно поправить маршрут\n"
        "💰 финансы — цена, оплаты, долги, маржа\n"
        "💬 жалобы — ответы клиентам\n"
        "🛠️ техподдержка — тикеты, ответы и закрытие обращений\n"
        "🚚 курьеры — выдача и доставка по городу\n"
        "📤 Excel — только для отчётов и массовых обновлений\n\n"
        "Для работы на компьютере откройте обычный <b>Telegram Desktop</b> и пользуйтесь этим же меню."
    )


def warehouse_menu_text() -> str:
    return (
        f"<b>🏭 Склад {safe(COMPANY_NAME)}</b>\n\n"
        "Здесь удобно работать с телефона:\n"
        "➕ принять груз\n"
        "⚖️ указать фактический вес\n"
        "📸 добавить фото коробки\n"
        "🔁 поменять статус и уведомить клиента\n\n"
        "Можно просто отправить номер груза <code>CG...</code>."
    )


def courier_menu_text() -> str:
    return (
        f"<b>🚚 Курьерское меню</b>\n\n"
        "Здесь курьер видит свои доставки, адреса и может отметить груз доставленным."
    )


def status_keyboard(order_id: int) -> InlineKeyboardMarkup:
    rows = []
    for status in STATUSES:
        rows.append([InlineKeyboardButton(text=status, callback_data=f"st:{order_id}:{status}")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def order_actions_keyboard(order_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🤖 Включить автостатусы", callback_data=f"auto_enable:{order_id}")],
            [InlineKeyboardButton(text="🔁 Изменить статус", callback_data=f"open_status:{order_id}")],
            [InlineKeyboardButton(text="✅ Доставлен", callback_data=f"st:{order_id}:доставлен")],
            [InlineKeyboardButton(text="⚠️ Проблема", callback_data=f"st:{order_id}:проблема")],
        ]
    )


def auto_status_keyboard(order_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="✅ Включить автостатусы", callback_data=f"auto_enable:{order_id}")],
            [InlineKeyboardButton(text="⏸ Выключить автостатусы", callback_data=f"auto_disable:{order_id}")],
            [InlineKeyboardButton(text="🔁 Изменить статус вручную", callback_data=f"open_status:{order_id}")],
        ]
    )


def auto_status_panel_keyboard(orders: list[aiosqlite.Row]) -> InlineKeyboardMarkup:
    """Главная панель автостатусов без команд.

    Админ не должен писать /autostatus или /autoroute. Он нажимает кнопку,
    а бот сам включает автостатусы и сам выбирает маршрут по направлению груза.
    """
    rows = [
        [InlineKeyboardButton(text="✅ Включить всем активным грузам", callback_data="auto_enable_all")],
        [InlineKeyboardButton(text="📋 Грузы без автостатусов", callback_data="auto_show_disabled")],
    ]
    for order in orders[:10]:
        enabled = bool(int(order["auto_status_enabled"] or 0))
        marker = "✅" if enabled else "🤖"
        code = order["tracking_code"]
        from_country = order["from_country"] or "?"
        to_city = order["to_city"] or "?"
        rows.append([
            InlineKeyboardButton(
                text=f"{marker} {code} · {from_country} → {to_city}",
                callback_data=f"open_auto:{order['id']}",
            )
        ])
    rows.append([InlineKeyboardButton(text="↩️ Назад в админ-меню", callback_data="admin_back")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def complaint_actions_keyboard(complaint_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="✅ Закрыть жалобу", callback_data=f"close_complaint:{complaint_id}")],
        ]
    )


def support_client_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🆕 Создать обращение", callback_data="support_new")],
            [InlineKeyboardButton(text="📋 Мои обращения", callback_data="support_my")],
        ]
    )


def support_topic_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📦 Вопрос по грузу", callback_data="support_topic:cargo")],
            [InlineKeyboardButton(text="💰 Оплата / долг", callback_data="support_topic:payment")],
            [InlineKeyboardButton(text="📸 Фото / повреждение", callback_data="support_topic:damage")],
            [InlineKeyboardButton(text="⚙️ Ошибка в боте", callback_data="support_topic:bot")],
            [InlineKeyboardButton(text="❓ Другое", callback_data="support_topic:other")],
        ]
    )


def support_ticket_actions_keyboard(ticket_id: int, include_client_reply: bool = False) -> InlineKeyboardMarkup:
    if include_client_reply:
        return InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="✍️ Ответить в обращение", callback_data=f"support_client_reply:{ticket_id}")],
            ]
        )
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="💬 Ответить", callback_data=f"support_reply:{ticket_id}")],
            [InlineKeyboardButton(text="✅ Закрыть", callback_data=f"support_close:{ticket_id}")],
        ]
    )


@asynccontextmanager
async def get_db():
    db = await aiosqlite.connect(DB_PATH)
    db.row_factory = aiosqlite.Row
    try:
        yield db
    finally:
        await db.close()




async def db_fetchone(db: aiosqlite.Connection, query: str, params: tuple = ()):
    cur = await db.execute(query, params)
    return await cur.fetchone()


async def db_fetchall(db: aiosqlite.Connection, query: str, params: tuple = ()) -> list[aiosqlite.Row]:
    cur = await db.execute(query, params)
    return await cur.fetchall()

async def init_db() -> None:
    async with get_db() as db:
        await db.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                telegram_id INTEGER PRIMARY KEY,
                username TEXT,
                full_name TEXT,
                role TEXT DEFAULT 'client',
                partner_code TEXT,
                ref_partner_id INTEGER,
                created_at TEXT,
                updated_at TEXT
            );

            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                tracking_code TEXT UNIQUE,
                order_type TEXT DEFAULT 'cargo',
                from_country TEXT,
                to_city TEXT,
                cargo_type TEXT,
                weight REAL,
                volume REAL,
                description TEXT,
                customer_name TEXT,
                phone TEXT,
                status TEXT DEFAULT 'новая заявка',
                price REAL DEFAULT 0,
                cost REAL DEFAULT 0,
                margin REAL DEFAULT 0,
                partner_id INTEGER,
                created_at TEXT,
                updated_at TEXT
            );

            CREATE TABLE IF NOT EXISTS status_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER,
                status TEXT,
                comment TEXT,
                created_by INTEGER,
                created_at TEXT
            );

            CREATE TABLE IF NOT EXISTS complaints (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                order_id INTEGER,
                tracking_code TEXT,
                text TEXT,
                urgency TEXT,
                status TEXT DEFAULT 'open',
                admin_reply TEXT,
                created_at TEXT,
                updated_at TEXT
            );


            CREATE TABLE IF NOT EXISTS support_tickets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                tracking_code TEXT,
                topic TEXT,
                status TEXT DEFAULT 'open',
                priority TEXT DEFAULT 'normal',
                last_message TEXT,
                created_at TEXT,
                updated_at TEXT,
                closed_at TEXT
            );

            CREATE TABLE IF NOT EXISTS support_messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id INTEGER NOT NULL,
                sender_id INTEGER,
                sender_role TEXT,
                text TEXT,
                created_at TEXT
            );

            CREATE TABLE IF NOT EXISTS partner_leads (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                partner_id INTEGER,
                client_id INTEGER,
                created_at TEXT,
                UNIQUE(partner_id, client_id)
            );

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            );

            CREATE TABLE IF NOT EXISTS tariffs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                country_key TEXT UNIQUE,
                country_name TEXT,
                rate REAL NOT NULL,
                created_at TEXT,
                updated_at TEXT
            );

            CREATE TABLE IF NOT EXISTS cargo_photos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER,
                file_id TEXT,
                comment TEXT,
                created_by INTEGER,
                created_at TEXT
            );

            CREATE TABLE IF NOT EXISTS payments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER,
                amount REAL,
                comment TEXT,
                created_by INTEGER,
                created_at TEXT
            );
            """
        )
        for sql in [
            "ALTER TABLE orders ADD COLUMN paid_amount REAL DEFAULT 0",
            "ALTER TABLE orders ADD COLUMN payment_status TEXT DEFAULT 'не оплачено'",
            "ALTER TABLE orders ADD COLUMN courier_id INTEGER",
            "ALTER TABLE orders ADD COLUMN delivery_address TEXT",
            "ALTER TABLE orders ADD COLUMN auto_status_enabled INTEGER DEFAULT 0",
            "ALTER TABLE orders ADD COLUMN auto_status_route TEXT",
            "ALTER TABLE orders ADD COLUMN auto_status_started_at TEXT",
            "ALTER TABLE orders ADD COLUMN auto_status_last_at TEXT",
        ]:
            try:
                await db.execute(sql)
            except Exception:
                pass
        for country, rate in DEFAULT_RATES.items():
            # В тарифы показываем русские названия, английские ключи оставляем только для распознавания.
            if re.search(r"[a-z]", country):
                continue
            await db.execute(
                "INSERT OR IGNORE INTO tariffs (country_key, country_name, rate, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
                (normalize(country), country.title(), rate, now_iso(), now_iso()),
            )
        await db.commit()


async def upsert_user(message: Message, ref_partner_id: Optional[int] = None) -> None:
    if not message.from_user:
        return
    user_id = message.from_user.id
    username = short_username(message)
    name = full_name(message)
    role = "admin" if is_admin(user_id) else "warehouse" if user_id in WAREHOUSE_IDS else "client"
    code = partner_code(user_id)
    async with get_db() as db:
        existing = await db_fetchone(db, 
            "SELECT telegram_id, ref_partner_id, role FROM users WHERE telegram_id=?",
            (user_id,),
        )
        if existing:
            current_ref = existing["ref_partner_id"]
            saved_role = existing["role"] or role
            if is_admin(user_id):
                saved_role = "admin"
            elif user_id in WAREHOUSE_IDS:
                saved_role = "warehouse"
            await db.execute(
                """
                UPDATE users
                SET username=?, full_name=?, role=?, partner_code=?,
                    ref_partner_id=COALESCE(ref_partner_id, ?), updated_at=?
                WHERE telegram_id=?
                """,
                (username, name, saved_role, code, ref_partner_id, now_iso(), user_id),
            )
            final_ref = current_ref or ref_partner_id
        else:
            await db.execute(
                """
                INSERT INTO users (telegram_id, username, full_name, role, partner_code, ref_partner_id, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (user_id, username, name, role, code, ref_partner_id, now_iso(), now_iso()),
            )
            final_ref = ref_partner_id
        if final_ref and final_ref != user_id:
            await db.execute(
                "INSERT OR IGNORE INTO partner_leads (partner_id, client_id, created_at) VALUES (?, ?, ?)",
                (final_ref, user_id, now_iso()),
            )
        await db.commit()


async def get_user(user_id: int) -> Optional[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchone(db, "SELECT * FROM users WHERE telegram_id=?", (user_id,))


async def get_order_by_code(code: str) -> Optional[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchone(db, "SELECT * FROM orders WHERE UPPER(tracking_code)=UPPER(?)", (code,))


async def get_order(order_id: int) -> Optional[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchone(db, "SELECT * FROM orders WHERE id=?", (order_id,))


def choose_auto_route(from_country: str, to_city: str = "") -> str:
    text = normalize(f"{from_country} {to_city}")
    if any(x in text for x in ["китай", "china", "guangzhou", "гуанчжоу", "иу", "yiwu", "1688", "taobao"]):
        return "china_cis"
    if any(x in text for x in ["турция", "turkey", "стамбул", "istanbul"]):
        return "turkey_cis"
    if any(x in text for x in ["дубай", "оаэ", "uae", "dubai", "emirates"]):
        return "uae_cis"
    return "cis_local"


def auto_route_label(route_key: str) -> str:
    route = AUTO_STATUS_ROUTES.get(route_key) or AUTO_STATUS_ROUTES["cis_local"]
    return route["name"]


async def enable_auto_status_for_order(order_id: int, actor_id: int = 0) -> Optional[aiosqlite.Row]:
    """Включает бесплатные автостатусы одной кнопкой.

    Маршрут выбирается автоматически по направлению груза, поэтому админу
    не нужно вводить команды /autoroute и выбирать шаблон вручную.
    """
    order = await get_order(order_id)
    if not order:
        return None
    route_key = choose_auto_route(order["from_country"] or "", order["to_city"] or "")
    async with get_db() as db:
        await db.execute(
            """
            UPDATE orders
            SET auto_status_enabled=1,
                auto_status_route=?,
                auto_status_started_at=?,
                auto_status_last_at=NULL,
                updated_at=?
            WHERE id=?
            """,
            (route_key, now_iso(), now_iso(), order_id),
        )
        await db.execute(
            "INSERT INTO status_history (order_id, status, comment, created_by, created_at) VALUES (?, ?, ?, ?, ?)",
            (order_id, order["status"] or "новая заявка", f"Автостатусы включены. Маршрут: {auto_route_label(route_key)}", actor_id, now_iso()),
        )
        await db.commit()
    return await get_order(order_id)


async def enable_auto_status_for_active_orders(actor_id: int = 0, limit: int = 500) -> int:
    """Включает автостатусы сразу для всех активных грузов.

    Это нужно для реальной карго-работы: админ не включает статусы каждому
    клиенту отдельно. Бот сам подбирает маршрут по направлению каждого груза.
    """
    async with get_db() as db:
        rows = await db_fetchall(
            db,
            """
            SELECT id FROM orders
            WHERE COALESCE(auto_status_enabled, 0)=0
              AND status NOT IN ('доставлен', 'отменён', 'проблема', 'передан курьеру')
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        )
    count = 0
    for row in rows:
        updated = await enable_auto_status_for_order(row["id"], actor_id)
        if updated:
            count += 1
    return count


async def list_orders_without_auto_status(limit: int = 10) -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(
            db,
            """
            SELECT * FROM orders
            WHERE COALESCE(auto_status_enabled, 0)=0
              AND status NOT IN ('доставлен', 'отменён', 'проблема', 'передан курьеру')
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        )


async def disable_auto_status_for_order(order_id: int, actor_id: int = 0) -> Optional[aiosqlite.Row]:
    order = await get_order(order_id)
    if not order:
        return None
    async with get_db() as db:
        await db.execute(
            "UPDATE orders SET auto_status_enabled=0, updated_at=? WHERE id=?",
            (now_iso(), order_id),
        )
        await db.execute(
            "INSERT INTO status_history (order_id, status, comment, created_by, created_at) VALUES (?, ?, ?, ?, ?)",
            (order_id, order["status"] or "новая заявка", "Автостатусы выключены", actor_id, now_iso()),
        )
        await db.commit()
    return await get_order(order_id)


def parse_iso_datetime(value: str) -> datetime:
    if not value:
        return datetime.now(timezone.utc)
    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return datetime.now(timezone.utc)


def get_auto_target_status(order: aiosqlite.Row) -> Optional[tuple[str, str, str]]:
    route_key = order["auto_status_route"] or choose_auto_route(order["from_country"] or "", order["to_city"] or "")
    route = AUTO_STATUS_ROUTES.get(route_key) or AUTO_STATUS_ROUTES["cis_local"]
    started_at = parse_iso_datetime(order["auto_status_started_at"] or order["created_at"] or now_iso())
    elapsed_seconds = (datetime.now(timezone.utc) - started_at).total_seconds()
    # В реальной работе этапы считаются днями. Для демонстрации можно поставить
    # AUTO_STATUS_TIME_UNIT=minutes, тогда те же значения станут минутами.
    elapsed_units = elapsed_seconds / 60 if AUTO_STATUS_TIME_UNIT == "minutes" else elapsed_seconds / 86400
    target = None
    target_index = -1
    for idx, (day, status, comment) in enumerate(route["stages"]):
        if elapsed_units >= day:
            target = (status, comment, route["name"])
            target_index = idx
    if not target:
        return None
    current_status = order["status"] or "новая заявка"
    if current_status in TERMINAL_STATUSES:
        return None
    stage_statuses = [s for _, s, _ in route["stages"]]
    current_index = stage_statuses.index(current_status) if current_status in stage_statuses else -1
    if target_index <= current_index or target[0] == current_status:
        return None
    return target


async def create_order(
    user_id: int,
    order_type: str,
    from_country: str,
    to_city: str,
    cargo_type: str,
    weight: float,
    volume: float,
    description: str,
    customer_name: str,
    phone: str,
    partner_id: Optional[int] = None,
) -> aiosqlite.Row:
    async with get_db() as db:
        cur = await db.execute(
            """
            INSERT INTO orders (
                user_id, order_type, from_country, to_city, cargo_type,
                weight, volume, description, customer_name, phone,
                status, partner_id, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'новая заявка', ?, ?, ?)
            """,
            (
                user_id,
                order_type,
                from_country,
                to_city,
                cargo_type,
                weight,
                volume,
                description,
                customer_name,
                phone,
                partner_id,
                now_iso(),
                now_iso(),
            ),
        )
        order_id = cur.lastrowid
        code = f"CG{datetime.now().strftime('%y%m%d')}{order_id:05d}"
        route_key = choose_auto_route(from_country, to_city)
        created_at = now_iso()
        auto_enabled = 1 if AUTO_STATUS_ON_NEW_ORDERS else 0
        auto_started_at = created_at if AUTO_STATUS_ON_NEW_ORDERS else None
        await db.execute(
            "UPDATE orders SET tracking_code=?, auto_status_route=?, auto_status_enabled=?, auto_status_started_at=? WHERE id=?",
            (code, route_key, auto_enabled, auto_started_at, order_id),
        )
        await db.execute(
            "INSERT INTO status_history (order_id, status, comment, created_by, created_at) VALUES (?, ?, ?, ?, ?)",
            (order_id, "новая заявка", "Заявка создана клиентом", user_id, now_iso()),
        )
        if AUTO_STATUS_ON_NEW_ORDERS:
            await db.execute(
                "INSERT INTO status_history (order_id, status, comment, created_by, created_at) VALUES (?, ?, ?, ?, ?)",
                (order_id, "новая заявка", f"Автостатусы включены автоматически. Маршрут: {auto_route_label(route_key)}", 0, now_iso()),
            )
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM orders WHERE id=?", (order_id,))


async def update_order_status(
    order_id: int,
    status: str,
    created_by: int,
    comment: str = "",
) -> Optional[aiosqlite.Row]:
    async with get_db() as db:
        await db.execute(
            "UPDATE orders SET status=?, updated_at=? WHERE id=?",
            (status, now_iso(), order_id),
        )
        await db.execute(
            "INSERT INTO status_history (order_id, status, comment, created_by, created_at) VALUES (?, ?, ?, ?, ?)",
            (order_id, status, comment, created_by, now_iso()),
        )
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM orders WHERE id=?", (order_id,))


async def set_order_price(code: str, price: float, cost: float) -> Optional[aiosqlite.Row]:
    margin = price - cost
    async with get_db() as db:
        existing = await db_fetchone(db, "SELECT paid_amount FROM orders WHERE UPPER(tracking_code)=UPPER(?)", (code,))
        if not existing:
            return None
        paid = float(existing["paid_amount"] or 0)
        payment_status = "оплачено" if price > 0 and paid >= price else "частично" if paid > 0 else "не оплачено"
        await db.execute(
            "UPDATE orders SET price=?, cost=?, margin=?, payment_status=?, updated_at=? WHERE UPPER(tracking_code)=UPPER(?)",
            (price, cost, margin, payment_status, now_iso(), code),
        )
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM orders WHERE UPPER(tracking_code)=UPPER(?)", (code,))


async def set_order_weight(code: str, weight: float, actor_id: int) -> Optional[aiosqlite.Row]:
    async with get_db() as db:
        row = await db_fetchone(db, "SELECT id FROM orders WHERE UPPER(tracking_code)=UPPER(?)", (code,))
        if not row:
            return None
        await db.execute("UPDATE orders SET weight=?, updated_at=? WHERE id=?", (weight, now_iso(), row["id"]))
        await db.execute(
            "INSERT INTO status_history (order_id, status, comment, created_by, created_at) VALUES (?, ?, ?, ?, ?)",
            (row["id"], "вес обновлён", f"Вес: {weight} кг", actor_id, now_iso()),
        )
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM orders WHERE id=?", (row["id"],))


async def get_rate_from_db(from_country: str) -> float:
    key = normalize(from_country)
    async with get_db() as db:
        row = await db_fetchone(db, "SELECT rate FROM tariffs WHERE country_key=?", (key,))
    if row:
        return float(row["rate"])
    return float(DEFAULT_RATES.get(key, 3.5))


async def estimate_delivery_price_db(from_country: str, weight: float, volume: float = 0) -> tuple[float, float]:
    rate = await get_rate_from_db(from_country)
    billable_weight = max(float(weight or 0), float(volume or 0) * 167 if volume else 0)
    if billable_weight <= 0:
        billable_weight = float(weight or 0)
    return round(rate, 2), round(rate * billable_weight, 2)


async def set_tariff(country_name: str, rate: float) -> None:
    key = normalize(country_name)
    async with get_db() as db:
        await db.execute(
            """
            INSERT INTO tariffs (country_key, country_name, rate, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(country_key) DO UPDATE SET country_name=excluded.country_name, rate=excluded.rate, updated_at=excluded.updated_at
            """,
            (key, country_name.strip(), rate, now_iso(), now_iso()),
        )
        await db.commit()


async def list_tariffs() -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(db, "SELECT * FROM tariffs ORDER BY country_name")


async def get_status_history(order_id: int) -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(
            db,
            "SELECT * FROM status_history WHERE order_id=? ORDER BY id ASC",
            (order_id,),
        )


async def add_cargo_photo(order_id: int, file_id: str, comment: str, created_by: int) -> None:
    async with get_db() as db:
        await db.execute(
            "INSERT INTO cargo_photos (order_id, file_id, comment, created_by, created_at) VALUES (?, ?, ?, ?, ?)",
            (order_id, file_id, comment, created_by, now_iso()),
        )
        await db.commit()


async def list_cargo_photos(order_id: int) -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(
            db,
            "SELECT * FROM cargo_photos WHERE order_id=? ORDER BY id DESC LIMIT 10",
            (order_id,),
        )


async def add_payment(code: str, amount: float, comment: str, created_by: int) -> Optional[aiosqlite.Row]:
    async with get_db() as db:
        order = await db_fetchone(db, "SELECT * FROM orders WHERE UPPER(tracking_code)=UPPER(?)", (code,))
        if not order:
            return None
        paid = float(order["paid_amount"] or 0) + amount
        price = float(order["price"] or 0)
        payment_status = "оплачено" if price > 0 and paid >= price else "частично" if paid > 0 else "не оплачено"
        await db.execute(
            "INSERT INTO payments (order_id, amount, comment, created_by, created_at) VALUES (?, ?, ?, ?, ?)",
            (order["id"], amount, comment, created_by, now_iso()),
        )
        await db.execute(
            "UPDATE orders SET paid_amount=?, payment_status=?, updated_at=? WHERE id=?",
            (paid, payment_status, now_iso(), order["id"]),
        )
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM orders WHERE id=?", (order["id"],))


async def list_debts(limit: int = 30) -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(
            db,
            """
            SELECT *, (COALESCE(price, 0) - COALESCE(paid_amount, 0)) AS debt
            FROM orders
            WHERE COALESCE(price, 0) > COALESCE(paid_amount, 0)
            ORDER BY updated_at DESC LIMIT ?
            """,
            (limit,),
        )


async def assign_courier(code: str, courier_id: int, address: str, actor_id: int) -> Optional[aiosqlite.Row]:
    async with get_db() as db:
        row = await db_fetchone(db, "SELECT id FROM orders WHERE UPPER(tracking_code)=UPPER(?)", (code,))
        if not row:
            return None
        await db.execute(
            "UPDATE orders SET courier_id=?, delivery_address=?, status=?, updated_at=? WHERE id=?",
            (courier_id, address, "передан курьеру", now_iso(), row["id"]),
        )
        await db.execute(
            "INSERT INTO status_history (order_id, status, comment, created_by, created_at) VALUES (?, ?, ?, ?, ?)",
            (row["id"], "передан курьеру", f"Курьер: {courier_id}. Адрес: {address}", actor_id, now_iso()),
        )
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM orders WHERE id=?", (row["id"],))


async def courier_orders(courier_id: int) -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(
            db,
            "SELECT * FROM orders WHERE courier_id=? AND status!='доставлен' ORDER BY updated_at DESC LIMIT 30",
            (courier_id,),
        )


async def list_orders(limit: int = 10, status: Optional[str] = None, user_id: Optional[int] = None) -> list[aiosqlite.Row]:
    query = "SELECT * FROM orders"
    params = []
    conditions = []
    if status:
        conditions.append("status=?")
        params.append(status)
    if user_id:
        conditions.append("user_id=?")
        params.append(user_id)
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    query += " ORDER BY id DESC LIMIT ?"
    params.append(limit)
    async with get_db() as db:
        return await db_fetchall(db, query, tuple(params))


async def orders_on_warehouse() -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(db, 
            """
            SELECT * FROM orders
            WHERE status IN ('принят на склад', 'ожидает отправки')
            ORDER BY id DESC LIMIT 20
            """
        )


async def create_complaint(user_id: int, code: str, text: str, urgency: str) -> aiosqlite.Row:
    order = await get_order_by_code(code) if code else None
    order_id = order["id"] if order else None
    async with get_db() as db:
        cur = await db.execute(
            """
            INSERT INTO complaints (user_id, order_id, tracking_code, text, urgency, status, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, 'open', ?, ?)
            """,
            (user_id, order_id, code, text, urgency, now_iso(), now_iso()),
        )
        cid = cur.lastrowid
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM complaints WHERE id=?", (cid,))


async def list_open_complaints(limit: int = 10) -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(db, 
            "SELECT * FROM complaints WHERE status='open' ORDER BY id DESC LIMIT ?",
            (limit,),
        )


async def close_complaint(complaint_id: int) -> None:
    async with get_db() as db:
        await db.execute(
            "UPDATE complaints SET status='closed', updated_at=? WHERE id=?",
            (now_iso(), complaint_id),
        )
        await db.commit()


async def reply_complaint(complaint_id: int, reply: str) -> Optional[aiosqlite.Row]:
    async with get_db() as db:
        await db.execute(
            "UPDATE complaints SET admin_reply=?, status='closed', updated_at=? WHERE id=?",
            (reply, now_iso(), complaint_id),
        )
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM complaints WHERE id=?", (complaint_id,))



def support_topic_label(topic: str) -> str:
    labels = {
        "cargo": "Вопрос по грузу",
        "payment": "Оплата / долг",
        "damage": "Фото / повреждение",
        "bot": "Ошибка в боте",
        "other": "Другое",
    }
    return labels.get(topic or "", topic or "Другое")


async def create_support_ticket(user_id: int, code: str, topic: str, text: str) -> aiosqlite.Row:
    async with get_db() as db:
        cur = await db.execute(
            """
            INSERT INTO support_tickets (user_id, tracking_code, topic, status, priority, last_message, created_at, updated_at)
            VALUES (?, ?, ?, 'open', 'normal', ?, ?, ?)
            """,
            (user_id, code, topic, text, now_iso(), now_iso()),
        )
        ticket_id = cur.lastrowid
        await db.execute(
            """
            INSERT INTO support_messages (ticket_id, sender_id, sender_role, text, created_at)
            VALUES (?, ?, 'client', ?, ?)
            """,
            (ticket_id, user_id, text, now_iso()),
        )
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM support_tickets WHERE id=?", (ticket_id,))


async def get_support_ticket(ticket_id: int) -> Optional[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchone(db, "SELECT * FROM support_tickets WHERE id=?", (ticket_id,))


async def list_open_support_tickets(limit: int = 15) -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(
            db,
            "SELECT * FROM support_tickets WHERE status!='closed' ORDER BY updated_at DESC LIMIT ?",
            (limit,),
        )


async def list_user_support_tickets(user_id: int, limit: int = 10) -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(
            db,
            "SELECT * FROM support_tickets WHERE user_id=? ORDER BY updated_at DESC LIMIT ?",
            (user_id, limit),
        )


async def list_support_messages(ticket_id: int, limit: int = 10) -> list[aiosqlite.Row]:
    async with get_db() as db:
        return await db_fetchall(
            db,
            "SELECT * FROM support_messages WHERE ticket_id=? ORDER BY id DESC LIMIT ?",
            (ticket_id, limit),
        )


async def add_support_message(ticket_id: int, sender_id: int, sender_role: str, text: str, new_status: Optional[str] = None) -> Optional[aiosqlite.Row]:
    ticket = await get_support_ticket(ticket_id)
    if not ticket:
        return None
    status = new_status or ("answered" if sender_role == "admin" else "open")
    async with get_db() as db:
        await db.execute(
            """
            INSERT INTO support_messages (ticket_id, sender_id, sender_role, text, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (ticket_id, sender_id, sender_role, text, now_iso()),
        )
        await db.execute(
            "UPDATE support_tickets SET last_message=?, status=?, updated_at=? WHERE id=?",
            (text, status, now_iso(), ticket_id),
        )
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM support_tickets WHERE id=?", (ticket_id,))


async def close_support_ticket(ticket_id: int) -> Optional[aiosqlite.Row]:
    async with get_db() as db:
        await db.execute(
            "UPDATE support_tickets SET status='closed', updated_at=?, closed_at=? WHERE id=?",
            (now_iso(), now_iso(), ticket_id),
        )
        await db.commit()
        return await db_fetchone(db, "SELECT * FROM support_tickets WHERE id=?", (ticket_id,))


def format_support_ticket(ticket: aiosqlite.Row) -> str:
    return (
        f"<b>🛠️ Обращение №{ticket['id']}</b>\n"
        f"Тема: {safe(support_topic_label(ticket['topic']))}\n"
        f"Груз: {safe(ticket['tracking_code'] or 'не указан')}\n"
        f"Клиент ID: <code>{ticket['user_id']}</code>\n"
        f"Статус: <b>{safe(ticket['status'])}</b>\n"
        f"Обновлено: {safe((ticket['updated_at'] or '')[:16].replace('T', ' '))}\n\n"
        f"Последнее сообщение:\n{safe(ticket['last_message'])}"
    )


async def finance_report() -> dict:
    async with get_db() as db:
        totals = await db_fetchone(db, 
            """
            SELECT COUNT(*) AS total_orders,
                   COALESCE(SUM(price), 0) AS revenue,
                   COALESCE(SUM(cost), 0) AS cost,
                   COALESCE(SUM(margin), 0) AS margin
            FROM orders
            """
        )
        today = datetime.now().strftime("%Y-%m-%d")
        today_row = await db_fetchone(db, 
            """
            SELECT COUNT(*) AS today_orders,
                   COALESCE(SUM(price), 0) AS today_revenue,
                   COALESCE(SUM(cost), 0) AS today_cost,
                   COALESCE(SUM(margin), 0) AS today_margin
            FROM orders
            WHERE created_at LIKE ?
            """,
            (f"%{today}%",),
        )
        by_status = await db_fetchall(db, 
            "SELECT status, COUNT(*) AS cnt FROM orders GROUP BY status ORDER BY cnt DESC"
        )
        return {"totals": totals, "today": today_row, "by_status": by_status}


async def partner_report(user_id: int) -> dict:
    async with get_db() as db:
        leads = await db_fetchone(db, 
            "SELECT COUNT(*) AS cnt FROM partner_leads WHERE partner_id=?",
            (user_id,),
        )
        orders = await db_fetchone(db, 
            """
            SELECT COUNT(*) AS cnt, COALESCE(SUM(price), 0) AS revenue
            FROM orders WHERE partner_id=?
            """,
            (user_id,),
        )
        partners = await db_fetchall(db, 
            "SELECT partner_id, COUNT(*) AS cnt FROM partner_leads GROUP BY partner_id ORDER BY cnt DESC LIMIT 10"
        )
        return {"leads": leads, "orders": orders, "partners": partners}


async def notify_admins(bot: Bot, text: str, reply_markup: Optional[InlineKeyboardMarkup] = None) -> None:
    for admin_id in ADMIN_IDS:
        try:
            await bot.send_message(admin_id, text, reply_markup=reply_markup)
        except Exception as e:
            logger.warning("Failed to notify admin %s: %s", admin_id, e)


async def notify_user(bot: Bot, user_id: int, text: str) -> None:
    try:
        await bot.send_message(user_id, text)
    except Exception as e:
        logger.warning("Failed to notify user %s: %s", user_id, e)


def detect_intent(text: str) -> Optional[str]:
    t = normalize(text)
    if any(w in t for w in ["где груз", "где посыл", "статус", "трек", "когда прид", "мой груз"]):
        return "tracking"
    if any(w in t for w in ["сколько", "цена", "стоимость", "тариф", "рассчитать", "за кг"]):
        return "calculator"
    if any(w in t for w in ["выкуп", "купить из китая", "1688", "таобао", "pinduoduo", "ссылка на товар"]):
        return "buyout"
    if any(w in t for w in ["опт", "партия", "короб", "контейнер", "большой груз"]):
        return "wholesale"
    if any(w in t for w in ["поддерж", "техпод", "оператор", "менеджер", "помощ", "саппорт"]):
        return "support"
    if any(w in t for w in ["жалоба", "проблем", "задерж", "повреж", "потерял"]):
        return "complaint"
    if any(w in t for w in ["документ", "тамож", "инвойс", "декларац"]):
        return "customs"
    return None


def estimate_delivery_price(from_country: str, weight: float, volume: float = 0) -> tuple[float, float]:
    key = normalize(from_country)
    rate = DEFAULT_RATES.get(key, 3.5)
    billable_weight = max(float(weight or 0), float(volume or 0) * 167 if volume else 0)
    if billable_weight <= 0:
        billable_weight = float(weight or 0)
    return round(rate, 2), round(rate * billable_weight, 2)


def format_order(order: aiosqlite.Row) -> str:
    return (
        f"<b>📦 Груз {safe(order['tracking_code'])}</b>\n"
        f"Тип заявки: {safe(order['order_type'])}\n"
        f"Маршрут: {safe(order['from_country'])} → {safe(order['to_city'])}\n"
        f"Товар: {safe(order['cargo_type'])}\n"
        f"Вес: {safe(order['weight'])} кг\n"
        f"Объём: {safe(order['volume'])} м³\n"
        f"Статус: <b>{safe(order['status'])}</b>\n"
        f"Автостатусы: {'вкл' if int(order['auto_status_enabled'] or 0) else 'выкл'} | {safe(auto_route_label(order['auto_status_route'] or choose_auto_route(order['from_country'] or '', order['to_city'] or '')))}\n"
        f"Цена клиенту: {safe(order['price'])} {safe(CURRENCY)}\n"
        f"Оплачено: {safe(order['paid_amount'])} {safe(CURRENCY)} | {safe(order['payment_status'])}\n"
        f"Курьер: {safe(order['courier_id']) if order['courier_id'] else 'не назначен'}\n"
        f"Адрес доставки: {safe(order['delivery_address']) if order['delivery_address'] else 'не указан'}\n"
        f"Описание: {safe(order['description'])}\n"
        f"Клиент: {safe(order['customer_name'])}, {safe(order['phone'])}"
    )


def format_short_order(order: aiosqlite.Row) -> str:
    return (
        f"{safe(order['tracking_code'])} | {safe(order['status'])} | "
        f"{safe(order['from_country'])} → {safe(order['to_city'])} | {safe(order['weight'])} кг"
    )


async def show_order_to_admin(bot: Bot, order: aiosqlite.Row) -> None:
    text = "<b>🆕 Новая заявка</b>\n\n" + format_order(order)
    await notify_admins(bot, text, order_actions_keyboard(order["id"]))


def format_history(rows: list[aiosqlite.Row]) -> str:
    if not rows:
        return "История пока пустая."
    lines = []
    for r in rows:
        dt = safe((r["created_at"] or "")[:16].replace("T", " "))
        comment = f" — {safe(r['comment'])}" if r["comment"] else ""
        lines.append(f"{dt}: <b>{safe(r['status'])}</b>{comment}")
    return "\n".join(lines)


def make_qr_bytes(text: str) -> bytes:
    img = qrcode.make(text)
    bio = BytesIO()
    img.save(bio, format="PNG")
    bio.seek(0)
    return bio.read()


def _register_pdf_font() -> str:
    for fp in [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]:
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont("CargoFont", fp))
                return "CargoFont"
            except Exception:
                pass
    return "Helvetica"


def generate_receipt_pdf(order: aiosqlite.Row) -> Path:
    font = _register_pdf_font()
    path = RECEIPT_DIR / f"receipt_{order['tracking_code']}.pdf"
    c = canvas.Canvas(str(path), pagesize=A4)
    width, height = A4
    y = height - 60
    c.setFont(font, 16)
    c.drawString(50, y, f"{COMPANY_NAME} — квитанция")
    y -= 40
    c.setFont(font, 11)
    lines = [
        f"Номер груза: {order['tracking_code']}",
        f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        f"Клиент: {order['customer_name']} / {order['phone']}",
        f"Маршрут: {order['from_country']} → {order['to_city']}",
        f"Товар: {order['cargo_type']}",
        f"Вес: {order['weight']} кг",
        f"Статус: {order['status']}",
        f"Цена: {order['price']} {CURRENCY}",
        f"Оплачено: {order['paid_amount']} {CURRENCY}",
        f"Статус оплаты: {order['payment_status']}",
        f"Адрес доставки: {order['delivery_address'] or 'не указан'}",
        "",
        "Квитанция сформирована автоматически в Telegram-системе CargoAI CRM.",
    ]
    for line in lines:
        c.drawString(50, y, str(line))
        y -= 22
    c.showPage()
    c.save()
    return path


async def export_orders_to_excel() -> Path:
    path = EXPORT_DIR / f"orders_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "orders"
    headers = [
        "tracking_code", "status", "order_type", "from_country", "to_city", "cargo_type",
        "weight", "volume", "price", "cost", "margin", "paid_amount", "payment_status",
        "customer_name", "phone", "courier_id", "delivery_address", "created_at", "updated_at", "description"
    ]
    ws.append(headers)
    async with get_db() as db:
        rows = await db_fetchall(db, "SELECT * FROM orders ORDER BY id DESC")
    for o in rows:
        ws.append([o[h] if h in o.keys() else "" for h in headers])
    widths = {"A": 18, "B": 18, "C": 14, "D": 14, "E": 14, "F": 18, "N": 18, "O": 18, "Q": 26, "T": 45}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    wb.save(path)
    return path


async def import_statuses_from_excel(path: Path, actor_id: int, bot: Bot) -> tuple[int, int]:
    wb = load_workbook(path)
    ws = wb.active
    headers = [str(c.value or "").strip().lower() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    def idx(name: str) -> Optional[int]:
        return headers.index(name) if name in headers else None
    i_code, i_status, i_comment = idx("tracking_code"), idx("status"), idx("comment")
    i_price, i_cost, i_paid = idx("price"), idx("cost"), idx("paid_amount")
    if i_code is None:
        return 0, 0
    ok = 0
    fail = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = str(row[i_code] or "").strip().upper()
        if not code:
            continue
        try:
            order = await get_order_by_code(code)
            if not order:
                fail += 1
                continue
            if i_status is not None and row[i_status]:
                status = str(row[i_status]).strip()
                comment = str(row[i_comment] or "") if i_comment is not None else "Импорт из Excel"
                updated = await update_order_status(order["id"], status, actor_id, comment)
                await notify_user(bot, updated["user_id"], f"📦 Статус груза <b>{safe(code)}</b> обновлён: <b>{safe(status)}</b>")
            if i_price is not None and row[i_price] is not None:
                price = float(row[i_price] or 0)
                cost = float(row[i_cost] or 0) if i_cost is not None else float(order["cost"] or 0)
                await set_order_price(code, price, cost)
            if i_paid is not None and row[i_paid] is not None:
                current = await get_order_by_code(code)
                target_paid = float(row[i_paid] or 0)
                delta = target_paid - float(current["paid_amount"] or 0)
                if abs(delta) > 0.0001:
                    await add_payment(code, delta, "Импорт оплаты из Excel", actor_id)
            ok += 1
        except Exception as e:
            logger.exception("Excel import failed for %s: %s", code, e)
            fail += 1
    return ok, fail


# =========================
# FSM STATES
# =========================
class CargoForm(StatesGroup):
    from_country = State()
    to_city = State()
    weight = State()
    volume = State()
    cargo_type = State()
    description = State()
    customer_name = State()
    phone = State()


class CalcForm(StatesGroup):
    from_country = State()
    weight = State()
    volume = State()


class TrackForm(StatesGroup):
    code = State()


class BuyoutForm(StatesGroup):
    link = State()
    details = State()
    customer_name = State()
    phone = State()


class WholesaleForm(StatesGroup):
    from_country = State()
    to_city = State()
    goods = State()
    boxes = State()
    weight = State()
    documents = State()
    customer_name = State()
    phone = State()


class ComplaintForm(StatesGroup):
    code = State()
    text = State()
    urgency = State()


class SupportForm(StatesGroup):
    code = State()
    topic = State()
    text = State()


class ClientSupportReplyForm(StatesGroup):
    ticket_id = State()
    text = State()


class AdminSupportReplyForm(StatesGroup):
    ticket_id = State()
    text = State()


class StatusForm(StatesGroup):
    code = State()
    status = State()
    comment = State()


class AutoStatusForm(StatesGroup):
    code = State()


class WeightForm(StatesGroup):
    code = State()
    weight = State()


class BroadcastForm(StatesGroup):
    text = State()


class HistoryForm(StatesGroup):
    code = State()


class PhotoAddForm(StatesGroup):
    code = State()
    photo = State()


class PhotoViewForm(StatesGroup):
    code = State()


class QrForm(StatesGroup):
    code = State()


class ReceiptForm(StatesGroup):
    code = State()


class PaymentForm(StatesGroup):
    code = State()
    amount = State()
    comment = State()


class ExcelImportForm(StatesGroup):
    file = State()


class TariffForm(StatesGroup):
    country = State()
    rate = State()


class CourierAssignForm(StatesGroup):
    code = State()
    courier_id = State()
    address = State()


class CourierDeliveredForm(StatesGroup):
    code = State()


# =========================
# COMMANDS / START
# =========================
@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext, command: CommandObject = None):
    await state.clear()
    ref_partner_id = None
    args = command.args if command else None
    if args and args.startswith("partner_"):
        raw = args.replace("partner_", "", 1)
        if raw.isdigit():
            ref_partner_id = int(raw)
    await upsert_user(message, ref_partner_id)

    if args and args.startswith("track_"):
        code = args.replace("track_", "", 1).strip().upper()
        order = await get_order_by_code(code)
        if order:
            rows = await get_status_history(order["id"])
            await message.answer(format_order(order) + "\n\n<b>🕓 История:</b>\n" + format_history(rows), reply_markup=client_keyboard())
            return

    name = message.from_user.first_name if message.from_user else ""
    await message.answer(client_menu_text(name), reply_markup=client_keyboard())


@router.message(Command("admin"))
async def cmd_admin(message: Message, state: FSMContext):
    await state.clear()
    await upsert_user(message)
    if not message.from_user or not is_admin(message.from_user.id):
        await message.answer("⛔ У вас нет доступа к админ‑меню.")
        return
    await message.answer(admin_menu_text(), reply_markup=admin_keyboard())


@router.message(Command("warehouse"))
async def cmd_warehouse(message: Message, state: FSMContext):
    await state.clear()
    await upsert_user(message)
    if not message.from_user or not is_warehouse(message.from_user.id):
        await message.answer("⛔ У вас нет доступа к складскому меню.")
        return
    await message.answer(warehouse_menu_text(), reply_markup=warehouse_keyboard())


@router.message(Command("help"))
async def cmd_help(message: Message):
    if message.from_user and is_admin(message.from_user.id):
        await message.answer(
            "<b>Помощь CargoPilot Pro</b>\n\n"
            "/start — клиентское меню\n"
            "/admin — админ-меню\n"
            "/warehouse — меню склада\n"
            "/courier — меню курьера\n"
            "/routes — маршруты автостатусов\n"
            "/support — открытые обращения\n"
            "/help_admin — все команды админа"
        )
    else:
        await message.answer(
            "<b>Помощь</b>\n\n"
            "Нажмите /start, чтобы открыть меню.\n"
            "Вы можете оформить доставку, рассчитать стоимость, проверить статус груза, создать обращение в техподдержку или посмотреть свои заказы."
        )


@router.message(Command("help_admin"))
async def cmd_help_admin(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    await message.answer(
        "<b>Команды админа</b>\n\n"
        "/admin — открыть админ‑меню\n"
        "/setstatus CG26050300001 в пути комментарий — сменить статус\n"
        "/setprice CG26050300001 100 70 — цена клиенту, себестоимость\n"
        "/replycomplaint 1 текст ответа — ответить на жалобу\n"
        "/support — список обращений техподдержки\n"
        "/replyticket 1 текст ответа — ответить в обращение\n"
        "/closeticket 1 — закрыть обращение\n"
        "/makepartner 123456789 — сделать пользователя партнёром\n"
        "/role 123456789 warehouse — выдать роль warehouse/client/partner/courier/admin\n"
        "/history CG... — история статусов\n"
        "/photos CG... — фото груза\n"
        "/qr CG... — QR-код груза\n"
        "/pay CG... 10000 комментарий — внести оплату\n"
        "/debts — список долгов\n"
        "/export_orders — выгрузка Excel\n"
        "/settariff Китай 3.5 — изменить тариф\n"
        "/tariffs — список тарифов\n"
        "/receipt CG... — PDF-квитанция\n"
        "/assigncourier CG... 123456789 адрес — назначить курьера\n"
        "🤖 Автостатусы — включить по кнопке без команд\n"
        "📦 Все грузы → карточка груза → 🤖 Включить автостатусы\n"
    )


@router.message(Command("setstatus"))
async def cmd_setstatus(message: Message, bot: Bot):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split(maxsplit=2)
    if len(parts) < 3:
        await message.answer("Формат: /setstatus CG26050300001 в пути")
        return
    code = parts[1].upper()
    rest = parts[2]
    status = None
    for s in sorted(STATUSES, key=len, reverse=True):
        if rest.lower().startswith(s.lower()):
            status = s
            comment = rest[len(s):].strip()
            break
    if not status:
        status = rest.strip()
        comment = ""
    order = await get_order_by_code(code)
    if not order:
        await message.answer("Груз не найден.")
        return
    updated = await update_order_status(order["id"], status, message.from_user.id, comment)
    await message.answer("✅ Статус обновлён.\n\n" + format_order(updated))
    await notify_user(
        bot,
        updated["user_id"],
        f"📦 Статус груза <b>{safe(updated['tracking_code'])}</b> обновлён: <b>{safe(status)}</b>\n{safe(comment)}",
    )


@router.message(Command("routes"))
async def cmd_routes(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    lines = ["<b>🤖 Шаблоны автостатусов</b>"]
    for key, route in AUTO_STATUS_ROUTES.items():
        steps = ", ".join(f"день {d}: {s}" for d, s, _ in route["stages"][:4])
        lines.append(f"<code>{safe(key)}</code> — {safe(route['name'])}\n{safe(steps)} ...")
    lines.append("\nТеперь всё работает без команд: откройте <b>/admin</b>, нажмите <b>🤖 Автостатусы</b> и выберите действие кнопкой. Бот сам определит маршрут по направлению груза.")
    await message.answer("\n\n".join(lines), reply_markup=admin_keyboard())


@router.message(Command("autostatus"))
async def cmd_autostatus(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split(maxsplit=2)
    if len(parts) < 3 or parts[2].lower() not in {"on", "off", "вкл", "выкл"}:
        await message.answer("Формат: /autostatus CG26050300001 on или /autostatus CG26050300001 off")
        return
    code = parts[1].upper()
    order = await get_order_by_code(code)
    if not order:
        await message.answer("Груз не найден.")
        return
    if parts[2].lower() in {"on", "вкл"}:
        updated = await enable_auto_status_for_order(order["id"], message.from_user.id)
        route_name = auto_route_label(updated["auto_status_route"] or choose_auto_route(updated["from_country"] or "", updated["to_city"] or ""))
        await message.answer(
            f"✅ Автостатусы включены для <b>{safe(code)}</b>.\n"
            f"Маршрут выбран автоматически: <b>{safe(route_name)}</b>.\n\n"
            "Теперь бот будет сам менять статусы по расписанию маршрута.",
            reply_markup=admin_keyboard(),
        )
    else:
        await disable_auto_status_for_order(order["id"], message.from_user.id)
        await message.answer(f"⏸ Автостатусы для <b>{safe(code)}</b> выключены.", reply_markup=admin_keyboard())


@router.message(Command("autostatus_all"))
async def cmd_autostatus_all(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split(maxsplit=1)
    if len(parts) < 2 or parts[1].lower() not in {"on", "вкл"}:
        await message.answer("Формат: /autostatus_all on")
        return
    count = await enable_auto_status_for_active_orders(message.from_user.id)
    await message.answer(
        f"✅ Автостатусы включены для активных грузов: {count}.\n\n"
        "Новые заявки также получают автостатусы автоматически, если AUTO_STATUS_ON_NEW_ORDERS=true.",
        reply_markup=admin_keyboard(),
    )


@router.message(Command("autoroute"))
async def cmd_autoroute(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split(maxsplit=2)
    if len(parts) < 3 or parts[2] not in AUTO_STATUS_ROUTES:
        await message.answer("Формат: /autoroute CG26050300001 china_cis\n\nСписок: /routes")
        return
    code = parts[1].upper()
    route_key = parts[2]
    async with get_db() as db:
        row = await db_fetchone(db, "SELECT id FROM orders WHERE UPPER(tracking_code)=UPPER(?)", (code,))
        if not row:
            await message.answer("Груз не найден.")
            return
        await db.execute(
            "UPDATE orders SET auto_status_route=?, auto_status_enabled=1, auto_status_started_at=?, updated_at=? WHERE id=?",
            (route_key, now_iso(), now_iso(), row["id"]),
        )
        await db.commit()
    await message.answer(f"✅ Маршрут автостатусов для {safe(code)}: {safe(auto_route_label(route_key))}")


@router.message(F.text == "🤖 Автостатусы")
async def admin_auto_status_menu(message: Message, state: FSMContext):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    await state.clear()
    orders = await list_orders(limit=10)
    text = (
        "<b>🤖 Автостатусы без команд</b>\n\n"
        "Админу ничего не нужно писать вручную. Нажмите кнопку ниже — бот сам включит автостатусы, "
        "сам определит маршрут по направлению груза и будет менять этапы по расписанию.\n\n"
        "<b>Как работает:</b>\n"
        "1) клиент оставляет заявку;\n"
        "2) бот выбирает маршрут: Китай/Турция/ОАЭ/СНГ;\n"
        "3) статус меняется автоматически;\n"
        "4) клиент получает уведомления.\n\n"
        "Выберите действие:"
    )
    await message.answer(text, reply_markup=auto_status_panel_keyboard(orders))


@router.message(AutoStatusForm.code)
async def auto_status_code(message: Message, state: FSMContext):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    raw = (message.text or "").strip()
    if normalize(raw) in {"все", "all", "всё"}:
        count = await enable_auto_status_for_active_orders(message.from_user.id)
        await state.clear()
        await message.answer(
            f"✅ Автостатусы включены для активных грузов: {count}.\n\n"
            "Бот сам выбрал маршруты по направлениям и будет уведомлять клиентов при смене статусов.",
            reply_markup=admin_keyboard(),
        )
        return
    code = parse_tracking_code(raw) or raw.upper()
    order = await get_order_by_code(code)
    if not order:
        await message.answer("Груз не найден. Нажмите кнопку «🤖 Автостатусы» ещё раз или введите другой номер CG...")
        return
    updated = await enable_auto_status_for_order(order["id"], message.from_user.id)
    await state.clear()
    route_name = auto_route_label(updated["auto_status_route"] or choose_auto_route(updated["from_country"] or "", updated["to_city"] or ""))
    await message.answer(
        f"✅ Автостатусы включены для <b>{safe(updated['tracking_code'])}</b>.\n"
        f"Маршрут выбран автоматически: <b>{safe(route_name)}</b>.\n\n"
        "Бот сам будет менять статусы по расписанию маршрута и уведомлять клиента.\n"
        "Если груз задержится, статус можно изменить вручную в любой момент.",
        reply_markup=admin_keyboard(),
    )


@router.message(Command("setprice"))
async def cmd_setprice(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split()
    if len(parts) != 4:
        await message.answer("Формат: /setprice CG26050300001 100 70")
        return
    code = parts[1].upper()
    try:
        price = float(parts[2].replace(",", "."))
        cost = float(parts[3].replace(",", "."))
    except ValueError:
        await message.answer("Цена и себестоимость должны быть числами.")
        return
    order = await set_order_price(code, price, cost)
    if not order:
        await message.answer("Груз не найден.")
        return
    await message.answer(
        f"✅ Финансы обновлены.\n"
        f"Груз: {safe(code)}\n"
        f"Цена: {price} {safe(CURRENCY)}\n"
        f"Себестоимость: {cost} {safe(CURRENCY)}\n"
        f"Маржа: {price - cost} {safe(CURRENCY)}"
    )


@router.message(Command("replycomplaint"))
async def cmd_reply_complaint(message: Message, bot: Bot):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split(maxsplit=2)
    if len(parts) < 3 or not parts[1].isdigit():
        await message.answer("Формат: /replycomplaint 1 текст ответа")
        return
    complaint = await reply_complaint(int(parts[1]), parts[2])
    if not complaint:
        await message.answer("Жалоба не найдена.")
        return
    await message.answer("✅ Ответ отправлен клиенту, жалоба закрыта.")
    await notify_user(
        bot,
        complaint["user_id"],
        f"💬 Ответ по вашей жалобе №{complaint['id']}:\n\n{safe(parts[2])}",
    )



@router.message(Command("support"))
async def cmd_support(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    tickets = await list_open_support_tickets(20)
    if not tickets:
        await message.answer("Открытых обращений в техподдержку нет.", reply_markup=admin_keyboard())
        return
    await message.answer("<b>🛠️ Открытые обращения техподдержки</b>", reply_markup=admin_keyboard())
    for t in tickets:
        await message.answer(format_support_ticket(t) + f"\n\nОтветить: /replyticket {t['id']} ваш текст", reply_markup=support_ticket_actions_keyboard(t["id"]))


@router.message(Command("replyticket"))
async def cmd_reply_ticket(message: Message, bot: Bot):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split(maxsplit=2)
    if len(parts) < 3 or not parts[1].isdigit():
        await message.answer("Формат: /replyticket 1 текст ответа")
        return
    ticket = await add_support_message(int(parts[1]), message.from_user.id, "admin", parts[2], "answered")
    if not ticket:
        await message.answer("Обращение не найдено.")
        return
    await message.answer(f"✅ Ответ отправлен по обращению №{ticket['id']}.")
    await notify_user(
        bot,
        ticket["user_id"],
        f"🛠️ Ответ техподдержки по обращению №{ticket['id']}:\n\n{safe(parts[2])}\n\nЕсли вопрос не решён, нажмите «🛠️ Техподдержка» → «📋 Мои обращения» и ответьте в этот тикет.",
    )


@router.message(Command("closeticket"))
async def cmd_close_ticket(message: Message, bot: Bot):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split()
    if len(parts) != 2 or not parts[1].isdigit():
        await message.answer("Формат: /closeticket 1")
        return
    ticket = await close_support_ticket(int(parts[1]))
    if not ticket:
        await message.answer("Обращение не найдено.")
        return
    await message.answer(f"✅ Обращение №{ticket['id']} закрыто.")
    await notify_user(bot, ticket["user_id"], f"✅ Ваше обращение №{ticket['id']} закрыто. Спасибо за обращение.")


@router.message(Command("makepartner"))
async def cmd_makepartner(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split()
    if len(parts) != 2 or not parts[1].lstrip("-").isdigit():
        await message.answer("Формат: /makepartner 123456789")
        return
    user_id = int(parts[1])
    async with get_db() as db:
        await db.execute(
            "INSERT OR IGNORE INTO users (telegram_id, username, full_name, role, partner_code, created_at, updated_at) VALUES (?, '', '', 'partner', ?, ?, ?)",
            (user_id, partner_code(user_id), now_iso(), now_iso()),
        )
        await db.execute("UPDATE users SET role='partner', updated_at=? WHERE telegram_id=?", (now_iso(), user_id))
        await db.commit()
    await message.answer(f"✅ Пользователь {user_id} теперь партнёр.")


@router.message(Command("role"))
async def cmd_role(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split()
    if len(parts) != 3 or not parts[1].lstrip("-").isdigit() or parts[2] not in {"client", "partner", "warehouse", "courier", "admin"}:
        await message.answer("Формат: /role 123456789 warehouse|courier|client|partner|admin")
        return
    user_id = int(parts[1])
    role = parts[2]
    async with get_db() as db:
        await db.execute(
            "INSERT OR IGNORE INTO users (telegram_id, username, full_name, role, partner_code, created_at, updated_at) VALUES (?, '', '', ?, ?, ?, ?)",
            (user_id, role, partner_code(user_id), now_iso(), now_iso()),
        )
        await db.execute("UPDATE users SET role=?, updated_at=? WHERE telegram_id=?", (role, now_iso(), user_id))
        await db.commit()
    await message.answer(f"✅ Роль пользователя {user_id}: {role}")


# =========================
# CLIENT FLOWS
# =========================
@router.message(Command("history"))
async def cmd_history(message: Message):
    parts = (message.text or "").split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Формат: /history CG26050300001")
        return
    order = await get_order_by_code(parts[1].strip().upper())
    if not order:
        await message.answer("Груз не найден.")
        return
    if message.from_user and not is_admin(message.from_user.id) and order["user_id"] != message.from_user.id:
        await message.answer("⛔ Историю можно смотреть только по своему грузу.")
        return
    rows = await get_status_history(order["id"])
    await message.answer(f"<b>🕓 История {safe(order['tracking_code'])}</b>\n\n" + format_history(rows))


@router.message(Command("photos"))
async def cmd_photos(message: Message):
    parts = (message.text or "").split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Формат: /photos CG26050300001")
        return
    order = await get_order_by_code(parts[1].strip().upper())
    if not order:
        await message.answer("Груз не найден.")
        return
    if message.from_user and not is_admin(message.from_user.id) and order["user_id"] != message.from_user.id:
        await message.answer("⛔ Фото можно смотреть только по своему грузу.")
        return
    photos = await list_cargo_photos(order["id"])
    if not photos:
        await message.answer("По этому грузу пока нет фото.")
        return
    for ph in photos:
        await message.answer_photo(ph["file_id"], caption=f"📷 {safe(order['tracking_code'])}\n{safe(ph['comment'])}\n{safe((ph['created_at'] or '')[:16].replace('T', ' '))}")


@router.message(Command("qr"))
async def cmd_qr(message: Message, bot: Bot):
    parts = (message.text or "").split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Формат: /qr CG26050300001")
        return
    order = await get_order_by_code(parts[1].strip().upper())
    if not order:
        await message.answer("Груз не найден.")
        return
    if message.from_user and not is_admin(message.from_user.id) and order["user_id"] != message.from_user.id:
        await message.answer("⛔ QR можно получить только по своему грузу.")
        return
    me = await bot.get_me()
    payload = f"https://t.me/{me.username}?start=track_{order['tracking_code']}"
    qr_bytes = make_qr_bytes(payload)
    await message.answer_photo(BufferedInputFile(qr_bytes, filename=f"{order['tracking_code']}.png"), caption=f"🔳 QR для груза {safe(order['tracking_code'])}")


@router.message(Command("pay"))
async def cmd_pay(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split(maxsplit=3)
    if len(parts) < 3:
        await message.answer("Формат: /pay CG26050300001 10000 комментарий")
        return
    code = parts[1].upper()
    try:
        amount = float(parts[2].replace(",", "."))
    except ValueError:
        await message.answer("Сумма должна быть числом.")
        return
    comment = parts[3] if len(parts) >= 4 else "Оплата"
    order = await add_payment(code, amount, comment, message.from_user.id)
    if not order:
        await message.answer("Груз не найден.")
        return
    debt = max(float(order["price"] or 0) - float(order["paid_amount"] or 0), 0)
    await message.answer(
        f"✅ Оплата внесена.\n\n"
        f"Груз: {safe(order['tracking_code'])}\n"
        f"Цена: {order['price']} {safe(CURRENCY)}\n"
        f"Оплачено: {order['paid_amount']} {safe(CURRENCY)}\n"
        f"Долг: {debt} {safe(CURRENCY)}\n"
        f"Статус оплаты: {safe(order['payment_status'])}"
    )


@router.message(Command("debts"))
async def cmd_debts(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    rows = await list_debts()
    if not rows:
        await message.answer("Долгов нет.")
        return
    lines = []
    for o in rows:
        lines.append(f"{safe(o['tracking_code'])} | долг: {o['debt']} {safe(CURRENCY)} | клиент: {safe(o['customer_name'])} {safe(o['phone'])}")
    await message.answer("<b>💸 Неоплаченные заказы</b>\n\n" + "\n".join(lines))


@router.message(Command("export_orders"))
async def cmd_export_orders(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    path = await export_orders_to_excel()
    await message.answer_document(FSInputFile(path), caption="📤 Excel-выгрузка заказов")


@router.message(Command("settariff"))
async def cmd_settariff(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split(maxsplit=2)
    if len(parts) < 3:
        await message.answer("Формат: /settariff Китай 3.5")
        return
    try:
        rate = float(parts[2].replace(",", "."))
    except ValueError:
        await message.answer("Тариф должен быть числом.")
        return
    await set_tariff(parts[1], rate)
    await message.answer(f"✅ Тариф обновлён: {safe(parts[1])} — {rate} {safe(CURRENCY)}/кг")


@router.message(Command("tariffs"))
async def cmd_tariffs(message: Message):
    rows = await list_tariffs()
    if not rows:
        await message.answer("Тарифов пока нет.")
        return
    lines = [f"— {safe(r['country_name'])}: {r['rate']} {safe(CURRENCY)}/кг" for r in rows]
    await message.answer("<b>⚙️ Тарифы</b>\n\n" + "\n".join(lines))


@router.message(Command("receipt"))
async def cmd_receipt(message: Message):
    if not message.from_user:
        return
    parts = (message.text or "").split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Формат: /receipt CG26050300001")
        return
    order = await get_order_by_code(parts[1].strip().upper())
    if not order:
        await message.answer("Груз не найден.")
        return
    if not is_admin(message.from_user.id) and order["user_id"] != message.from_user.id:
        await message.answer("⛔ Можно получить квитанцию только по своему грузу.")
        return
    path = generate_receipt_pdf(order)
    await message.answer_document(FSInputFile(path), caption=f"📄 Квитанция {safe(order['tracking_code'])}")


@router.message(Command("assigncourier"))
async def cmd_assign_courier(message: Message, bot: Bot):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split(maxsplit=3)
    if len(parts) < 4 or not parts[2].lstrip("-").isdigit():
        await message.answer("Формат: /assigncourier CG26050300001 123456789 адрес доставки")
        return
    code = parts[1].upper()
    courier_id = int(parts[2])
    address = parts[3]
    order = await assign_courier(code, courier_id, address, message.from_user.id)
    if not order:
        await message.answer("Груз не найден.")
        return
    async with get_db() as db:
        await db.execute(
            "INSERT OR IGNORE INTO users (telegram_id, username, full_name, role, partner_code, created_at, updated_at) VALUES (?, '', '', 'courier', ?, ?, ?)",
            (courier_id, partner_code(courier_id), now_iso(), now_iso()),
        )
        await db.execute("UPDATE users SET role='courier', updated_at=? WHERE telegram_id=?", (now_iso(), courier_id))
        await db.commit()
    await message.answer("✅ Курьер назначен.\n\n" + format_order(order))
    await notify_user(bot, order["user_id"], f"🚚 Ваш груз <b>{safe(order['tracking_code'])}</b> передан курьеру. Адрес: {safe(address)}")
    await notify_user(bot, courier_id, f"🚚 Вам назначена доставка: {safe(order['tracking_code'])}\nАдрес: {safe(address)}")


@router.message(Command("courier"))
async def cmd_courier(message: Message, state: FSMContext):
    await state.clear()
    if not message.from_user or not (is_courier_static(message.from_user.id) or await has_role(message.from_user.id, "courier")):
        await message.answer("⛔ У вас нет доступа к курьерскому меню.")
        return
    await message.answer(courier_menu_text(), reply_markup=courier_keyboard())


@router.message(F.text == "👤 Клиентское меню")
async def show_client_menu(message: Message, state: FSMContext):
    await state.clear()
    name = message.from_user.first_name if message.from_user else ""
    await message.answer(client_menu_text(name), reply_markup=client_keyboard())


@router.message(F.text == "📦 Отправить груз")
@router.message(F.text == "📦 Оформить доставку")
@router.message(F.text == "📦 Оставить заявку")
async def cargo_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(CargoForm.from_country)
    await message.answer("Оформим заявку на доставку. Откуда груз? Например: Китай, Турция, Дубай, Россия")


@router.message(CargoForm.from_country)
async def cargo_from_country(message: Message, state: FSMContext):
    await state.update_data(from_country=message.text)
    await state.set_state(CargoForm.to_city)
    await message.answer("В какой город доставить? Например: Алматы, Астана, Шымкент")


@router.message(CargoForm.to_city)
async def cargo_to_city(message: Message, state: FSMContext):
    await state.update_data(to_city=message.text)
    await state.set_state(CargoForm.weight)
    await message.answer("Примерный вес в кг? Например: 12.5")


@router.message(CargoForm.weight)
async def cargo_weight(message: Message, state: FSMContext):
    try:
        weight = float((message.text or "").replace(",", "."))
        if weight <= 0:
            raise ValueError
    except ValueError:
        await message.answer("Введите вес числом. Например: 12.5")
        return
    await state.update_data(weight=weight)
    await state.set_state(CargoForm.volume)
    await message.answer("Объём в м³, если знаете. Если не знаете — напишите 0")


@router.message(CargoForm.volume)
async def cargo_volume(message: Message, state: FSMContext):
    try:
        volume = float((message.text or "0").replace(",", "."))
        if volume < 0:
            raise ValueError
    except ValueError:
        await message.answer("Введите объём числом. Если не знаете — 0")
        return
    await state.update_data(volume=volume)
    await state.set_state(CargoForm.cargo_type)
    await message.answer("Что за товар? Например: одежда, техника, косметика, запчасти")


@router.message(CargoForm.cargo_type)
async def cargo_type(message: Message, state: FSMContext):
    await state.update_data(cargo_type=message.text)
    await state.set_state(CargoForm.description)
    await message.answer("Добавьте комментарий: количество, особенности, нужна ли упаковка/забор товара.")


@router.message(CargoForm.description)
async def cargo_description(message: Message, state: FSMContext):
    await state.update_data(description=message.text)
    await state.set_state(CargoForm.customer_name)
    await message.answer("Ваше имя?")


@router.message(CargoForm.customer_name)
async def cargo_name(message: Message, state: FSMContext):
    await state.update_data(customer_name=message.text)
    await state.set_state(CargoForm.phone)
    await message.answer("Ваш телефон или WhatsApp/Telegram для связи?")


@router.message(CargoForm.phone)
async def cargo_phone(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user:
        return
    data = await state.get_data()
    await upsert_user(message)
    user = await get_user(message.from_user.id)
    partner_id = user["ref_partner_id"] if user else None
    order = await create_order(
        user_id=message.from_user.id,
        order_type="cargo",
        from_country=data["from_country"],
        to_city=data["to_city"],
        cargo_type=data["cargo_type"],
        weight=float(data["weight"]),
        volume=float(data["volume"]),
        description=data["description"],
        customer_name=data["customer_name"],
        phone=message.text or "",
        partner_id=partner_id,
    )
    rate, estimate = await estimate_delivery_price_db(data["from_country"], float(data["weight"]), float(data["volume"]))
    await state.clear()
    await message.answer(
        f"✅ Заявка создана.\n\n"
        f"Ваш номер груза: <b>{safe(order['tracking_code'])}</b>\n"
        f"Предварительный тариф: {rate} {safe(CURRENCY)}/кг\n"
        f"Предварительная стоимость: около <b>{estimate} {safe(CURRENCY)}</b>\n\n"
        "Итоговая цена может измениться после взвешивания и проверки груза на складе.",
        reply_markup=client_keyboard(),
    )
    await show_order_to_admin(bot, order)


@router.message(F.text == "🧮 Рассчитать доставку")
async def calc_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(CalcForm.from_country)
    await message.answer("Откуда груз? Китай, Турция, Дубай или Россия?")


@router.message(CalcForm.from_country)
async def calc_country(message: Message, state: FSMContext):
    await state.update_data(from_country=message.text)
    await state.set_state(CalcForm.weight)
    await message.answer("Вес в кг?")


@router.message(CalcForm.weight)
async def calc_weight(message: Message, state: FSMContext):
    try:
        weight = float((message.text or "").replace(",", "."))
        if weight <= 0:
            raise ValueError
    except ValueError:
        await message.answer("Введите вес числом. Например: 10")
        return
    await state.update_data(weight=weight)
    await state.set_state(CalcForm.volume)
    await message.answer("Объём в м³, если знаете. Если не знаете — напишите 0")


@router.message(CalcForm.volume)
async def calc_finish(message: Message, state: FSMContext):
    try:
        volume = float((message.text or "0").replace(",", "."))
    except ValueError:
        await message.answer("Введите объём числом. Если не знаете — 0")
        return
    data = await state.get_data()
    rate, estimate = await estimate_delivery_price_db(data["from_country"], float(data["weight"]), volume)
    await state.clear()
    await message.answer(
        f"🧮 <b>Предварительный расчёт</b>\n\n"
        f"Страна: {safe(data['from_country'])}\n"
        f"Вес: {safe(data['weight'])} кг\n"
        f"Объём: {safe(volume)} м³\n"
        f"Тариф: {rate} {safe(CURRENCY)}/кг\n"
        f"Примерная стоимость: <b>{estimate} {safe(CURRENCY)}</b>\n\n"
        "Это предварительный расчёт. Финальная цена зависит от фактического веса, категории товара и условий доставки.",
        reply_markup=client_keyboard(),
    )


@router.message(F.text == "🔎 Где мой груз?")
async def track_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(TrackForm.code)
    await message.answer("Введите номер груза. Например: CG26050300001")


@router.message(TrackForm.code)
async def track_finish(message: Message, state: FSMContext):
    code = parse_tracking_code(message.text or "") or (message.text or "").strip().upper()
    order = await get_order_by_code(code)
    await state.clear()
    if not order:
        await message.answer("Груз не найден. Проверьте номер или напишите менеджеру.", reply_markup=client_keyboard())
        return
    await message.answer(format_order(order), reply_markup=client_keyboard())


@router.message(F.text == "📋 Мои заказы")
async def my_orders(message: Message):
    if not message.from_user:
        return
    orders = await list_orders(limit=10, user_id=message.from_user.id)
    if not orders:
        await message.answer("У вас пока нет заказов.", reply_markup=client_keyboard())
        return
    text = "<b>📋 Ваши последние заказы</b>\n\n" + "\n".join(format_short_order(o) for o in orders)
    await message.answer(text, reply_markup=client_keyboard())


@router.message(F.text == "🛒 Выкуп товара")
async def buyout_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(BuyoutForm.link)
    await message.answer("Отправьте ссылку на товар или описание товара, который нужно выкупить.")


@router.message(BuyoutForm.link)
async def buyout_link(message: Message, state: FSMContext):
    await state.update_data(link=message.text)
    await state.set_state(BuyoutForm.details)
    await message.answer("Укажите детали: размер, цвет, количество, бюджет, комментарий.")


@router.message(BuyoutForm.details)
async def buyout_details(message: Message, state: FSMContext):
    await state.update_data(details=message.text)
    await state.set_state(BuyoutForm.customer_name)
    await message.answer("Ваше имя?")


@router.message(BuyoutForm.customer_name)
async def buyout_name(message: Message, state: FSMContext):
    await state.update_data(customer_name=message.text)
    await state.set_state(BuyoutForm.phone)
    await message.answer("Телефон или Telegram для связи?")


@router.message(BuyoutForm.phone)
async def buyout_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user:
        return
    data = await state.get_data()
    await upsert_user(message)
    user = await get_user(message.from_user.id)
    partner_id = user["ref_partner_id"] if user else None
    description = f"Ссылка/товар: {data['link']}\nДетали: {data['details']}"
    order = await create_order(
        user_id=message.from_user.id,
        order_type="buyout",
        from_country="Китай/маркетплейс",
        to_city="уточнить",
        cargo_type="выкуп товара",
        weight=0,
        volume=0,
        description=description,
        customer_name=data["customer_name"],
        phone=message.text or "",
        partner_id=partner_id,
    )
    await state.clear()
    await message.answer(
        f"✅ Заявка на выкуп создана.\nНомер: <b>{safe(order['tracking_code'])}</b>\n\n"
        "Менеджер проверит товар, стоимость, комиссию и доставку.",
        reply_markup=client_keyboard(),
    )
    await show_order_to_admin(bot, order)


@router.message(F.text == "🏢 Оптовая доставка")
async def wholesale_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(WholesaleForm.from_country)
    await message.answer("Откуда партия? Например: Китай, Турция, Дубай")


@router.message(WholesaleForm.from_country)
async def wholesale_country(message: Message, state: FSMContext):
    await state.update_data(from_country=message.text)
    await state.set_state(WholesaleForm.to_city)
    await message.answer("Куда доставить партию?")


@router.message(WholesaleForm.to_city)
async def wholesale_city(message: Message, state: FSMContext):
    await state.update_data(to_city=message.text)
    await state.set_state(WholesaleForm.goods)
    await message.answer("Что за товар и какая партия?")


@router.message(WholesaleForm.goods)
async def wholesale_goods(message: Message, state: FSMContext):
    await state.update_data(goods=message.text)
    await state.set_state(WholesaleForm.boxes)
    await message.answer("Сколько коробок/мест примерно?")


@router.message(WholesaleForm.boxes)
async def wholesale_boxes(message: Message, state: FSMContext):
    await state.update_data(boxes=message.text)
    await state.set_state(WholesaleForm.weight)
    await message.answer("Общий вес примерно в кг?")


@router.message(WholesaleForm.weight)
async def wholesale_weight(message: Message, state: FSMContext):
    try:
        weight = float((message.text or "").replace(",", "."))
    except ValueError:
        await message.answer("Введите вес числом. Например: 250")
        return
    await state.update_data(weight=weight)
    await state.set_state(WholesaleForm.documents)
    await message.answer("Нужны документы/таможенное сопровождение? Напишите что есть: инвойс, упаковочный лист, договор и т.д.")


@router.message(WholesaleForm.documents)
async def wholesale_docs(message: Message, state: FSMContext):
    await state.update_data(documents=message.text)
    await state.set_state(WholesaleForm.customer_name)
    await message.answer("Ваше имя?")


@router.message(WholesaleForm.customer_name)
async def wholesale_name(message: Message, state: FSMContext):
    await state.update_data(customer_name=message.text)
    await state.set_state(WholesaleForm.phone)
    await message.answer("Телефон или Telegram для связи?")


@router.message(WholesaleForm.phone)
async def wholesale_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user:
        return
    data = await state.get_data()
    await upsert_user(message)
    user = await get_user(message.from_user.id)
    partner_id = user["ref_partner_id"] if user else None
    description = (
        f"Товар: {data['goods']}\n"
        f"Коробки/места: {data['boxes']}\n"
        f"Документы: {data['documents']}"
    )
    order = await create_order(
        user_id=message.from_user.id,
        order_type="wholesale",
        from_country=data["from_country"],
        to_city=data["to_city"],
        cargo_type="оптовая партия",
        weight=float(data["weight"]),
        volume=0,
        description=description,
        customer_name=data["customer_name"],
        phone=message.text or "",
        partner_id=partner_id,
    )
    await state.clear()
    await message.answer(
        f"✅ Заявка на оптовую доставку создана.\nНомер: <b>{safe(order['tracking_code'])}</b>\n\n"
        "Менеджер проверит маршрут, документы, тариф и условия.",
        reply_markup=client_keyboard(),
    )
    await show_order_to_admin(bot, order)


@router.message(F.text == "⚠️ Жалоба / проблема")
async def complaint_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(ComplaintForm.code)
    await message.answer("Введите номер груза. Если номера нет — напишите 0")


@router.message(ComplaintForm.code)
async def complaint_code(message: Message, state: FSMContext):
    code = "" if (message.text or "").strip() == "0" else (parse_tracking_code(message.text or "") or (message.text or "").strip().upper())
    await state.update_data(code=code)
    await state.set_state(ComplaintForm.text)
    await message.answer("Опишите проблему коротко и понятно.")


@router.message(ComplaintForm.text)
async def complaint_text(message: Message, state: FSMContext):
    await state.update_data(text=message.text)
    await state.set_state(ComplaintForm.urgency)
    await message.answer("Срочность: низкая / средняя / высокая")


@router.message(ComplaintForm.urgency)
async def complaint_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user:
        return
    data = await state.get_data()
    complaint = await create_complaint(message.from_user.id, data.get("code", ""), data.get("text", ""), message.text or "средняя")
    await state.clear()
    await message.answer(
        f"✅ Жалоба №{complaint['id']} создана. Менеджер увидит её в админ‑панели.",
        reply_markup=client_keyboard(),
    )
    await notify_admins(
        bot,
        f"<b>⚠️ Новая жалоба №{complaint['id']}</b>\n"
        f"Груз: {safe(complaint['tracking_code'])}\n"
        f"Срочность: {safe(complaint['urgency'])}\n"
        f"Клиент ID: {complaint['user_id']}\n\n"
        f"{safe(complaint['text'])}",
        complaint_actions_keyboard(complaint["id"]),
    )



@router.message(F.text == "🛠️ Техподдержка")
async def support_start(message: Message, state: FSMContext):
    await state.clear()
    if message.from_user and is_admin(message.from_user.id):
        tickets = await list_open_support_tickets(20)
        if not tickets:
            await message.answer("Открытых обращений в техподдержку нет.", reply_markup=admin_keyboard())
            return
        await message.answer("<b>🛠️ Открытые обращения техподдержки</b>", reply_markup=admin_keyboard())
        for t in tickets:
            await message.answer(format_support_ticket(t) + f"\n\nОтветить: /replyticket {t['id']} ваш текст", reply_markup=support_ticket_actions_keyboard(t["id"]))
        return
    await message.answer(
        "<b>🛠️ Техподдержка</b>\n\n"
        "Здесь можно создать обращение, если нужен ответ менеджера: вопрос по грузу, оплате, фото, повреждению или работе бота.\n\n"
        "Выберите действие:",
        reply_markup=support_client_menu_keyboard(),
    )


@router.callback_query(F.data == "support_new")
async def cb_support_new(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await state.set_state(SupportForm.code)
    await call.message.answer("Введите номер груза <code>CG...</code>. Если вопрос общий — напишите <b>0</b>.")
    await call.answer()


@router.message(SupportForm.code)
async def support_code(message: Message, state: FSMContext):
    code = "" if (message.text or "").strip() == "0" else (parse_tracking_code(message.text or "") or (message.text or "").strip().upper())
    await state.update_data(code=code)
    await state.set_state(SupportForm.topic)
    await message.answer("Выберите тему обращения:", reply_markup=support_topic_keyboard())


@router.callback_query(F.data.startswith("support_topic:"))
async def cb_support_topic(call: CallbackQuery, state: FSMContext):
    topic = call.data.split(":", 1)[1]
    await state.update_data(topic=topic)
    await state.set_state(SupportForm.text)
    await call.message.answer(
        f"Тема: <b>{safe(support_topic_label(topic))}</b>\n\n"
        "Опишите вопрос одним сообщением. Укажите детали, чтобы менеджер быстрее понял ситуацию."
    )
    await call.answer()


@router.message(SupportForm.text)
async def support_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user:
        return
    data = await state.get_data()
    await upsert_user(message)
    ticket = await create_support_ticket(
        message.from_user.id,
        data.get("code", ""),
        data.get("topic", "other"),
        message.text or "",
    )
    await state.clear()
    await message.answer(
        f"✅ Обращение №{ticket['id']} создано. Менеджер получил уведомление.\n\n"
        "Чтобы посмотреть свои обращения или дописать сообщение, нажмите «🛠️ Техподдержка» → «📋 Мои обращения».",
        reply_markup=client_keyboard(),
    )
    await notify_admins(
        bot,
        "<b>🛠️ Новое обращение в техподдержку</b>\n\n" + format_support_ticket(ticket),
        support_ticket_actions_keyboard(ticket["id"]),
    )


@router.callback_query(F.data == "support_my")
async def cb_support_my(call: CallbackQuery):
    tickets = await list_user_support_tickets(call.from_user.id, 10)
    if not tickets:
        await call.message.answer("У вас пока нет обращений в техподдержку.")
        await call.answer()
        return
    await call.message.answer("<b>📋 Ваши обращения</b>")
    for t in tickets:
        await call.message.answer(format_support_ticket(t), reply_markup=support_ticket_actions_keyboard(t["id"], include_client_reply=True))
    await call.answer()


@router.callback_query(F.data.startswith("support_client_reply:"))
async def cb_support_client_reply(call: CallbackQuery, state: FSMContext):
    ticket_id = int(call.data.split(":", 1)[1])
    ticket = await get_support_ticket(ticket_id)
    if not ticket or ticket["user_id"] != call.from_user.id:
        await call.answer("Это не ваше обращение", show_alert=True)
        return
    if ticket["status"] == "closed":
        await call.answer("Обращение уже закрыто", show_alert=True)
        return
    await state.clear()
    await state.update_data(ticket_id=ticket_id)
    await state.set_state(ClientSupportReplyForm.text)
    await call.message.answer(f"Напишите ответ в обращение №{ticket_id} одним сообщением.")
    await call.answer()


@router.message(ClientSupportReplyForm.text)
async def support_client_reply_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user:
        return
    data = await state.get_data()
    ticket_id = int(data["ticket_id"])
    ticket = await get_support_ticket(ticket_id)
    if not ticket or ticket["user_id"] != message.from_user.id:
        await state.clear()
        await message.answer("Обращение не найдено.", reply_markup=client_keyboard())
        return
    updated = await add_support_message(ticket_id, message.from_user.id, "client", message.text or "", "open")
    await state.clear()
    await message.answer(f"✅ Сообщение добавлено в обращение №{ticket_id}.", reply_markup=client_keyboard())
    await notify_admins(
        bot,
        "<b>🛠️ Новое сообщение в обращении</b>\n\n" + format_support_ticket(updated),
        support_ticket_actions_keyboard(ticket_id),
    )


@router.message(F.text == "❓ FAQ")
async def faq(message: Message):
    await message.answer(
        "<b>❓ FAQ</b>\n\n"
        "<b>Как узнать статус?</b> Нажмите «🔎 Где мой груз?» и введите номер CG...\n\n"
        "<b>Цена финальная?</b> Калькулятор даёт предварительный расчёт. Финальная цена зависит от фактического веса, объёма и категории товара.\n\n"
        "<b>Какие документы нужны?</b> Обычно нужны описание товара, количество, стоимость, инвойс/ссылка/упаковочный лист, если есть.\n\n"
        "<b>Что делать при проблеме?</b> Нажмите «⚠️ Жалоба / проблема», укажите номер груза и описание.",
        reply_markup=client_keyboard(),
    )


@router.message(F.text == "🤝 Партнёрка")
async def partner_menu(message: Message, bot: Bot):
    if not message.from_user:
        return
    await upsert_user(message)
    me = await bot.get_me()
    code = message.from_user.id
    link = f"https://t.me/{me.username}?start=partner_{code}"
    report = await partner_report(message.from_user.id)
    orders = report["orders"]
    commission = round((orders["revenue"] or 0) * DEFAULT_COMMISSION_PERCENT / 100, 2)
    await message.answer(
        f"<b>🤝 Партнёрская система</b>\n\n"
        f"Ваша ссылка:\n{safe(link)}\n\n"
        f"Приведено клиентов: {report['leads']['cnt']}\n"
        f"Заказов от ваших клиентов: {orders['cnt']}\n"
        f"Оборот: {orders['revenue']} {safe(CURRENCY)}\n"
        f"Примерная комиссия {DEFAULT_COMMISSION_PERCENT}%: {commission} {safe(CURRENCY)}",
        reply_markup=client_keyboard(),
    )


@router.message(F.text == "👤 Мой кабинет")
async def client_cabinet(message: Message):
    if not message.from_user:
        return
    await upsert_user(message)
    orders = await list_orders(limit=50, user_id=message.from_user.id)
    active = [o for o in orders if o["status"] != "доставлен"]
    debt = sum(max(float(o["price"] or 0) - float(o["paid_amount"] or 0), 0) for o in orders)
    await message.answer(
        f"<b>👤 Мой кабинет</b>\n\n"
        f"ID: {message.from_user.id}\n"
        f"Активных грузов: {len(active)}\n"
        f"Всего заказов: {len(orders)}\n"
        f"Долг: {debt} {safe(CURRENCY)}\n\n"
        "Для статуса нажмите «🔎 Где мой груз?». Для фото — «📷 Фото груза».",
        reply_markup=client_keyboard(),
    )


@router.message(F.text == "🕓 История статусов")
async def history_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(HistoryForm.code)
    await message.answer("Введите номер груза CG...")


@router.message(HistoryForm.code)
async def history_finish(message: Message, state: FSMContext):
    if not message.from_user:
        return
    code = parse_tracking_code(message.text or "") or (message.text or "").strip().upper()
    order = await get_order_by_code(code)
    await state.clear()
    if not order:
        await message.answer("Груз не найден.", reply_markup=client_keyboard())
        return
    if not is_admin(message.from_user.id) and order["user_id"] != message.from_user.id:
        await message.answer("⛔ Историю можно смотреть только по своему грузу.", reply_markup=client_keyboard())
        return
    rows = await get_status_history(order["id"])
    await message.answer(f"<b>🕓 История {safe(order['tracking_code'])}</b>\n\n" + format_history(rows), reply_markup=client_keyboard())


@router.message(F.text == "📷 Фото груза")
async def photo_view_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(PhotoViewForm.code)
    await message.answer("Введите номер груза CG...")


@router.message(PhotoViewForm.code)
async def photo_view_finish(message: Message, state: FSMContext):
    if not message.from_user:
        return
    code = parse_tracking_code(message.text or "") or (message.text or "").strip().upper()
    order = await get_order_by_code(code)
    await state.clear()
    if not order:
        await message.answer("Груз не найден.", reply_markup=client_keyboard())
        return
    if not is_admin(message.from_user.id) and order["user_id"] != message.from_user.id:
        await message.answer("⛔ Фото можно смотреть только по своему грузу.", reply_markup=client_keyboard())
        return
    photos = await list_cargo_photos(order["id"])
    if not photos:
        await message.answer("По этому грузу пока нет фото.", reply_markup=client_keyboard())
        return
    for ph in photos:
        await message.answer_photo(ph["file_id"], caption=f"📷 {safe(order['tracking_code'])}\n{safe(ph['comment'])}")


@router.message(F.text == "🔳 QR груза")
async def qr_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(QrForm.code)
    await message.answer("Введите номер груза CG...")


@router.message(QrForm.code)
async def qr_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user:
        return
    code = parse_tracking_code(message.text or "") or (message.text or "").strip().upper()
    order = await get_order_by_code(code)
    await state.clear()
    if not order:
        await message.answer("Груз не найден.", reply_markup=client_keyboard())
        return
    if not is_admin(message.from_user.id) and order["user_id"] != message.from_user.id:
        await message.answer("⛔ QR можно получить только по своему грузу.", reply_markup=client_keyboard())
        return
    me = await bot.get_me()
    payload = f"https://t.me/{me.username}?start=track_{order['tracking_code']}"
    await message.answer_photo(BufferedInputFile(make_qr_bytes(payload), filename=f"{order['tracking_code']}.png"), caption=f"🔳 QR для {safe(order['tracking_code'])}", reply_markup=client_keyboard())


@router.message(F.text == "💸 Долги/оплаты")
async def admin_debts_button(message: Message):
    await cmd_debts(message)


@router.message(F.text == "📤 Excel экспорт")
async def admin_export_button(message: Message):
    await cmd_export_orders(message)


@router.message(F.text == "📥 Обновить из Excel")
async def excel_import_start(message: Message, state: FSMContext):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    await state.clear()
    await state.set_state(ExcelImportForm.file)
    await message.answer(
        "Отправьте .xlsx файл с колонками:\n"
        "tracking_code, status, comment, price, cost, paid_amount\n\n"
        "Обязательная колонка только tracking_code. Остальные можно не заполнять."
    )


@router.message(ExcelImportForm.file, F.document)
async def excel_import_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user or not is_admin(message.from_user.id) or not message.document:
        return
    if not (message.document.file_name or "").lower().endswith(".xlsx"):
        await message.answer("Нужен файл .xlsx")
        return
    path = EXPORT_DIR / f"import_{message.from_user.id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    await bot.download(message.document, destination=path)
    ok, fail = await import_statuses_from_excel(path, message.from_user.id, bot)
    await state.clear()
    await message.answer(f"✅ Импорт завершён. Обновлено: {ok}, ошибок: {fail}", reply_markup=admin_keyboard())


@router.message(F.text == "⚙️ Тарифы")
async def tariffs_menu(message: Message, state: FSMContext):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    await state.clear()
    rows = await list_tariffs()
    lines = [f"— {safe(r['country_name'])}: {r['rate']} {safe(CURRENCY)}/кг" for r in rows]
    await state.set_state(TariffForm.country)
    await message.answer(
        "<b>⚙️ Тарифы</b>\n\n" + ("\n".join(lines) if lines else "Тарифов пока нет.") +
        "\n\nЧтобы добавить/изменить тариф, напишите страну. Например: Китай\nДля отмены: /admin"
    )


@router.message(TariffForm.country)
async def tariff_country(message: Message, state: FSMContext):
    await state.update_data(country=message.text)
    await state.set_state(TariffForm.rate)
    await message.answer("Введите тариф за 1 кг. Например: 3.5")


@router.message(TariffForm.rate)
async def tariff_rate(message: Message, state: FSMContext):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    try:
        rate = float((message.text or "").replace(",", "."))
    except ValueError:
        await message.answer("Введите число. Например: 3.5")
        return
    data = await state.get_data()
    await set_tariff(data["country"], rate)
    await state.clear()
    await message.answer(f"✅ Тариф сохранён: {safe(data['country'])} — {rate} {safe(CURRENCY)}/кг", reply_markup=admin_keyboard())


@router.message(F.text == "👥 Роли")
async def roles_info(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    await message.answer(
        "<b>👥 Роли сотрудников</b>\n\n"
        "Выдать роль:\n"
        "/role 123456789 warehouse\n"
        "/role 123456789 courier\n"
        "/role 123456789 partner\n"
        "/role 123456789 client\n\n"
        "Склад: /warehouse\nКурьер: /courier",
        reply_markup=admin_keyboard(),
    )


@router.message(F.text == "📄 PDF квитанция")
async def receipt_start(message: Message, state: FSMContext):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    await state.clear()
    await state.set_state(ReceiptForm.code)
    await message.answer("Введите номер груза CG...")


@router.message(ReceiptForm.code)
async def receipt_finish(message: Message, state: FSMContext):
    code = parse_tracking_code(message.text or "") or (message.text or "").strip().upper()
    order = await get_order_by_code(code)
    await state.clear()
    if not order:
        await message.answer("Груз не найден.", reply_markup=admin_keyboard())
        return
    path = generate_receipt_pdf(order)
    await message.answer_document(FSInputFile(path), caption=f"📄 Квитанция {safe(order['tracking_code'])}", reply_markup=admin_keyboard())


@router.message(F.text == "🚚 Курьеры")
async def courier_assign_start(message: Message, state: FSMContext):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    await state.clear()
    await state.set_state(CourierAssignForm.code)
    await message.answer("Введите номер груза, который нужно передать курьеру.")


@router.message(CourierAssignForm.code)
async def courier_assign_code(message: Message, state: FSMContext):
    code = parse_tracking_code(message.text or "") or (message.text or "").strip().upper()
    order = await get_order_by_code(code)
    if not order:
        await message.answer("Груз не найден. Введите номер ещё раз.")
        return
    await state.update_data(code=code)
    await state.set_state(CourierAssignForm.courier_id)
    await message.answer("Введите Telegram ID курьера.")


@router.message(CourierAssignForm.courier_id)
async def courier_assign_id(message: Message, state: FSMContext):
    if not (message.text or "").lstrip("-").isdigit():
        await message.answer("ID должен быть числом.")
        return
    await state.update_data(courier_id=int(message.text))
    await state.set_state(CourierAssignForm.address)
    await message.answer("Введите адрес доставки или комментарий для курьера.")


@router.message(CourierAssignForm.address)
async def courier_assign_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    data = await state.get_data()
    order = await assign_courier(data["code"], int(data["courier_id"]), message.text or "", message.from_user.id)
    await state.clear()
    if not order:
        await message.answer("Груз не найден.", reply_markup=admin_keyboard())
        return
    async with get_db() as db:
        cid = int(data["courier_id"])
        await db.execute(
            "INSERT OR IGNORE INTO users (telegram_id, username, full_name, role, partner_code, created_at, updated_at) VALUES (?, '', '', 'courier', ?, ?, ?)",
            (cid, partner_code(cid), now_iso(), now_iso()),
        )
        await db.execute("UPDATE users SET role='courier', updated_at=? WHERE telegram_id=?", (now_iso(), cid))
        await db.commit()
    await message.answer("✅ Курьер назначен.\n\n" + format_order(order), reply_markup=admin_keyboard())
    await notify_user(bot, order["user_id"], f"🚚 Ваш груз <b>{safe(order['tracking_code'])}</b> передан курьеру.")
    await notify_user(bot, int(data["courier_id"]), f"🚚 Вам назначена доставка: {safe(order['tracking_code'])}\nАдрес: {safe(message.text)}")


@router.message(F.text == "🚚 Курьерское меню")
async def show_courier_menu(message: Message, state: FSMContext):
    await cmd_courier(message, state)


@router.message(F.text == "🚚 Мои доставки")
async def courier_my_deliveries(message: Message):
    if not message.from_user or not (is_courier_static(message.from_user.id) or await has_role(message.from_user.id, "courier")):
        return
    rows = await courier_orders(message.from_user.id)
    if not rows:
        await message.answer("У вас нет активных доставок.", reply_markup=courier_keyboard())
        return
    text = "<b>🚚 Мои доставки</b>\n\n" + "\n".join(
        f"{safe(o['tracking_code'])} | {safe(o['delivery_address'])} | {safe(o['phone'])}" for o in rows
    )
    await message.answer(text, reply_markup=courier_keyboard())


@router.message(F.text == "✅ Отметить доставлено")
async def courier_delivered_start(message: Message, state: FSMContext):
    if not message.from_user or not (is_courier_static(message.from_user.id) or await has_role(message.from_user.id, "courier")):
        return
    await state.clear()
    await state.set_state(CourierDeliveredForm.code)
    await message.answer("Введите номер груза CG...")


@router.message(CourierDeliveredForm.code)
async def courier_delivered_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user:
        return
    code = parse_tracking_code(message.text or "") or (message.text or "").strip().upper()
    order = await get_order_by_code(code)
    await state.clear()
    if not order or int(order["courier_id"] or 0) != message.from_user.id:
        await message.answer("Груз не найден среди ваших доставок.", reply_markup=courier_keyboard())
        return
    updated = await update_order_status(order["id"], "доставлен", message.from_user.id, "Курьер отметил доставку")
    await message.answer("✅ Доставка отмечена.\n\n" + format_order(updated), reply_markup=courier_keyboard())
    await notify_user(bot, updated["user_id"], f"✅ Ваш груз <b>{safe(updated['tracking_code'])}</b> доставлен.")


# =========================
# ADMIN / WAREHOUSE MENUS
# =========================
@router.message(F.text == "🏭 Меню склада")
async def show_warehouse_menu(message: Message, state: FSMContext):
    await state.clear()
    if not message.from_user or not is_warehouse(message.from_user.id):
        await message.answer("⛔ Нет доступа.")
        return
    await message.answer(warehouse_menu_text(), reply_markup=warehouse_keyboard())


@router.message(F.text == "📥 Новые заявки")
async def admin_new_orders(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    orders = await list_orders(limit=10, status="новая заявка")
    if not orders:
        await message.answer("Новых заявок нет.", reply_markup=admin_keyboard())
        return
    for order in orders:
        await message.answer(format_order(order), reply_markup=order_actions_keyboard(order["id"]))


@router.message(F.text == "📦 Все грузы")
async def admin_all_orders(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    orders = await list_orders(limit=15)
    if not orders:
        await message.answer("Грузов пока нет.", reply_markup=admin_keyboard())
        return
    text = "<b>📦 Последние грузы</b>\n\n" + "\n".join(format_short_order(o) for o in orders)
    await message.answer(text, reply_markup=admin_keyboard())


@router.message(F.text == "🔎 Найти груз")
async def admin_find_start(message: Message, state: FSMContext):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    await state.clear()
    await state.set_state(TrackForm.code)
    await message.answer("Введите номер груза CG...")


@router.message(F.text == "🔁 Изменить статус")
async def status_start(message: Message, state: FSMContext):
    if not message.from_user or not (is_admin(message.from_user.id) or is_warehouse(message.from_user.id)):
        return
    await state.clear()
    await state.set_state(StatusForm.code)
    await message.answer("Введите номер груза CG...")


@router.message(StatusForm.code)
async def status_code(message: Message, state: FSMContext):
    code = parse_tracking_code(message.text or "") or (message.text or "").strip().upper()
    order = await get_order_by_code(code)
    if not order:
        await message.answer("Груз не найден. Введите номер ещё раз или /start для отмены.")
        return
    await state.update_data(order_id=order["id"], code=code)
    await state.set_state(StatusForm.status)
    await message.answer("Выберите статус:", reply_markup=status_keyboard(order["id"]))


@router.message(StatusForm.status)
async def status_manual_status(message: Message, state: FSMContext):
    await state.update_data(status=message.text)
    await state.set_state(StatusForm.comment)
    await message.answer("Комментарий к статусу. Если не нужен — напишите 0")


@router.message(StatusForm.comment)
async def status_manual_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user:
        return
    data = await state.get_data()
    comment = "" if (message.text or "").strip() == "0" else (message.text or "")
    updated = await update_order_status(data["order_id"], data["status"], message.from_user.id, comment)
    await state.clear()
    await message.answer("✅ Статус обновлён.\n\n" + format_order(updated), reply_markup=admin_keyboard() if is_admin(message.from_user.id) else warehouse_keyboard())
    await notify_user(
        bot,
        updated["user_id"],
        f"📦 Статус груза <b>{safe(updated['tracking_code'])}</b> обновлён: <b>{safe(updated['status'])}</b>\n{safe(comment)}",
    )


@router.message(F.text == "💬 Жалобы")
async def admin_complaints(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    complaints = await list_open_complaints(10)
    if not complaints:
        await message.answer("Открытых жалоб нет.", reply_markup=admin_keyboard())
        return
    for c in complaints:
        await message.answer(
            f"<b>⚠️ Жалоба №{c['id']}</b>\n"
            f"Груз: {safe(c['tracking_code'])}\n"
            f"Клиент ID: {c['user_id']}\n"
            f"Срочность: {safe(c['urgency'])}\n"
            f"Статус: {safe(c['status'])}\n\n"
            f"{safe(c['text'])}\n\n"
            f"Ответить: /replycomplaint {c['id']} ваш текст",
            reply_markup=complaint_actions_keyboard(c["id"]),
        )



@router.message(F.text == "🛠️ Техподдержка")
async def admin_support_menu(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    tickets = await list_open_support_tickets(20)
    if not tickets:
        await message.answer("Открытых обращений в техподдержку нет.", reply_markup=admin_keyboard())
        return
    await message.answer("<b>🛠️ Открытые обращения техподдержки</b>", reply_markup=admin_keyboard())
    for t in tickets:
        await message.answer(format_support_ticket(t) + f"\n\nОтветить: /replyticket {t['id']} ваш текст", reply_markup=support_ticket_actions_keyboard(t["id"]))


@router.message(F.text == "💰 Финансы")
@router.message(F.text == "📊 Отчёт за день")
async def admin_finance(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    report = await finance_report()
    totals = report["totals"]
    today = report["today"]
    status_lines = "\n".join(f"— {safe(r['status'])}: {r['cnt']}" for r in report["by_status"])
    await message.answer(
        f"<b>📊 Отчёт</b>\n\n"
        f"<b>Сегодня:</b>\n"
        f"Заявок: {today['today_orders']}\n"
        f"Оборот: {today['today_revenue']} {safe(CURRENCY)}\n"
        f"Себестоимость: {today['today_cost']} {safe(CURRENCY)}\n"
        f"Маржа: {today['today_margin']} {safe(CURRENCY)}\n\n"
        f"<b>Всего:</b>\n"
        f"Заявок: {totals['total_orders']}\n"
        f"Оборот: {totals['revenue']} {safe(CURRENCY)}\n"
        f"Себестоимость: {totals['cost']} {safe(CURRENCY)}\n"
        f"Маржа: {totals['margin']} {safe(CURRENCY)}\n\n"
        f"<b>По статусам:</b>\n{status_lines or 'нет данных'}\n\n"
        "Чтобы внести финансы по грузу: /setprice CG26050300001 100 70",
        reply_markup=admin_keyboard(),
    )


@router.message(F.text == "🤝 Партнёры")
async def admin_partners(message: Message):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    report = await partner_report(message.from_user.id)
    top = report["partners"]
    if not top:
        await message.answer("Партнёрских лидов пока нет.", reply_markup=admin_keyboard())
        return
    lines = [f"Партнёр {r['partner_id']}: {r['cnt']} клиентов" for r in top]
    await message.answer("<b>🤝 Партнёры</b>\n\n" + "\n".join(lines), reply_markup=admin_keyboard())


@router.message(F.text == "📢 Рассылка")
async def broadcast_start(message: Message, state: FSMContext):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    await state.clear()
    await state.set_state(BroadcastForm.text)
    await message.answer("Введите текст рассылки. Для отмены: /start")


@router.message(BroadcastForm.text)
async def broadcast_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    text = message.text or ""
    async with get_db() as db:
        users = await db_fetchall(db, "SELECT telegram_id FROM users WHERE role IN ('client', 'partner')")
    sent = 0
    failed = 0
    for user in users:
        try:
            await bot.send_message(user["telegram_id"], text)
            sent += 1
            await asyncio.sleep(0.04)
        except Exception:
            failed += 1
    await state.clear()
    await message.answer(f"✅ Рассылка завершена. Отправлено: {sent}, ошибок: {failed}", reply_markup=admin_keyboard())


@router.message(F.text == "📸 Добавить фото")
async def warehouse_photo_start(message: Message, state: FSMContext):
    if not message.from_user or not is_warehouse(message.from_user.id):
        return
    await state.clear()
    await state.set_state(PhotoAddForm.code)
    await message.answer("Введите номер груза CG...")


@router.message(PhotoAddForm.code)
async def warehouse_photo_code(message: Message, state: FSMContext):
    code = parse_tracking_code(message.text or "") or (message.text or "").strip().upper()
    order = await get_order_by_code(code)
    if not order:
        await message.answer("Груз не найден. Введите номер ещё раз.")
        return
    await state.update_data(code=code, order_id=order["id"])
    await state.set_state(PhotoAddForm.photo)
    await message.answer("Отправьте фото груза. В подписи можно написать комментарий: склад, упаковка, повреждение и т.д.")


@router.message(PhotoAddForm.photo, F.photo)
async def warehouse_photo_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user or not is_warehouse(message.from_user.id) or not message.photo:
        return
    data = await state.get_data()
    order = await get_order(data["order_id"])
    if not order:
        await state.clear()
        await message.answer("Груз не найден.", reply_markup=warehouse_keyboard())
        return
    file_id = message.photo[-1].file_id
    comment = message.caption or "Фото груза"
    await add_cargo_photo(order["id"], file_id, comment, message.from_user.id)
    await state.clear()
    await message.answer("✅ Фото добавлено к грузу.", reply_markup=warehouse_keyboard())
    await notify_user(bot, order["user_id"], f"📷 По вашему грузу <b>{safe(order['tracking_code'])}</b> добавлено новое фото.")


@router.message(PhotoAddForm.photo)
async def warehouse_photo_not_photo(message: Message):
    await message.answer("Отправьте именно фото, не файл. Можно добавить подпись к фото.")


@router.message(F.text == "➕ Принять груз")
@router.message(F.text == "⚖️ Указать вес")
async def warehouse_weight_start(message: Message, state: FSMContext):
    if not message.from_user or not is_warehouse(message.from_user.id):
        return
    await state.clear()
    await state.set_state(WeightForm.code)
    await message.answer("Введите номер груза CG...")


@router.message(WeightForm.code)
async def warehouse_weight_code(message: Message, state: FSMContext):
    code = parse_tracking_code(message.text or "") or (message.text or "").strip().upper()
    order = await get_order_by_code(code)
    if not order:
        await message.answer("Груз не найден. Введите номер ещё раз.")
        return
    await state.update_data(code=code)
    await state.set_state(WeightForm.weight)
    await message.answer("Укажите фактический вес в кг.")


@router.message(WeightForm.weight)
async def warehouse_weight_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user:
        return
    try:
        weight = float((message.text or "").replace(",", "."))
    except ValueError:
        await message.answer("Введите вес числом. Например: 8.4")
        return
    data = await state.get_data()
    updated = await set_order_weight(data["code"], weight, message.from_user.id)
    if not updated:
        await message.answer("Груз не найден.")
        await state.clear()
        return
    updated = await update_order_status(updated["id"], "принят на склад", message.from_user.id, f"Вес: {weight} кг")
    await state.clear()
    await message.answer("✅ Груз принят/вес обновлён.\n\n" + format_order(updated), reply_markup=warehouse_keyboard())
    await notify_user(
        bot,
        updated["user_id"],
        f"📦 Ваш груз <b>{safe(updated['tracking_code'])}</b> принят на склад.\nФактический вес: <b>{weight} кг</b>.",
    )


@router.message(F.text == "📦 Грузы на складе")
async def warehouse_orders(message: Message):
    if not message.from_user or not is_warehouse(message.from_user.id):
        return
    rows = await orders_on_warehouse()
    if not rows:
        await message.answer("На складе пока нет активных грузов.", reply_markup=warehouse_keyboard())
        return
    text = "<b>📦 Грузы на складе</b>\n\n" + "\n".join(format_short_order(o) for o in rows)
    await message.answer(text, reply_markup=warehouse_keyboard())


# =========================
# CALLBACKS
# =========================
@router.callback_query(F.data == "admin_back")
async def cb_admin_back(call: CallbackQuery):
    if not call.from_user or not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    await call.message.answer(admin_menu_text(), reply_markup=admin_keyboard())
    await call.answer()


@router.callback_query(F.data == "auto_enable_all")
async def cb_auto_enable_all(call: CallbackQuery):
    if not call.from_user or not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    count = await enable_auto_status_for_active_orders(call.from_user.id)
    orders = await list_orders(limit=10)
    await call.message.answer(
        f"✅ Автостатусы включены для активных грузов: <b>{count}</b>.\n\n"
        "Бот сам выбрал маршруты по направлениям и будет менять статусы по расписанию. "
        "Клиенты получат уведомления автоматически.",
        reply_markup=auto_status_panel_keyboard(orders),
    )
    await call.answer("Готово")


@router.callback_query(F.data == "auto_show_disabled")
async def cb_auto_show_disabled(call: CallbackQuery):
    if not call.from_user or not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    orders = await list_orders_without_auto_status(limit=10)
    if not orders:
        await call.message.answer(
            "✅ У всех активных грузов уже включены автостатусы.",
            reply_markup=auto_status_panel_keyboard(await list_orders(limit=10)),
        )
    else:
        await call.message.answer(
            "<b>📋 Грузы без автостатусов</b>\n\n"
            "Нажмите на нужный груз — дальше можно включить автостатусы одной кнопкой.",
            reply_markup=auto_status_panel_keyboard(orders),
        )
    await call.answer()


@router.callback_query(F.data.startswith("auto_enable:"))
async def cb_auto_enable(call: CallbackQuery):
    if not call.from_user or not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    order_id = int(call.data.split(":", 1)[1])
    updated = await enable_auto_status_for_order(order_id, call.from_user.id)
    if not updated:
        await call.answer("Груз не найден", show_alert=True)
        return
    route_name = auto_route_label(updated["auto_status_route"] or choose_auto_route(updated["from_country"] or "", updated["to_city"] or ""))
    await call.message.answer(
        f"✅ Автостатусы включены для <b>{safe(updated['tracking_code'])}</b>.\n"
        f"Маршрут выбран автоматически: <b>{safe(route_name)}</b>.\n\n"
        "Бот сам будет менять статусы по расписанию маршрута и уведомлять клиента.",
        reply_markup=order_actions_keyboard(order_id),
    )
    await call.answer("Автостатусы включены")


@router.callback_query(F.data.startswith("auto_disable:"))
async def cb_auto_disable(call: CallbackQuery):
    if not call.from_user or not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    order_id = int(call.data.split(":", 1)[1])
    updated = await disable_auto_status_for_order(order_id, call.from_user.id)
    if not updated:
        await call.answer("Груз не найден", show_alert=True)
        return
    await call.message.answer(f"⏸ Автостатусы для <b>{safe(updated['tracking_code'])}</b> выключены.")
    await call.answer("Автостатусы выключены")


@router.callback_query(F.data.startswith("open_auto:"))
async def cb_open_auto_status(call: CallbackQuery):
    if not call.from_user or not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    order_id = int(call.data.split(":", 1)[1])
    order = await get_order(order_id)
    if not order:
        await call.answer("Груз не найден", show_alert=True)
        return
    route_key = order["auto_status_route"] or choose_auto_route(order["from_country"] or "", order["to_city"] or "")
    await call.message.answer(
        f"<b>🤖 Автостатусы для {safe(order['tracking_code'])}</b>\n\n"
        f"Сейчас: {'включены' if int(order['auto_status_enabled'] or 0) else 'выключены'}\n"
        f"Маршрут: {safe(auto_route_label(route_key))}\n\n"
        "Нажмите кнопку ниже, чтобы бот сам вёл этот груз по статусам.",
        reply_markup=auto_status_keyboard(order_id),
    )
    await call.answer()


@router.callback_query(F.data.startswith("open_status:"))
async def cb_open_status(call: CallbackQuery):
    if not call.from_user or not (is_admin(call.from_user.id) or is_warehouse(call.from_user.id)):
        await call.answer("Нет доступа", show_alert=True)
        return
    order_id = int(call.data.split(":", 1)[1])
    await call.message.answer("Выберите новый статус:", reply_markup=status_keyboard(order_id))
    await call.answer()


@router.callback_query(F.data.startswith("st:"))
async def cb_set_status(call: CallbackQuery, bot: Bot, state: FSMContext):
    if not call.from_user or not (is_admin(call.from_user.id) or is_warehouse(call.from_user.id)):
        await call.answer("Нет доступа", show_alert=True)
        return
    _, order_id_raw, status = call.data.split(":", 2)
    order_id = int(order_id_raw)
    updated = await update_order_status(order_id, status, call.from_user.id, "Статус изменён через кнопку")
    if not updated:
        await call.answer("Груз не найден", show_alert=True)
        return
    await state.clear()
    await call.message.answer("✅ Статус обновлён.\n\n" + format_order(updated))
    await notify_user(
        bot,
        updated["user_id"],
        f"📦 Статус груза <b>{safe(updated['tracking_code'])}</b> обновлён: <b>{safe(status)}</b>",
    )
    await call.answer("Статус обновлён")



@router.callback_query(F.data.startswith("support_reply:"))
async def cb_support_reply(call: CallbackQuery, state: FSMContext):
    if not call.from_user or not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    ticket_id = int(call.data.split(":", 1)[1])
    ticket = await get_support_ticket(ticket_id)
    if not ticket:
        await call.answer("Обращение не найдено", show_alert=True)
        return
    await state.clear()
    await state.update_data(ticket_id=ticket_id)
    await state.set_state(AdminSupportReplyForm.text)
    await call.message.answer(f"Напишите ответ клиенту по обращению №{ticket_id} одним сообщением.")
    await call.answer()


@router.message(AdminSupportReplyForm.text)
async def admin_support_reply_finish(message: Message, state: FSMContext, bot: Bot):
    if not message.from_user or not is_admin(message.from_user.id):
        return
    data = await state.get_data()
    ticket_id = int(data["ticket_id"])
    ticket = await add_support_message(ticket_id, message.from_user.id, "admin", message.text or "", "answered")
    await state.clear()
    if not ticket:
        await message.answer("Обращение не найдено.", reply_markup=admin_keyboard())
        return
    await message.answer(f"✅ Ответ отправлен клиенту по обращению №{ticket_id}.", reply_markup=admin_keyboard())
    await notify_user(
        bot,
        ticket["user_id"],
        f"🛠️ Ответ техподдержки по обращению №{ticket_id}:\n\n{safe(message.text)}\n\nЕсли вопрос не решён, нажмите «🛠️ Техподдержка» → «📋 Мои обращения» и ответьте в этот тикет.",
    )


@router.callback_query(F.data.startswith("support_close:"))
async def cb_support_close(call: CallbackQuery, bot: Bot):
    if not call.from_user or not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    ticket_id = int(call.data.split(":", 1)[1])
    ticket = await close_support_ticket(ticket_id)
    if not ticket:
        await call.answer("Обращение не найдено", show_alert=True)
        return
    await notify_user(bot, ticket["user_id"], f"✅ Ваше обращение №{ticket_id} закрыто. Спасибо за обращение.")
    await call.answer("Обращение закрыто")
    try:
        await call.message.edit_text((call.message.text or "") + "\n\n✅ Закрыто")
    except TelegramBadRequest:
        pass


@router.callback_query(F.data.startswith("close_complaint:"))
async def cb_close_complaint(call: CallbackQuery):
    if not call.from_user or not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    complaint_id = int(call.data.split(":", 1)[1])
    await close_complaint(complaint_id)
    await call.answer("Жалоба закрыта")
    try:
        await call.message.edit_text((call.message.text or "") + "\n\n✅ Закрыто")
    except TelegramBadRequest:
        pass


# =========================
# FREE AI / INTENT FALLBACK
# =========================
@router.message(F.text)
async def free_ai_router(message: Message, state: FSMContext):
    # This is the free AI-like layer: no paid LLM API, just intent detection + scenarios.
    if not message.from_user:
        return
    await upsert_user(message)
    text = message.text or ""
    code = parse_tracking_code(text)
    if code:
        order = await get_order_by_code(code)
        if order:
            await message.answer(format_order(order), reply_markup=client_keyboard())
            return

    intent = detect_intent(text)
    if intent == "tracking":
        await track_start(message, state)
        return
    if intent == "calculator":
        await calc_start(message, state)
        return
    if intent == "buyout":
        await buyout_start(message, state)
        return
    if intent == "wholesale":
        await wholesale_start(message, state)
        return
    if intent == "support":
        await support_start(message, state)
        return
    if intent == "complaint":
        await complaint_start(message, state)
        return
    if intent == "customs":
        await message.answer(
            "📄 <b>Чек‑лист для документов</b>\n\n"
            "Обычно нужны:\n"
            "— описание товара\n"
            "— количество\n"
            "— примерная стоимость\n"
            "— страна отправки\n"
            "— инвойс или ссылка на товар\n"
            "— упаковочный лист, если есть\n\n"
            "Я не оформляю юридическое решение по таможне, но могу собрать данные для менеджера.",
            reply_markup=client_keyboard(),
        )
        return

    await message.answer(
        "Я могу помочь с доставкой, расчётом, статусом груза, выкупом товара, оптовой партией или жалобой.\n\n"
        "Выберите действие в меню 👇",
        reply_markup=client_keyboard(),
    )



# =========================
# FREE AUTO STATUS ENGINE
# =========================
async def process_auto_statuses(bot: Bot) -> int:
    if not AUTO_STATUS_ENABLED:
        return 0
    async with get_db() as db:
        rows = await db_fetchall(
            db,
            """
            SELECT * FROM orders
            WHERE COALESCE(auto_status_enabled, 0)=1
              AND status NOT IN ('доставлен', 'отменён', 'проблема', 'передан курьеру')
            ORDER BY id ASC
            LIMIT 100
            """,
        )
    changed = 0
    for order in rows:
        try:
            target = get_auto_target_status(order)
            if not target:
                continue
            status, comment, route_name = target
            updated = await update_order_status(order["id"], status, 0, f"Автостатус: {comment}. Маршрут: {route_name}")
            if not updated:
                continue
            async with get_db() as db:
                await db.execute("UPDATE orders SET auto_status_last_at=? WHERE id=?", (now_iso(), order["id"]))
                await db.commit()
            changed += 1
            await notify_user(
                bot,
                updated["user_id"],
                f"📦 Статус груза <b>{safe(updated['tracking_code'])}</b> обновлён: <b>{safe(status)}</b>\n{safe(comment)}",
            )
        except Exception as e:
            logger.exception("Auto status failed for order %s: %s", order["id"], e)
    return changed


async def auto_status_loop(bot: Bot) -> None:
    await asyncio.sleep(15)
    logger.info("Auto status engine: enabled=%s on_new=%s interval=%s unit=%s", AUTO_STATUS_ENABLED, AUTO_STATUS_ON_NEW_ORDERS, AUTO_STATUS_INTERVAL_SECONDS, AUTO_STATUS_TIME_UNIT)
    while True:
        try:
            changed = await process_auto_statuses(bot)
            if changed:
                logger.info("Auto status engine changed %s orders", changed)
        except Exception as e:
            logger.exception("Auto status loop error: %s", e)
        # Для реальной работы обычно 300 сек. Для демо можно поставить 10-30 сек.
        await asyncio.sleep(max(10, AUTO_STATUS_INTERVAL_SECONDS))


# =========================
# HEALTH SERVER
# =========================
async def health(_: web.Request) -> web.Response:
    return web.json_response({"ok": True, "service": COMPANY_NAME})


async def start_health_server(bot: Bot) -> None:
    # Нужен только для health-check на Render/Koyeb/VPS. Веб-панели в этой версии нет.
    app = web.Application()
    app["bot"] = bot
    app.router.add_get("/", health)
    app.router.add_get("/health", health)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    logger.info("Health server started on port %s", PORT)


async def main() -> None:
    await init_db()
    bot = Bot(BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    asyncio.create_task(start_health_server(bot))
    asyncio.create_task(auto_status_loop(bot))
    logger.info("Starting polling. DB=%s", DB_PATH)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
