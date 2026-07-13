import os
import io
import json
import base64
import logging
from datetime import datetime, timedelta

import qrcode
import gspread
from google.oauth2.service_account import Credentials
from apscheduler.schedulers.asyncio import AsyncIOScheduler

from contract import generate_contract_pdf, ru_date, extract_serial
from datetime import date as date_cls

from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandStart
from aiogram.types import (
    Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton, BufferedInputFile,
    ReplyKeyboardMarkup, KeyboardButton, BotCommand, BotCommandScopeDefault, BotCommandScopeChat
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("bot")

# ------------------ НАСТРОЙКИ ------------------
BOT_TOKEN = os.environ["BOT_TOKEN"]
ADMIN_ID = int(os.environ["ADMIN_ID"])
SHEET_ID = os.environ["SHEET_ID"]
_friend_env = os.environ.get("FRIEND_ID")
FRIEND_ID = int(_friend_env) if _friend_env else None

OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
openai_client = None
if OPENAI_API_KEY:
    from openai import AsyncOpenAI
    openai_client = AsyncOpenAI(api_key=OPENAI_API_KEY)

BOT_USERNAME = "elektro_vlg_bot"
BUSINESS_ADDRESS = "г. Волгоград, ул. Киргизская 2"
BUSINESS_PHONE = "8-960-896-06-06"
PAYMENT_INFO = "💳 Оплата на номер +7 995 404-39-63 (Сбербанк, Емикова Наталья Анатольевна)"
DELIVERY_PRICE = 500
TO_INTERVAL_DAYS = 30

creds_dict = json.loads(os.environ["GOOGLE_CREDS"])
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
gc = gspread.authorize(credentials)
sh = gc.open_by_key(SHEET_ID)

bikes_ws = sh.worksheet("Велосипеды")
clients_ws = sh.worksheet("Клиенты")
rentals_ws = sh.worksheet("Аренды")
finances_ws = sh.worksheet("Финансы")

PERIOD_DAYS = {"неделя": 7, "месяц": 30}


def ensure_column(ws, column_name):
    headers = ws.row_values(1)
    if column_name not in headers:
        ws.update_cell(1, len(headers) + 1, column_name)


def ensure_worksheet(title, headers):
    try:
        ws = sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=200, cols=len(headers))
        ws.append_row(headers)
    return ws


def ensure_all_columns():
    ensure_column(clients_ws, "blacklist")
    ensure_column(rentals_ws, "delivery")
    ensure_column(rentals_ws, "delivery_address")
    ensure_column(finances_ws, "bike_id")
    ensure_column(bikes_ws, "photo_id")
    ensure_column(clients_ws, "passport")
    ensure_column(clients_ws, "passport_series")
    ensure_column(clients_ws, "passport_number")
    ensure_column(clients_ws, "registration_address")
    ensure_column(clients_ws, "passport_photo_main")
    ensure_column(clients_ws, "passport_photo_reg")
    ensure_column(clients_ws, "dob")
    ensure_column(clients_ws, "passport_issued_by")
    ensure_column(clients_ws, "passport_issue_date")
    ensure_column(clients_ws, "department_code")
    ensure_column(clients_ws, "actual_address")


ensure_all_columns()
waiting_ws = ensure_worksheet("Ожидание", ["id", "telegram_id", "name", "date_added"])
contracts_ws = ensure_worksheet("Договоры", ["id", "rental_id", "client_id", "date", "file_id"])

# ------------------ ВЕЛОСИПЕДЫ ------------------

def get_free_bikes():
    rows = bikes_ws.get_all_records()
    return [r for r in rows if str(r.get("status", "")).strip() == "свободен"]


def get_bike_by_id(bike_id):
    rows = bikes_ws.get_all_records()
    for i, r in enumerate(rows, start=2):
        if str(r.get("id")) == str(bike_id):
            return r, i
    return None, None


def set_bike_status(bike_id, status):
    _, row_idx = get_bike_by_id(bike_id)
    if row_idx:
        col = bikes_ws.find("status").col
        bikes_ws.update_cell(row_idx, col, status)


def set_bike_photo(bike_id, file_id):
    _, row_idx = get_bike_by_id(bike_id)
    if row_idx:
        col = bikes_ws.find("photo_id").col
        bikes_ws.update_cell(row_idx, col, file_id)
        return True
    return False


# ------------------ КЛИЕНТЫ ------------------

def get_or_create_client(tg_id, name, phone, passport_series="", passport_number="",
                          registration_address="", photo_main="", photo_reg="",
                          dob="", issued_by="", issue_date="", department_code="",
                          actual_address=""):
    rows = clients_ws.get_all_records()
    for r in rows:
        if str(r.get("telegram_id")) == str(tg_id):
            return r.get("id")
    new_id = len(rows) + 1
    clients_ws.append_row([
        new_id, tg_id, name, phone, datetime.now().strftime("%d.%m.%Y"), "нет", "",
        passport_series, passport_number, registration_address, photo_main, photo_reg,
        dob, issued_by, issue_date, department_code, actual_address
    ])
    return new_id


def get_client_by_tg(tg_id):
    rows = clients_ws.get_all_records()
    for r in rows:
        if str(r.get("telegram_id")) == str(tg_id):
            return r
    return None


def get_client_telegram_id(client_id):
    rows = clients_ws.get_all_records()
    for r in rows:
        if str(r.get("id")) == str(client_id):
            return r.get("telegram_id")
    return None


def get_client_by_id(client_id):
    rows = clients_ws.get_all_records()
    for r in rows:
        if str(r.get("id")) == str(client_id):
            return r
    return None


def is_manual_tg_id(tg_id):
    return str(tg_id).startswith("manual_")


def next_manual_id():
    rows = clients_ws.get_all_records()
    manual_count = len([r for r in rows if is_manual_tg_id(r.get("telegram_id", ""))])
    return f"manual_{manual_count + 1}"


def create_manual_client(name, phone, passport_series="", passport_number="",
                          registration_address="", photo_main="", photo_reg="",
                          dob="", issued_by="", issue_date="", department_code="",
                          actual_address=""):
    rows = clients_ws.get_all_records()
    new_id = len(rows) + 1
    tg_id = next_manual_id()
    clients_ws.append_row([
        new_id, tg_id, name, phone, datetime.now().strftime("%d.%m.%Y"), "нет", "",
        passport_series, passport_number, registration_address, photo_main, photo_reg,
        dob, issued_by, issue_date, department_code, actual_address
    ])
    return new_id, tg_id


def get_manual_clients():
    rows = clients_ws.get_all_records()
    return [r for r in rows if is_manual_tg_id(r.get("telegram_id", ""))]


def is_blacklisted(client_id):
    rows = clients_ws.get_all_records()
    for r in rows:
        if str(r.get("id")) == str(client_id):
            return str(r.get("blacklist", "")).strip() == "да"
    return False


def get_all_clients():
    return clients_ws.get_all_records()


# ------------------ ЛИСТ ОЖИДАНИЯ ------------------

def add_to_waitlist(tg_id, name):
    rows = waiting_ws.get_all_records()
    for r in rows:
        if str(r.get("telegram_id")) == str(tg_id):
            return False
    new_id = len(rows) + 1
    waiting_ws.append_row([new_id, tg_id, name, datetime.now().strftime("%d.%m.%Y %H:%M")])
    return True


def get_waitlist():
    return waiting_ws.get_all_records()


def clear_waitlist():
    waiting_ws.clear()
    waiting_ws.append_row(["id", "telegram_id", "name", "date_added"])


# ------------------ АРЕНДЫ ------------------

def create_rental(bike_id, client_id, period, price, owner, delivery="нет", delivery_address=""):
    rows = rentals_ws.get_all_records()
    new_id = len(rows) + 1
    start = datetime.now()
    end = start + timedelta(days=PERIOD_DAYS[period])
    rentals_ws.append_row([
        new_id, bike_id, client_id,
        start.strftime("%d.%m.%Y"), period, end.strftime("%d.%m.%Y"),
        price, owner, "не оплачено", "арендован", delivery, delivery_address
    ])
    set_bike_status(bike_id, "в аренде")
    return new_id, start, end


def get_rental_by_id(rental_id):
    rows = rentals_ws.get_all_records()
    for i, r in enumerate(rows, start=2):
        if str(r.get("id")) == str(rental_id):
            return r, i
    return None, None


def get_active_rentals():
    return rentals_ws.get_all_records()


def get_bike_rentals(bike_id):
    rows = rentals_ws.get_all_records()
    return [r for r in rows if str(r.get("bike_id")) == str(bike_id)]


def set_payment_status(rental_id, status):
    _, row_idx = get_rental_by_id(rental_id)
    if row_idx:
        col = rentals_ws.find("payment_status").col
        rentals_ws.update_cell(row_idx, col, status)
        return True
    return False


def set_return_status(rental_id, status):
    _, row_idx = get_rental_by_id(rental_id)
    if row_idx:
        col = rentals_ws.find("return_status").col
        rentals_ws.update_cell(row_idx, col, status)
        return True
    return False


def mark_paid(rental_id):
    return set_payment_status(rental_id, "оплачено")


def mark_pending(rental_id):
    return set_payment_status(rental_id, "ожидает подтверждения")


def mark_returned(rental_id):
    rental, row_idx = get_rental_by_id(rental_id)
    if row_idx:
        col = rentals_ws.find("return_status").col
        rentals_ws.update_cell(row_idx, col, "возвращён")
        set_bike_status(rental.get("bike_id"), "свободен")
        return True
    return False


def cancel_rental(rental_id):
    rental, row_idx = get_rental_by_id(rental_id)
    if row_idx and rental.get("return_status") == "арендован":
        col = rentals_ws.find("return_status").col
        rentals_ws.update_cell(row_idx, col, "отменена")
        set_bike_status(rental["bike_id"], "свободен")
        return True
    return False


# ------------------ ФИНАНСЫ ------------------

def add_finance_row(ftype, category, amount, owner, comment, bike_id=""):
    rows = finances_ws.get_all_records()
    new_id = len(rows) + 1
    finances_ws.append_row([
        new_id, datetime.now().strftime("%d.%m.%Y"), ftype,
        category, amount, owner, comment, bike_id
    ])
    return new_id


def add_income(rental_id, amount, owner, bike_id):
    return add_finance_row("доход", "аренда", amount, owner, f"аренда #{rental_id}", bike_id)


def add_expense(category, amount, owner, comment, bike_id=""):
    return add_finance_row("расход", category, amount, owner, comment, bike_id)


def get_finance_rows():
    return finances_ws.get_all_records()


def get_bike_finances(bike_id):
    rows = finances_ws.get_all_records()
    return [r for r in rows if str(r.get("bike_id")) == str(bike_id)]


def log_contract(rental_id, client_id, file_id):
    rows = contracts_ws.get_all_records()
    new_id = len(rows) + 1
    contracts_ws.append_row([new_id, rental_id, client_id, datetime.now().strftime("%d.%m.%Y"), file_id])


def get_contract_by_rental(rental_id):
    rows = contracts_ws.get_all_records()
    for r in rows:
        if str(r.get("rental_id")) == str(rental_id):
            return r
    return None


def clear_sheet_keep_header(ws):
    headers = ws.row_values(1)
    ws.clear()
    ws.append_row(headers)


def reset_test_data():
    clear_sheet_keep_header(rentals_ws)
    clear_sheet_keep_header(finances_ws)
    clear_sheet_keep_header(waiting_ws)
    bikes = bikes_ws.get_all_records()
    status_col = bikes_ws.find("status").col
    for i in range(len(bikes)):
        bikes_ws.update_cell(i + 2, status_col, "свободен")


# ------------------ ДОГОВОР ------------------

# (текстовый договор заменён на PDF — см. contract.py)


# ------------------ QR-КОД ------------------

def generate_qr_image(bike_id):
    link = f"https://t.me/{BOT_USERNAME}?start=bike_{bike_id}"
    img = qrcode.make(link)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf, link


# ------------------ КЛАВИАТУРЫ ------------------

def booking_buttons(rental_id):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Я оплатил(а)", callback_data=f"claim_paid_{rental_id}")],
        [InlineKeyboardButton(text="❌ Отменить бронь", callback_data=f"cancel_rental_{rental_id}")],
    ])


def paid_button(rental_id):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Я оплатил(а)", callback_data=f"claim_paid_{rental_id}")]
    ])


def rental_action_buttons(rental_id, unpaid):
    buttons = []
    if unpaid:
        buttons.append([InlineKeyboardButton(text="✅ Я оплатил(а)", callback_data=f"claim_paid_{rental_id}")])
        buttons.append([InlineKeyboardButton(text="❌ Отменить бронь", callback_data=f"cancel_rental_{rental_id}")])
    buttons.append([InlineKeyboardButton(text="🔄 Продлить", callback_data=f"extend_{rental_id}")])
    buttons.append([InlineKeyboardButton(text="🚲 Вернуть велосипед", callback_data=f"claim_return_{rental_id}")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def waitlist_button():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔔 Уведомить когда освободится", callback_data="waitlist_join")]
    ])


async def download_photo_b64(bot_instance, file_id):
    file = await bot_instance.get_file(file_id)
    buf = await bot_instance.download_file(file.file_path)
    return base64.b64encode(buf.read()).decode()


PASSPORT_PROMPT = (
    "Ты помощник по распознаванию паспорта гражданина РФ. Первое фото — главная "
    "страница (разворот с фото и личными данными), второе — страница с отметкой "
    "о регистрации (прописка). Извлеки данные и верни СТРОГО валидный JSON без "
    "пояснений и без markdown-разметки, в формате:\n"
    '{"full_name": "Фамилия Имя Отчество", "dob": "ДД.ММ.ГГГГ", '
    '"passport_series": "0000", "passport_number": "000000", '
    '"issued_by": "текст кем выдан", "issue_date": "ДД.ММ.ГГГГ", '
    '"department_code": "000-000", "registration_address": "полный адрес регистрации"}\n'
    "Если что-то не удалось разобрать, оставь пустую строку в этом поле."
)


async def recognize_passport(bot_instance, photo_main_id, photo_reg_id):
    if not openai_client:
        return None
    try:
        img1 = await download_photo_b64(bot_instance, photo_main_id)
        img2 = await download_photo_b64(bot_instance, photo_reg_id)
        response = await openai_client.chat.completions.create(
            model="gpt-4o",
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": PASSPORT_PROMPT},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img1}"}},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img2}"}},
                ]
            }],
            response_format={"type": "json_object"},
            max_tokens=500,
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        log.warning(f"Ошибка распознавания паспорта: {e}")
        return None


def period_keyboard(bike, prefix="period_"):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"Неделя — {bike['price_week']}₽", callback_data=f"{prefix}неделя")],
        [InlineKeyboardButton(text=f"Месяц — {bike['price_month']}₽", callback_data=f"{prefix}месяц")],
    ])


def main_menu_kb():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🚲 Арендовать велосипед")],
            [KeyboardButton(text="📋 Мои аренды"), KeyboardButton(text="ℹ️ Помощь")],
        ],
        resize_keyboard=True
    )


def admin_menu_kb():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📋 Активные аренды"), KeyboardButton(text="📊 Статистика")],
            [KeyboardButton(text="💰 Отчёт"), KeyboardButton(text="🧹 Сброс теста")],
            [KeyboardButton(text="❓ Все команды")],
        ],
        resize_keyboard=True
    )


# ------------------ БОТ ------------------

storage = MemoryStorage()
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=storage)
router = Router()
dp.include_router(router)


class Registration(StatesGroup):
    waiting_name = State()
    waiting_phone = State()
    waiting_photo_main = State()
    waiting_photo_reg = State()
    confirming_recognition = State()
    waiting_dob = State()
    waiting_passport = State()
    waiting_issued_by = State()
    waiting_issue_date = State()
    waiting_department_code = State()
    waiting_reg_address = State()
    waiting_actual_address_choice = State()
    waiting_actual_address_text = State()


class Rent(StatesGroup):
    choosing_bike = State()
    choosing_period = State()
    choosing_delivery = State()
    entering_delivery_address = State()


class Extend(StatesGroup):
    choosing_period = State()


class SetPhoto(StatesGroup):
    waiting_photo = State()


class AddClient(StatesGroup):
    waiting_name = State()
    waiting_phone = State()
    waiting_photo_main = State()
    waiting_photo_reg = State()
    confirming_recognition = State()
    waiting_dob = State()
    waiting_passport = State()
    waiting_issued_by = State()
    waiting_issue_date = State()
    waiting_department_code = State()
    waiting_reg_address = State()
    waiting_actual_address_choice = State()
    waiting_actual_address_text = State()


class ManualRent(StatesGroup):
    choosing_client = State()
    choosing_bike = State()
    choosing_period = State()


class ReturnFlow(StatesGroup):
    waiting_photo = State()


def admin_only(user_id):
    return user_id == ADMIN_ID


def get_role(user_id):
    if user_id == ADMIN_ID:
        return "admin"
    if FRIEND_ID and user_id == FRIEND_ID:
        return "friend"
    return None


@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    parts = message.text.split(maxsplit=1)
    payload = parts[1] if len(parts) > 1 else None
    bike_id = payload.split("_", 1)[1] if payload and payload.startswith("bike_") else None

    client = get_client_by_tg(message.from_user.id)
    if client:
        if bike_id:
            await start_bike_flow(message, state, bike_id)
        elif message.from_user.id == ADMIN_ID:
            await message.answer("Привет! Админ-панель ниже 👇", reply_markup=admin_menu_kb())
        else:
            await message.answer(
                "С возвращением! 🙂 Выберите действие в меню внизу.",
                reply_markup=main_menu_kb()
            )
        return

    if bike_id:
        await state.update_data(pending_bike_id=bike_id)
    await message.answer(
        "Привет! Это бот проката электровелосипедов.\n\n"
        "Давай зарегистрируемся. Напиши своё ФИО полностью, как в паспорте:"
    )
    await state.set_state(Registration.waiting_name)


@router.message(Registration.waiting_name)
async def reg_name(message: Message, state: FSMContext):
    await state.update_data(name=message.text)
    await message.answer("Отлично! Теперь отправь номер телефона (просто текстом).")
    await state.set_state(Registration.waiting_phone)


@router.message(Registration.waiting_phone)
async def reg_phone(message: Message, state: FSMContext):
    await state.update_data(phone=message.text)
    await message.answer("Пришлите фото главной страницы паспорта 📸")
    await state.set_state(Registration.waiting_photo_main)


@router.message(Registration.waiting_photo_main, F.photo)
async def reg_photo_main(message: Message, state: FSMContext):
    await state.update_data(photo_main=message.photo[-1].file_id)
    await message.answer("Теперь пришлите фото страницы с пропиской 📸")
    await state.set_state(Registration.waiting_photo_reg)


@router.message(Registration.waiting_photo_reg, F.photo)
async def reg_photo_reg(message: Message, state: FSMContext):
    data = await state.get_data()
    photo_reg_id = message.photo[-1].file_id
    await state.update_data(photo_reg=photo_reg_id)

    recognized = None
    if openai_client:
        wait_msg = await message.answer("Распознаю паспорт, секунду... ⏳")
        recognized = await recognize_passport(message.bot, data["photo_main"], photo_reg_id)
        try:
            await wait_msg.delete()
        except Exception:
            pass

    if recognized:
        await state.update_data(recognized=recognized)
        text = (
            "Вот что удалось распознать:\n\n"
            f"ФИО: {recognized.get('full_name') or '—'}\n"
            f"Дата рождения: {recognized.get('dob') or '—'}\n"
            f"Паспорт: {recognized.get('passport_series') or '—'} {recognized.get('passport_number') or ''}\n"
            f"Кем выдан: {recognized.get('issued_by') or '—'}\n"
            f"Дата выдачи: {recognized.get('issue_date') or '—'}\n"
            f"Код подразделения: {recognized.get('department_code') or '—'}\n"
            f"Адрес регистрации: {recognized.get('registration_address') or '—'}\n\n"
            "Всё верно?"
        )
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Всё верно", callback_data="recog_ok")],
            [InlineKeyboardButton(text="✏️ Заполнить вручную", callback_data="recog_manual")],
        ])
        await message.answer(text, reply_markup=kb)
        await state.set_state(Registration.confirming_recognition)
    else:
        if openai_client:
            await message.answer("Не удалось распознать автоматически, заполним вручную.")
        await message.answer("Укажите дату рождения (например: 15.03.1990):")
        await state.set_state(Registration.waiting_dob)


@router.callback_query(F.data == "recog_ok", Registration.confirming_recognition)
async def recog_confirm(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    rec = data.get("recognized", {})
    await state.update_data(
        name=rec.get("full_name") or data.get("name"),
        dob=rec.get("dob", ""),
        passport_series=rec.get("passport_series", ""),
        passport_number=rec.get("passport_number", ""),
        issued_by=rec.get("issued_by", ""),
        issue_date=rec.get("issue_date", ""),
        department_code=rec.get("department_code", ""),
        registration_address=rec.get("registration_address", "")
    )
    await callback.message.edit_text("Принято ✅")
    await ask_actual_address(callback.message, state)


@router.callback_query(F.data == "recog_manual", Registration.confirming_recognition)
async def recog_manual(callback: CallbackQuery, state: FSMContext):
    await callback.message.edit_text("Хорошо, заполним вручную.")
    await callback.message.answer("Укажите дату рождения (например: 15.03.1990):")
    await state.set_state(Registration.waiting_dob)


@router.message(Registration.waiting_dob)
async def reg_dob(message: Message, state: FSMContext):
    await state.update_data(dob=message.text)
    await message.answer(
        "Укажите серию и номер паспорта одним сообщением (например: 1234 567890):"
    )
    await state.set_state(Registration.waiting_passport)


@router.message(Registration.waiting_passport)
async def reg_passport(message: Message, state: FSMContext):
    parts = message.text.split(maxsplit=1)
    series = parts[0] if parts else message.text
    number = parts[1] if len(parts) > 1 else ""
    await state.update_data(passport_series=series, passport_number=number)
    await message.answer("Кем выдан паспорт?")
    await state.set_state(Registration.waiting_issued_by)


@router.message(Registration.waiting_issued_by)
async def reg_issued_by(message: Message, state: FSMContext):
    await state.update_data(issued_by=message.text)
    await message.answer("Дата выдачи паспорта (например: 20.05.2015):")
    await state.set_state(Registration.waiting_issue_date)


@router.message(Registration.waiting_issue_date)
async def reg_issue_date(message: Message, state: FSMContext):
    await state.update_data(issue_date=message.text)
    await message.answer("Код подразделения (например: 340-001):")
    await state.set_state(Registration.waiting_department_code)


@router.message(Registration.waiting_department_code)
async def reg_department_code(message: Message, state: FSMContext):
    await state.update_data(department_code=message.text)
    await message.answer("Укажите адрес регистрации (прописку) как в паспорте:")
    await state.set_state(Registration.waiting_reg_address)


@router.message(Registration.waiting_reg_address)
async def reg_address(message: Message, state: FSMContext):
    await state.update_data(registration_address=message.text)
    await ask_actual_address(message, state)


async def ask_actual_address(target, state: FSMContext):
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Да, такой же", callback_data="actual_same")],
        [InlineKeyboardButton(text="✏️ Другой адрес", callback_data="actual_diff")],
    ])
    await target.answer("Фактический адрес проживания такой же, как прописка?", reply_markup=kb)
    await state.set_state(Registration.waiting_actual_address_choice)


@router.callback_query(F.data == "actual_same", Registration.waiting_actual_address_choice)
async def reg_actual_same(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    await state.update_data(actual_address=data.get("registration_address", ""))
    await callback.message.edit_text("Принято ✅")
    await finalize_registration(callback.message, callback.from_user.id, state)


@router.callback_query(F.data == "actual_diff", Registration.waiting_actual_address_choice)
async def reg_actual_diff(callback: CallbackQuery, state: FSMContext):
    await callback.message.edit_text("Введите фактический адрес проживания:")
    await state.set_state(Registration.waiting_actual_address_text)


@router.message(Registration.waiting_actual_address_text)
async def reg_actual_address_text(message: Message, state: FSMContext):
    await state.update_data(actual_address=message.text)
    await finalize_registration(message, message.from_user.id, state)


async def finalize_registration(target: Message, user_id: int, state: FSMContext):
    data = await state.get_data()
    get_or_create_client(
        user_id, data.get("name"), data.get("phone"),
        data.get("passport_series", ""), data.get("passport_number", ""),
        data.get("registration_address", ""), data.get("photo_main", ""),
        data.get("photo_reg", ""),
        data.get("dob", ""), data.get("issued_by", ""), data.get("issue_date", ""),
        data.get("department_code", ""), data.get("actual_address", "")
    )
    if data.get("photo_main"):
        await bot.send_photo(
            ADMIN_ID, data["photo_main"],
            caption=f"🪪 Паспорт нового клиента: {data.get('name', '—')}, тел. {data.get('phone', '—')}"
        )
    if data.get("photo_reg"):
        await bot.send_photo(ADMIN_ID, data["photo_reg"], caption="🪪 Страница с пропиской")
    bike_id = data.get("pending_bike_id")
    await state.clear()
    if bike_id:
        await start_bike_flow(target, state, bike_id)
    else:
        await target.answer(
            "Регистрация завершена! ✅ Выберите действие в меню внизу.",
            reply_markup=main_menu_kb()
        )


async def start_bike_flow(message: Message, state: FSMContext, bike_id):
    bike, _ = get_bike_by_id(bike_id)
    if not bike or bike.get("status") != "свободен":
        await message.answer("Этот велосипед сейчас недоступен, вот что есть свободного:")
        await cmd_rent(message, state)
        return
    await state.update_data(bike_id=bike_id)
    if bike.get("photo_id"):
        await message.answer_photo(bike["photo_id"], caption=bike["name_model"])
    await message.answer(f"Велосипед: {bike['name_model']}\nНа какой срок?", reply_markup=period_keyboard(bike))
    await state.set_state(Rent.choosing_period)


@router.message(F.text == "🚲 Арендовать велосипед")
async def menu_rent(message: Message, state: FSMContext):
    await cmd_rent(message, state)


@router.message(F.text == "📋 Мои аренды")
async def menu_my(message: Message):
    await cmd_my(message)


@router.message(F.text == "ℹ️ Помощь")
async def menu_help(message: Message):
    await message.answer(
        "🚲 Как это работает:\n\n"
        "1. Нажмите «Арендовать велосипед», выберите модель и срок\n"
        "2. Заберите велосипед сами или закажите доставку (+500₽)\n"
        "3. Переведите оплату и нажмите «Я оплатил(а)»\n"
        "4. Когда закончите — нажмите «Вернуть велосипед» в разделе «Мои аренды»\n\n"
        f"По всем вопросам: {BUSINESS_PHONE}"
    )


@router.message(Command("admin"))
async def cmd_admin_menu(message: Message):
    if not admin_only(message.from_user.id):
        return
    await message.answer("Админ-панель:", reply_markup=admin_menu_kb())


@router.message(F.text == "📋 Активные аренды")
async def menu_admin_rentals(message: Message):
    await cmd_rentals(message)


@router.message(F.text == "📊 Статистика")
async def menu_admin_stats(message: Message):
    await cmd_stats(message)


@router.message(F.text == "💰 Отчёт")
async def menu_admin_report(message: Message):
    await cmd_report(message)


@router.message(F.text == "🧹 Сброс теста")
async def menu_admin_reset(message: Message):
    await cmd_reset_test(message)


@router.message(F.text == "❓ Все команды")
async def menu_admin_help(message: Message):
    if not admin_only(message.from_user.id):
        return
    await message.answer(
        "📖 Все команды администратора\n\n"
        "— Без параметров (есть кнопки) —\n"
        "/rentals — список активных аренд\n"
        "/stats — статистика (велики, клиенты, доход за сегодня)\n"
        "/report — финансовый отчёт (доход/расход, я и Денц)\n"
        "/reset_test — очистить тестовые данные (с подтверждением)\n\n"
        "— С параметрами (наберите вручную) —\n"
        "/paid ID_аренды — вручную отметить аренду оплаченной\n"
        "  пример: /paid 5\n\n"
        "/return ID_аренды — вручную закрыть аренду, освободить велик\n"
        "  пример: /return 5\n\n"
        "/expense Владелец ID_велика Сумма Категория Комментарий\n"
        "  Владелец: Я или Денц. ID_велика: число, или - если расход общий\n"
        "  пример: /expense Я 3 1500 ремонт замена камеры\n\n"
        "/bike_history ID_велика — история аренд и расходов по велику\n"
        "  пример: /bike_history 3\n\n"
        "/blacklist add ID_клиента — добавить в чёрный список\n"
        "/blacklist remove ID_клиента — убрать из чёрного списка\n"
        "  пример: /blacklist add 7\n\n"
        "/broadcast текст — разослать сообщение всем клиентам\n"
        "  пример: /broadcast Завтра не работаем, извините за неудобства\n\n"
        "/set_photo ID_велика — затем прислать фото следующим сообщением\n"
        "  пример: /set_photo 3\n\n"
        "/qr ID_велика — получить QR-код для наклейки на велосипед\n"
        "  пример: /qr 3\n\n"
        "/contract ID_аренды — заново получить PDF-договор по аренде\n"
        "  пример: /contract 12\n\n"
        "/add_client — добавить клиента без Telegram (пошагово, с фото паспорта)\n\n"
        "/manual_rent — оформить аренду для клиента без Telegram (выбор кнопками)\n\n"
        "— Автоматически, без команд —\n"
        "🔔 Уведомления о новых арендах, заявках на оплату и возврат приходят сами\n"
        "🔧 Напоминания о ТО — раз в неделю, если велик давно не обслуживался\n"
        "🔴 Напоминания о просрочках — приходят автоматически вам и клиенту"
    )


@router.message(Command("rent"))
async def cmd_rent(message: Message, state: FSMContext):
    client = get_client_by_tg(message.from_user.id)
    if client and is_blacklisted(client["id"]):
        await bot.send_message(
            ADMIN_ID,
            f"⚠️ Клиент из чёрного списка пытается арендовать велосипед!\n"
            f"Имя: {client['name']}, телефон: {client['phone']}, id {client['id']}"
        )
    free_bikes = get_free_bikes()
    if not free_bikes:
        await message.answer("Сейчас все велосипеды заняты, попробуйте позже 🙁", reply_markup=waitlist_button())
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=b["name_model"], callback_data=f"bike_{b['id']}")]
        for b in free_bikes
    ])
    await message.answer("Выберите велосипед:", reply_markup=kb)
    await state.set_state(Rent.choosing_bike)


@router.callback_query(F.data == "waitlist_join")
async def waitlist_join(callback: CallbackQuery):
    get_or_create_client(callback.from_user.id, callback.from_user.full_name, "")
    added = add_to_waitlist(callback.from_user.id, callback.from_user.full_name)
    if added:
        await callback.answer("Записали! Уведомим как только освободится велосипед 🔔", show_alert=True)
    else:
        await callback.answer("Вы уже в списке ожидания", show_alert=True)


@router.callback_query(F.data.startswith("bike_"))
async def choose_bike(callback: CallbackQuery, state: FSMContext):
    bike_id = callback.data.split("_")[1]
    bike, _ = get_bike_by_id(bike_id)
    await state.update_data(bike_id=bike_id)
    if bike.get("photo_id"):
        await callback.message.answer_photo(bike["photo_id"], caption=bike["name_model"])
    await callback.message.edit_text(f"Выбран: {bike['name_model']}\nНа какой срок?", reply_markup=period_keyboard(bike))
    await state.set_state(Rent.choosing_period)


@router.callback_query(F.data.startswith("period_"), Rent.choosing_period)
async def choose_period(callback: CallbackQuery, state: FSMContext):
    period = callback.data.split("_")[1]
    await state.update_data(period=period)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🚶 Заберу сам(а), бесплатно", callback_data="pickup_self")],
        [InlineKeyboardButton(text=f"🚚 Доставка (+{DELIVERY_PRICE}₽)", callback_data="pickup_delivery")],
    ])
    await callback.message.edit_text(f"Срок: {period}\nКак заберёте велосипед?", reply_markup=kb)
    await state.set_state(Rent.choosing_delivery)


@router.callback_query(F.data == "pickup_self", Rent.choosing_delivery)
async def pickup_self(callback: CallbackQuery, state: FSMContext):
    await finalize_rental(callback, state, delivery="нет", delivery_address="")


@router.callback_query(F.data == "pickup_delivery", Rent.choosing_delivery)
async def pickup_delivery(callback: CallbackQuery, state: FSMContext):
    await callback.message.edit_text("Напишите адрес доставки (улица, дом):")
    await state.set_state(Rent.entering_delivery_address)


@router.message(Rent.entering_delivery_address)
async def delivery_address_entered(message: Message, state: FSMContext):
    await finalize_rental(message, state, delivery="да", delivery_address=message.text)


async def finalize_rental(event, state: FSMContext, delivery, delivery_address):
    data = await state.get_data()
    bike_id = data["bike_id"]
    period = data["period"]
    bike, _ = get_bike_by_id(bike_id)
    price = int(bike["price_week"] if period == "неделя" else bike["price_month"])
    if delivery == "да":
        price += DELIVERY_PRICE

    user = event.from_user
    client_id = get_or_create_client(user.id, user.full_name, "")
    client = get_client_by_tg(user.id)
    rental_id, start_dt, end_dt = create_rental(bike_id, client_id, period, price, bike["owner"], delivery, delivery_address)

    text = (
        f"Готово! ✅\n\n"
        f"Велосипед: {bike['name_model']}\n"
        f"Срок: {period}\n"
        f"Сумма: {price}₽" + (f" (включая доставку {DELIVERY_PRICE}₽)" if delivery == "да" else "") + "\n"
        f"Вернуть/продлить до: {end_dt.strftime('%d.%m.%Y')}\n\n"
    )
    if delivery == "да":
        text += f"🚚 Доставим по адресу: {delivery_address}\nТелефон для связи: {BUSINESS_PHONE}\n\n"
    else:
        text += f"📍 Забрать по адресу: {BUSINESS_ADDRESS}\nТелефон для связи: {BUSINESS_PHONE}\n\n"
    text += "Когда оплатите — нажмите кнопку ниже 👇"

    if isinstance(event, CallbackQuery):
        await event.message.edit_text(text)
        chat = event.message.chat.id
    else:
        await event.answer(text)
        chat = event.chat.id

    await bot.send_message(chat, PAYMENT_INFO)
    await bot.send_message(chat, "Управление бронью:", reply_markup=booking_buttons(rental_id))

    contract_data = {
        "rental_id": rental_id,
        "contract_date": ru_date(date_cls.today()),
        "business_phone": BUSINESS_PHONE,
        "full_name": client["name"] if client else user.full_name,
        "dob": client.get("dob", "") if client else "",
        "passport_series": client.get("passport_series", "") if client else "",
        "passport_number": client.get("passport_number", "") if client else "",
        "issued_by": client.get("passport_issued_by", "") if client else "",
        "issue_date": client.get("passport_issue_date", "") if client else "",
        "department_code": client.get("department_code", "") if client else "",
        "registration_address": client.get("registration_address", "") if client else "",
        "actual_address": client.get("actual_address", "") if client else "",
        "phone": client["phone"] if client else "",
        "bike_name": bike["name_model"],
        "serial": extract_serial(bike["name_model"]),
        "start_date": start_dt.strftime("%d.%m.%Y"),
        "end_date": end_dt.strftime("%d.%m.%Y"),
        "price_week": bike["price_week"],
        "price_month": bike["price_month"],
        "daily_penalty": round(int(bike["price_week"]) / 7),
    }
    pdf_buf = generate_contract_pdf(contract_data)
    pdf_bytes = pdf_buf.read()
    sent_doc = await bot.send_document(
        chat,
        BufferedInputFile(pdf_bytes, filename=f"dogovor_{rental_id}.pdf"),
        caption="📄 Договор аренды — ознакомьтесь перед встречей, подпишем при получении велосипеда"
    )
    log_contract(rental_id, client_id, sent_doc.document.file_id)
    await bot.send_document(
        ADMIN_ID,
        BufferedInputFile(pdf_bytes, filename=f"dogovor_{rental_id}.pdf"),
        caption=f"📄 Копия договора #{rental_id} — можно распечатать"
    )

    await bot.send_message(
        ADMIN_ID,
        f"🆕 Новая аренда #{rental_id}\n"
        f"{bike['name_model']} — {period} ({price}₽)\n"
        f"Клиент: {user.full_name} (id {client_id})\n"
        f"Владелец велика: {bike['owner']}\n"
        + (f"Доставка: {delivery_address}" if delivery == "да" else "Самовывоз")
    )
    await state.clear()


# ------------------ ОТМЕНА БРОНИ ------------------

@router.callback_query(F.data.startswith("cancel_rental_"))
async def cancel_rental_handler(callback: CallbackQuery):
    rental_id = callback.data.split("_")[-1]
    rental, _ = get_rental_by_id(rental_id)
    if not rental or rental.get("return_status") != "арендован":
        await callback.answer("Бронь уже не активна", show_alert=True)
        return
    if rental.get("payment_status") == "оплачено":
        await callback.answer("Аренда уже оплачена, для отмены свяжитесь с нами", show_alert=True)
        return
    cancel_rental(rental_id)
    await callback.message.edit_text(callback.message.text + "\n\n❌ Бронь отменена")
    await bot.send_message(ADMIN_ID, f"❌ Клиент отменил бронь #{rental_id} (велосипед снова свободен)")
    await notify_and_clear_waitlist()


# ------------------ ОПЛАТА ------------------

@router.callback_query(F.data.startswith("claim_paid_"))
async def claim_paid(callback: CallbackQuery):
    rental_id = callback.data.split("_")[-1]
    rental, _ = get_rental_by_id(rental_id)
    if not rental:
        await callback.answer("Аренда не найдена", show_alert=True)
        return
    if rental.get("payment_status") == "оплачено":
        await callback.answer("Уже подтверждено ✅", show_alert=True)
        return
    mark_pending(rental_id)
    await callback.message.edit_text(callback.message.text + "\n\n⏳ Ждём подтверждения от администратора...")
    admin_kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Подтвердить", callback_data=f"confirm_paid_{rental_id}"),
        InlineKeyboardButton(text="❌ Отклонить", callback_data=f"reject_paid_{rental_id}"),
    ]])
    await bot.send_message(
        ADMIN_ID,
        f"💰 Клиент отметил оплату по аренде #{rental_id}\n"
        f"Сумма: {rental['amount']}₽ | Владелец: {rental['owner']}\n"
        f"Проверьте поступление и подтвердите:",
        reply_markup=admin_kb
    )


@router.callback_query(F.data.startswith("confirm_paid_"))
async def confirm_paid(callback: CallbackQuery):
    if not admin_only(callback.from_user.id):
        await callback.answer("Только для администратора", show_alert=True)
        return
    rental_id = callback.data.split("_")[-1]
    rental, _ = get_rental_by_id(rental_id)
    mark_paid(rental_id)
    add_income(rental_id, rental["amount"], rental["owner"], rental["bike_id"])
    await callback.message.edit_text(callback.message.text + "\n\n✅ Оплата подтверждена")
    tg_id = get_client_telegram_id(rental["client_id"])
    if tg_id and not is_manual_tg_id(tg_id):
        await bot.send_message(tg_id, f"Оплата по аренде #{rental_id} подтверждена ✅ Спасибо!")


@router.callback_query(F.data.startswith("reject_paid_"))
async def reject_paid(callback: CallbackQuery):
    if not admin_only(callback.from_user.id):
        await callback.answer("Только для администратора", show_alert=True)
        return
    rental_id = callback.data.split("_")[-1]
    rental, _ = get_rental_by_id(rental_id)
    set_payment_status(rental_id, "не оплачено")
    await callback.message.edit_text(callback.message.text + "\n\n❌ Отклонено")
    tg_id = get_client_telegram_id(rental["client_id"])
    if tg_id and not is_manual_tg_id(tg_id):
        await bot.send_message(tg_id, f"По аренде #{rental_id} мы не нашли поступление оплаты. Свяжитесь с нами.")


# ------------------ ПРОДЛЕНИЕ ------------------

@router.callback_query(F.data.startswith("extend_"))
async def extend_start(callback: CallbackQuery, state: FSMContext):
    rental_id = callback.data.split("_")[-1]
    rental, _ = get_rental_by_id(rental_id)
    bike, _ = get_bike_by_id(rental["bike_id"])
    await state.update_data(old_rental_id=rental_id, bike_id=rental["bike_id"])
    await callback.message.answer("На какой срок продлить?", reply_markup=period_keyboard(bike, prefix="ext_period_"))
    await state.set_state(Extend.choosing_period)


@router.callback_query(F.data.startswith("ext_period_"), Extend.choosing_period)
async def extend_confirm(callback: CallbackQuery, state: FSMContext):
    period = callback.data.split("_")[-1]
    data = await state.get_data()
    old_rental_id = data["old_rental_id"]
    bike_id = data["bike_id"]
    bike, _ = get_bike_by_id(bike_id)
    price = int(bike["price_week"] if period == "неделя" else bike["price_month"])

    set_return_status(old_rental_id, "продлена")
    client_id = get_or_create_client(callback.from_user.id, callback.from_user.full_name, "")
    new_id, start_dt, end_dt = create_rental(bike_id, client_id, period, price, bike["owner"])

    await callback.message.edit_text(
        f"Продлено ✅\n\n"
        f"Новая аренда #{new_id}, срок: {period}, сумма: {price}₽\n"
        f"До: {end_dt.strftime('%d.%m.%Y')}\n\n"
        f"Когда оплатите — нажмите кнопку ниже 👇"
    )
    await callback.message.answer(f"Оплата за продление:\n{PAYMENT_INFO}", reply_markup=paid_button(new_id))
    await bot.send_message(
        ADMIN_ID,
        f"🔄 Продление: старая аренда #{old_rental_id} закрыта, новая #{new_id} "
        f"({bike['name_model']}, {period}, {price}₽)"
    )
    await state.clear()


# ------------------ ВОЗВРАТ (с фотофиксацией) ------------------

@router.callback_query(F.data.startswith("claim_return_"))
async def claim_return(callback: CallbackQuery, state: FSMContext):
    rental_id = callback.data.split("_")[-1]
    rental, _ = get_rental_by_id(rental_id)
    if not rental or rental.get("return_status") != "арендован":
        await callback.answer("Аренда уже закрыта", show_alert=True)
        return
    await state.update_data(return_rental_id=rental_id)
    await callback.message.answer("Пришлите, пожалуйста, фото велосипеда при возврате 📸")
    await state.set_state(ReturnFlow.waiting_photo)
    await callback.answer()


@router.message(ReturnFlow.waiting_photo, F.photo)
async def return_photo_received(message: Message, state: FSMContext):
    data = await state.get_data()
    rental_id = data["return_rental_id"]
    file_id = message.photo[-1].file_id
    admin_kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Подтвердить возврат", callback_data=f"confirm_return_{rental_id}"),
    ]])
    await bot.send_photo(
        ADMIN_ID, file_id,
        caption=f"🚲 Клиент хочет вернуть велосипед по аренде #{rental_id}\nПроверьте состояние и подтвердите:",
        reply_markup=admin_kb
    )
    await message.answer("Фото получено, ждите подтверждения администратора ✅")
    await state.clear()


@router.callback_query(F.data.startswith("confirm_return_"))
async def confirm_return(callback: CallbackQuery):
    if not admin_only(callback.from_user.id):
        await callback.answer("Только для администратора", show_alert=True)
        return
    rental_id = callback.data.split("_")[-1]
    rental, _ = get_rental_by_id(rental_id)
    mark_returned(rental_id)
    await callback.message.edit_caption(caption=(callback.message.caption or "") + "\n\n✅ Возврат подтверждён")
    tg_id = get_client_telegram_id(rental["client_id"])
    if tg_id and not is_manual_tg_id(tg_id):
        await bot.send_message(tg_id, f"Возврат велосипеда по аренде #{rental_id} подтверждён, спасибо! 🚲")
    await notify_and_clear_waitlist()


# ------------------ МОИ АРЕНДЫ ------------------

@router.message(Command("my"))
async def cmd_my(message: Message):
    rows = get_active_rentals()
    client = get_client_by_tg(message.from_user.id)
    if not client:
        await message.answer("Вы ещё не регистрировались — введите /start")
        return
    my_rentals = [
        r for r in rows
        if str(r.get("client_id")) == str(client["id"]) and r.get("return_status") == "арендован"
    ]
    if not my_rentals:
        await message.answer("Активных аренд нет.")
        return
    for r in my_rentals:
        text = (
            f"Аренда #{r['id']}\n"
            f"До: {r['end_date']}\n"
            f"Сумма: {r['amount']}₽\n"
            f"Оплата: {r['payment_status']}"
        )
        unpaid = r["payment_status"] != "оплачено"
        await message.answer(text, reply_markup=rental_action_buttons(r["id"], unpaid))


# ------------------ АДМИНСКИЕ КОМАНДЫ ------------------

@router.message(Command("rentals"))
async def cmd_rentals(message: Message):
    if not admin_only(message.from_user.id):
        return
    rows = get_active_rentals()
    active = [r for r in rows if r.get("return_status") == "арендован"]
    if not active:
        await message.answer("Активных аренд нет.")
        return
    text = "Активные аренды:\n\n"
    for r in active:
        text += f"#{r['id']} | {r['end_date']} | {r['amount']}₽ | {r['payment_status']} | владелец: {r['owner']}\n"
    await message.answer(text)


@router.message(Command("paid"))
async def cmd_paid(message: Message):
    if not admin_only(message.from_user.id):
        return
    parts = message.text.split()
    if len(parts) != 2:
        await message.answer("Использование: /paid ID_аренды")
        return
    rental, _ = get_rental_by_id(parts[1])
    if rental and mark_paid(parts[1]):
        add_income(parts[1], rental["amount"], rental["owner"], rental["bike_id"])
        await message.answer(f"Аренда #{parts[1]} отмечена как оплаченная ✅")
    else:
        await message.answer("Не найдено такой аренды.")


@router.message(Command("return"))
async def cmd_return(message: Message):
    if not admin_only(message.from_user.id):
        return
    parts = message.text.split()
    if len(parts) != 2:
        await message.answer("Использование: /return ID_аренды")
        return
    if mark_returned(parts[1]):
        await message.answer(f"Аренда #{parts[1]} закрыта, велосипед снова свободен ✅")
        await notify_and_clear_waitlist()
    else:
        await message.answer("Не найдено такой аренды.")


@router.message(Command("expense"))
async def cmd_expense(message: Message):
    if not admin_only(message.from_user.id):
        return
    parts = message.text.split(maxsplit=5)
    if len(parts) < 5:
        await message.answer(
            "Использование: /expense <владелец: Я/Денц> <ID велосипеда или -> <сумма> <категория> <комментарий>\n"
            "Пример: /expense Я 3 1500 ремонт замена камеры"
        )
        return
    _, owner, bike_id, amount, category = parts[:5]
    comment = parts[5] if len(parts) > 5 else ""
    if owner not in ("Я", "Денц"):
        await message.answer("Владелец должен быть 'Я' или 'Денц'")
        return
    try:
        amount = int(amount)
    except ValueError:
        await message.answer("Сумма должна быть числом")
        return
    bike_id_val = "" if bike_id == "-" else bike_id
    exp_id = add_expense(category, amount, owner, comment, bike_id_val)
    await message.answer(f"Расход #{exp_id} записан: {owner}, велик {bike_id}, {amount}₽, {category}")


@router.message(Command("report"))
async def cmd_report(message: Message):
    role = get_role(message.from_user.id)
    if role is None:
        return
    finances = get_finance_rows()
    income_me = sum(int(r["sum"]) for r in finances if r.get("type") == "доход" and r.get("owner") == "Я")
    income_friend = sum(int(r["sum"]) for r in finances if r.get("type") == "доход" and r.get("owner") == "Денц")
    expense_me = sum(int(r["sum"]) for r in finances if r.get("type") == "расход" and r.get("owner") == "Я")
    expense_friend = sum(int(r["sum"]) for r in finances if r.get("type") == "расход" and r.get("owner") == "Денц")

    if role == "admin":
        await message.answer(
            f"📊 Отчёт\n\n"
            f"— Доходы —\nМоя доля: {income_me}₽\nДоля Денца: {income_friend}₽\n\n"
            f"— Расходы —\nМои: {expense_me}₽\nДенца: {expense_friend}₽\n\n"
            f"— Итого чистыми —\nЯ: {income_me - expense_me}₽\nДенц: {income_friend - expense_friend}₽"
        )
    else:
        await message.answer(
            f"📊 Ваш отчёт\n\nДоход: {income_friend}₽\nРасходы: {expense_friend}₽\n"
            f"Чистыми: {income_friend - expense_friend}₽"
        )


@router.message(Command("stats"))
async def cmd_stats(message: Message):
    if not admin_only(message.from_user.id):
        return
    bikes = bikes_ws.get_all_records()
    total_bikes = len(bikes)
    free = len([b for b in bikes if b.get("status") == "свободен"])
    rented = len([b for b in bikes if b.get("status") == "в аренде"])
    maintenance = len([b for b in bikes if b.get("status") == "на обслуживании"])

    clients = get_all_clients()
    total_clients = len(clients)

    finances = get_finance_rows()
    today = datetime.now().strftime("%d.%m.%Y")
    today_income = sum(int(r["sum"]) for r in finances if r.get("type") == "доход" and r.get("date") == today)

    await message.answer(
        f"📈 Статистика\n\n"
        f"Велосипедов всего: {total_bikes}\n"
        f"Свободно: {free} | В аренде: {rented} | На ТО: {maintenance}\n\n"
        f"Клиентов: {total_clients}\n\n"
        f"Доход сегодня: {today_income}₽"
    )


@router.message(Command("bike_history"))
async def cmd_bike_history(message: Message):
    if not admin_only(message.from_user.id):
        return
    parts = message.text.split()
    if len(parts) != 2:
        await message.answer("Использование: /bike_history ID_велосипеда")
        return
    bike_id = parts[1]
    bike, _ = get_bike_by_id(bike_id)
    if not bike:
        await message.answer("Велосипед не найден")
        return
    rentals = get_bike_rentals(bike_id)
    finances = get_bike_finances(bike_id)
    income = sum(int(r["sum"]) for r in finances if r.get("type") == "доход")
    expense = sum(int(r["sum"]) for r in finances if r.get("type") == "расход")

    text = f"🚲 {bike['name_model']} (id {bike_id})\n\n"
    text += f"Всего аренд: {len(rentals)}\n"
    text += f"Доход: {income}₽ | Расходы (ремонт): {expense}₽ | Чистыми: {income - expense}₽\n\n"
    if expense > 0:
        text += "Расходы:\n"
        for f in finances:
            if f.get("type") == "расход":
                text += f"- {f['date']}: {f['sum']}₽ ({f['category']}) {f.get('comment', '')}\n"
    await message.answer(text)


@router.message(Command("blacklist"))
async def cmd_blacklist(message: Message):
    if not admin_only(message.from_user.id):
        return
    parts = message.text.split(maxsplit=2)
    if len(parts) != 3 or parts[1] not in ("add", "remove"):
        await message.answer("Использование: /blacklist add ID_клиента  или  /blacklist remove ID_клиента")
        return
    _, action, client_id = parts
    rows = clients_ws.get_all_records()
    for i, r in enumerate(rows, start=2):
        if str(r.get("id")) == str(client_id):
            col = clients_ws.find("blacklist").col
            clients_ws.update_cell(i, col, "да" if action == "add" else "нет")
            await message.answer(f"Клиент {client_id} {'добавлен в' if action == 'add' else 'убран из'} чёрный список")
            return
    await message.answer("Клиент не найден")


@router.message(Command("broadcast"))
async def cmd_broadcast(message: Message):
    if not admin_only(message.from_user.id):
        return
    text = message.text.replace("/broadcast", "", 1).strip()
    if not text:
        await message.answer("Использование: /broadcast текст сообщения")
        return
    clients = get_all_clients()
    sent = 0
    for r in clients:
        tg = r.get("telegram_id")
        if not tg:
            continue
        try:
            await bot.send_message(tg, f"📢 {text}")
            sent += 1
        except Exception as e:
            log.warning(f"Не удалось отправить {tg}: {e}")
    await message.answer(f"Разослано {sent} клиентам")


@router.message(Command("add_client"))
async def cmd_add_client(message: Message, state: FSMContext):
    if not admin_only(message.from_user.id):
        return
    await message.answer("Добавляем клиента без Telegram.\n\nФ.И.О. клиента:")
    await state.set_state(AddClient.waiting_name)


@router.message(AddClient.waiting_name)
async def ac_name(message: Message, state: FSMContext):
    await state.update_data(name=message.text)
    await message.answer("Номер телефона:")
    await state.set_state(AddClient.waiting_phone)


@router.message(AddClient.waiting_phone)
async def ac_phone(message: Message, state: FSMContext):
    await state.update_data(phone=message.text)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⏭ Пропустить фото", callback_data="ac_skip_photos")]
    ])
    await message.answer("Пришлите фото главной страницы паспорта 📸 (или пропустите)", reply_markup=kb)
    await state.set_state(AddClient.waiting_photo_main)


@router.callback_query(F.data == "ac_skip_photos", AddClient.waiting_photo_main)
async def ac_skip_photos(callback: CallbackQuery, state: FSMContext):
    await callback.message.edit_text("Фото пропущены.")
    await callback.message.answer("Укажите дату рождения (например: 15.03.1990), или «-» если нет данных:")
    await state.set_state(AddClient.waiting_dob)


@router.message(AddClient.waiting_photo_main, F.photo)
async def ac_photo_main(message: Message, state: FSMContext):
    await state.update_data(photo_main=message.photo[-1].file_id)
    await message.answer("Теперь фото страницы с пропиской 📸")
    await state.set_state(AddClient.waiting_photo_reg)


@router.message(AddClient.waiting_photo_reg, F.photo)
async def ac_photo_reg(message: Message, state: FSMContext):
    data = await state.get_data()
    photo_reg_id = message.photo[-1].file_id
    await state.update_data(photo_reg=photo_reg_id)

    recognized = None
    if openai_client:
        wait_msg = await message.answer("Распознаю паспорт, секунду... ⏳")
        recognized = await recognize_passport(message.bot, data["photo_main"], photo_reg_id)
        try:
            await wait_msg.delete()
        except Exception:
            pass

    if recognized:
        await state.update_data(recognized=recognized)
        text = (
            "Распознано:\n\n"
            f"ФИО: {recognized.get('full_name') or '—'}\n"
            f"Дата рождения: {recognized.get('dob') or '—'}\n"
            f"Паспорт: {recognized.get('passport_series') or '—'} {recognized.get('passport_number') or ''}\n"
            f"Кем выдан: {recognized.get('issued_by') or '—'}\n"
            f"Дата выдачи: {recognized.get('issue_date') or '—'}\n"
            f"Код подразделения: {recognized.get('department_code') or '—'}\n"
            f"Адрес регистрации: {recognized.get('registration_address') or '—'}\n\n"
            "Всё верно?"
        )
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Всё верно", callback_data="ac_recog_ok")],
            [InlineKeyboardButton(text="✏️ Заполнить вручную", callback_data="ac_recog_manual")],
        ])
        await message.answer(text, reply_markup=kb)
        await state.set_state(AddClient.confirming_recognition)
    else:
        await message.answer("Укажите дату рождения (например: 15.03.1990):")
        await state.set_state(AddClient.waiting_dob)


@router.callback_query(F.data == "ac_recog_ok", AddClient.confirming_recognition)
async def ac_recog_confirm(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    rec = data.get("recognized", {})
    await state.update_data(
        name=rec.get("full_name") or data.get("name"),
        dob=rec.get("dob", ""), passport_series=rec.get("passport_series", ""),
        passport_number=rec.get("passport_number", ""), issued_by=rec.get("issued_by", ""),
        issue_date=rec.get("issue_date", ""), department_code=rec.get("department_code", ""),
        registration_address=rec.get("registration_address", "")
    )
    await callback.message.edit_text("Принято ✅")
    await ac_ask_actual_address(callback.message, state)


@router.callback_query(F.data == "ac_recog_manual", AddClient.confirming_recognition)
async def ac_recog_manual(callback: CallbackQuery, state: FSMContext):
    await callback.message.edit_text("Заполним вручную.")
    await callback.message.answer("Укажите дату рождения (например: 15.03.1990):")
    await state.set_state(AddClient.waiting_dob)


@router.message(AddClient.waiting_dob)
async def ac_dob(message: Message, state: FSMContext):
    await state.update_data(dob=message.text)
    await message.answer("Серия и номер паспорта одним сообщением (например: 1234 567890):")
    await state.set_state(AddClient.waiting_passport)


@router.message(AddClient.waiting_passport)
async def ac_passport(message: Message, state: FSMContext):
    parts = message.text.split(maxsplit=1)
    series = parts[0] if parts else message.text
    number = parts[1] if len(parts) > 1 else ""
    await state.update_data(passport_series=series, passport_number=number)
    await message.answer("Кем выдан паспорт?")
    await state.set_state(AddClient.waiting_issued_by)


@router.message(AddClient.waiting_issued_by)
async def ac_issued_by(message: Message, state: FSMContext):
    await state.update_data(issued_by=message.text)
    await message.answer("Дата выдачи паспорта:")
    await state.set_state(AddClient.waiting_issue_date)


@router.message(AddClient.waiting_issue_date)
async def ac_issue_date(message: Message, state: FSMContext):
    await state.update_data(issue_date=message.text)
    await message.answer("Код подразделения:")
    await state.set_state(AddClient.waiting_department_code)


@router.message(AddClient.waiting_department_code)
async def ac_department_code(message: Message, state: FSMContext):
    await state.update_data(department_code=message.text)
    await message.answer("Адрес регистрации (прописка):")
    await state.set_state(AddClient.waiting_reg_address)


@router.message(AddClient.waiting_reg_address)
async def ac_reg_address(message: Message, state: FSMContext):
    await state.update_data(registration_address=message.text)
    await ac_ask_actual_address(message, state)


async def ac_ask_actual_address(target, state: FSMContext):
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Да, такой же", callback_data="ac_actual_same")],
        [InlineKeyboardButton(text="✏️ Другой адрес", callback_data="ac_actual_diff")],
    ])
    await target.answer("Фактический адрес такой же, как прописка?", reply_markup=kb)
    await state.set_state(AddClient.waiting_actual_address_choice)


@router.callback_query(F.data == "ac_actual_same", AddClient.waiting_actual_address_choice)
async def ac_actual_same(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    await state.update_data(actual_address=data.get("registration_address", ""))
    await callback.message.edit_text("Принято ✅")
    await ac_finalize(callback.message, state)


@router.callback_query(F.data == "ac_actual_diff", AddClient.waiting_actual_address_choice)
async def ac_actual_diff(callback: CallbackQuery, state: FSMContext):
    await callback.message.edit_text("Введите фактический адрес:")
    await state.set_state(AddClient.waiting_actual_address_text)


@router.message(AddClient.waiting_actual_address_text)
async def ac_actual_address_text(message: Message, state: FSMContext):
    await state.update_data(actual_address=message.text)
    await ac_finalize(message, state)


async def ac_finalize(target: Message, state: FSMContext):
    data = await state.get_data()
    client_id, tg_id = create_manual_client(
        data.get("name"), data.get("phone"),
        data.get("passport_series", ""), data.get("passport_number", ""),
        data.get("registration_address", ""), data.get("photo_main", ""),
        data.get("photo_reg", ""), data.get("dob", ""), data.get("issued_by", ""),
        data.get("issue_date", ""), data.get("department_code", ""),
        data.get("actual_address", "")
    )
    await state.clear()
    await target.answer(
        f"✅ Клиент добавлен (id {client_id}): {data.get('name')}, тел. {data.get('phone')}\n\n"
        f"Теперь можно оформить аренду командой /manual_rent"
    )


@router.message(Command("manual_rent"))
async def cmd_manual_rent(message: Message, state: FSMContext):
    if not admin_only(message.from_user.id):
        return
    clients = get_manual_clients()
    if not clients:
        await message.answer("Нет клиентов без Telegram. Сначала добавьте через /add_client")
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"{c['name']} ({c['phone']})", callback_data=f"manrent_client_{c['id']}")]
        for c in clients
    ])
    await message.answer("Выберите клиента:", reply_markup=kb)
    await state.set_state(ManualRent.choosing_client)


@router.callback_query(F.data.startswith("manrent_client_"), ManualRent.choosing_client)
async def manrent_choose_client(callback: CallbackQuery, state: FSMContext):
    client_id = callback.data.split("_")[-1]
    await state.update_data(manual_client_id=client_id)
    free_bikes = get_free_bikes()
    if not free_bikes:
        await callback.message.edit_text("Сейчас нет свободных велосипедов.")
        await state.clear()
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=b["name_model"], callback_data=f"manrent_bike_{b['id']}")]
        for b in free_bikes
    ])
    await callback.message.edit_text("Выберите велосипед:", reply_markup=kb)
    await state.set_state(ManualRent.choosing_bike)


@router.callback_query(F.data.startswith("manrent_bike_"), ManualRent.choosing_bike)
async def manrent_choose_bike(callback: CallbackQuery, state: FSMContext):
    bike_id = callback.data.split("_")[-1]
    bike, _ = get_bike_by_id(bike_id)
    await state.update_data(manual_bike_id=bike_id)
    await callback.message.edit_text(
        f"Велосипед: {bike['name_model']}\nНа какой срок?",
        reply_markup=period_keyboard(bike, prefix="manrent_period_")
    )
    await state.set_state(ManualRent.choosing_period)


@router.callback_query(F.data.startswith("manrent_period_"), ManualRent.choosing_period)
async def manrent_finalize(callback: CallbackQuery, state: FSMContext):
    period = callback.data.split("_")[-1]
    data = await state.get_data()
    bike_id = data["manual_bike_id"]
    client_id = data["manual_client_id"]
    bike, _ = get_bike_by_id(bike_id)
    price = int(bike["price_week"] if period == "неделя" else bike["price_month"])
    client = get_client_by_id(client_id)

    rental_id, start_dt, end_dt = create_rental(bike_id, client_id, period, price, bike["owner"])

    contract_data = {
        "rental_id": rental_id,
        "contract_date": ru_date(date_cls.today()),
        "business_phone": BUSINESS_PHONE,
        "full_name": client["name"],
        "dob": client.get("dob", ""),
        "passport_series": client.get("passport_series", ""),
        "passport_number": client.get("passport_number", ""),
        "issued_by": client.get("passport_issued_by", ""),
        "issue_date": client.get("passport_issue_date", ""),
        "department_code": client.get("department_code", ""),
        "registration_address": client.get("registration_address", ""),
        "actual_address": client.get("actual_address", ""),
        "phone": client.get("phone", ""),
        "bike_name": bike["name_model"],
        "serial": extract_serial(bike["name_model"]),
        "start_date": start_dt.strftime("%d.%m.%Y"),
        "end_date": end_dt.strftime("%d.%m.%Y"),
        "price_week": bike["price_week"],
        "price_month": bike["price_month"],
        "daily_penalty": round(int(bike["price_week"]) / 7),
    }
    pdf_buf = generate_contract_pdf(contract_data)
    sent_doc = await bot.send_document(
        ADMIN_ID,
        BufferedInputFile(pdf_buf.read(), filename=f"dogovor_{rental_id}.pdf"),
        caption=f"📄 Договор аренды #{rental_id} (клиент без Telegram) — распечатайте для подписи"
    )
    log_contract(rental_id, client_id, sent_doc.document.file_id)

    await callback.message.edit_text(
        f"✅ Аренда #{rental_id} создана\n\n"
        f"Клиент: {client['name']}\n"
        f"Велосипед: {bike['name_model']}\n"
        f"Срок: {period}, сумма: {price}₽\n"
        f"До: {end_dt.strftime('%d.%m.%Y')}\n\n"
        f"Когда получите оплату — отметьте: /paid {rental_id}\n"
        f"Когда велосипед вернут — закройте: /return {rental_id}"
    )
    await state.clear()


@router.message(Command("set_photo"))
async def cmd_set_photo(message: Message, state: FSMContext):
    if not admin_only(message.from_user.id):
        return
    parts = message.text.split()
    if len(parts) != 2:
        await message.answer("Использование: /set_photo ID_велосипеда, затем пришлите фото следующим сообщением")
        return
    bike, _ = get_bike_by_id(parts[1])
    if not bike:
        await message.answer("Велосипед не найден")
        return
    await state.update_data(photo_bike_id=parts[1])
    await message.answer(f"Пришлите фото для {bike['name_model']}:")
    await state.set_state(SetPhoto.waiting_photo)


@router.message(SetPhoto.waiting_photo, F.photo)
async def photo_received(message: Message, state: FSMContext):
    data = await state.get_data()
    bike_id = data["photo_bike_id"]
    file_id = message.photo[-1].file_id
    if set_bike_photo(bike_id, file_id):
        await message.answer("Фото сохранено ✅")
    else:
        await message.answer("Не удалось сохранить, велосипед не найден")
    await state.clear()


@router.message(Command("reset_test"))
async def cmd_reset_test(message: Message):
    if not admin_only(message.from_user.id):
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="⚠️ Да, очистить", callback_data="confirm_reset"),
        InlineKeyboardButton(text="Отмена", callback_data="cancel_reset"),
    ]])
    await message.answer(
        "Это удалит все записи в листах «Аренды», «Финансы», «Ожидание» и вернёт "
        "все велосипеды в статус «свободен».\n"
        "Лист «Клиенты» не тронется.\n\nПродолжить?",
        reply_markup=kb
    )


@router.callback_query(F.data == "cancel_reset")
async def cancel_reset(callback: CallbackQuery):
    await callback.message.edit_text("Отменено, ничего не удалено.")


@router.callback_query(F.data == "confirm_reset")
async def confirm_reset(callback: CallbackQuery):
    if not admin_only(callback.from_user.id):
        await callback.answer("Только для администратора", show_alert=True)
        return
    reset_test_data()
    await callback.message.edit_text(
        "✅ Тестовые данные очищены.\nАренды, финансы и лист ожидания пусты, все велосипеды свободны."
    )


@router.message(Command("contract"))
async def cmd_contract(message: Message):
    if not admin_only(message.from_user.id):
        return
    parts = message.text.split()
    if len(parts) != 2:
        await message.answer("Использование: /contract ID_аренды")
        return
    contract = get_contract_by_rental(parts[1])
    if not contract:
        await message.answer("Договор для этой аренды не найден")
        return
    await message.answer_document(contract["file_id"], caption=f"Договор по аренде #{parts[1]}")


@router.message(Command("qr"))
async def cmd_qr(message: Message):
    if not admin_only(message.from_user.id):
        return
    parts = message.text.split()
    if len(parts) != 2:
        await message.answer("Использование: /qr ID_велосипеда")
        return
    bike, _ = get_bike_by_id(parts[1])
    if not bike:
        await message.answer("Велосипед не найден")
        return
    buf, link = generate_qr_image(parts[1])
    await message.answer_photo(
        BufferedInputFile(buf.read(), filename=f"qr_{parts[1]}.png"),
        caption=f"QR для {bike['name_model']}\nСсылка: {link}\n\nРаспечатайте и приклейте на велосипед"
    )


# ------------------ НАПОМИНАНИЯ ------------------

async def notify_and_clear_waitlist():
    rows = get_waitlist()
    if not rows:
        return
    for r in rows:
        tg_id = r.get("telegram_id")
        if not tg_id:
            continue
        try:
            await bot.send_message(tg_id, "🚲 Велосипед освободился! Успейте забронировать: /rent")
        except Exception as e:
            log.warning(f"Не удалось уведомить {tg_id}: {e}")
    clear_waitlist()


async def check_reminders():
    rows = get_active_rentals()
    today = datetime.now().date()
    for r in rows:
        if r.get("return_status") != "арендован":
            continue
        try:
            end_date = datetime.strptime(r["end_date"], "%d.%m.%Y").date()
        except Exception:
            continue
        days_left = (end_date - today).days
        tg_id = get_client_telegram_id(r["client_id"])
        client_reachable = bool(tg_id) and not is_manual_tg_id(tg_id)
        unpaid = r.get("payment_status") != "оплачено"
        if days_left == 1:
            if client_reachable:
                await bot.send_message(
                    tg_id,
                    f"⏰ Завтра ({r['end_date']}) заканчивается аренда (сумма {r['amount']}₽).",
                    reply_markup=rental_action_buttons(r["id"], unpaid)
                )
            else:
                await bot.send_message(
                    ADMIN_ID,
                    f"⏰ Клиент без Telegram: завтра ({r['end_date']}) заканчивается аренда #{r['id']}"
                )
        elif days_left == 0:
            if client_reachable:
                await bot.send_message(
                    tg_id,
                    "⚠️ Сегодня последний день аренды. Продлите или верните велосипед.",
                    reply_markup=rental_action_buttons(r["id"], unpaid)
                )
            await bot.send_message(ADMIN_ID, f"⚠️ Сегодня истекает аренда #{r['id']}")
        elif days_left < 0:
            await bot.send_message(
                ADMIN_ID,
                f"🔴 Просрочка! Аренда #{r['id']} истекла {r['end_date']}, оплата: {r['payment_status']}"
            )
            if unpaid and client_reachable:
                await bot.send_message(
                    tg_id,
                    f"🔴 У вас долг по аренде #{r['id']}: {r['amount']}₽, просрочено с {r['end_date']}. "
                    f"Пожалуйста, оплатите или свяжитесь с нами: {BUSINESS_PHONE}\n\n{PAYMENT_INFO}",
                    reply_markup=paid_button(r["id"])
                )


async def check_maintenance():
    bikes = bikes_ws.get_all_records()
    today = datetime.now().date()
    for b in bikes:
        if b.get("status") == "на обслуживании":
            continue
        last_to = b.get("last_TO")
        if not last_to:
            continue
        try:
            last_to_date = datetime.strptime(str(last_to), "%d.%m.%Y").date()
        except Exception:
            continue
        days_since = (today - last_to_date).days
        if days_since >= TO_INTERVAL_DAYS and (days_since - TO_INTERVAL_DAYS) % 7 == 0:
            await bot.send_message(
                ADMIN_ID,
                f"🔧 Пора на ТО: {b['name_model']} (id {b['id']}) — последнее обслуживание {last_to}, "
                f"прошло {days_since} дней"
            )


async def setup_commands():
    client_commands = [
        BotCommand(command="start", description="Начать / регистрация"),
        BotCommand(command="rent", description="Арендовать велосипед"),
        BotCommand(command="my", description="Мои аренды"),
    ]
    await bot.set_my_commands(client_commands, scope=BotCommandScopeDefault())

    admin_commands = client_commands + [
        BotCommand(command="admin", description="Открыть админ-панель"),
        BotCommand(command="rentals", description="Активные аренды"),
        BotCommand(command="paid", description="Отметить оплату (ID)"),
        BotCommand(command="return", description="Закрыть аренду (ID)"),
        BotCommand(command="expense", description="Записать расход"),
        BotCommand(command="report", description="Финансовый отчёт"),
        BotCommand(command="stats", description="Статистика"),
        BotCommand(command="bike_history", description="История велосипеда (ID)"),
        BotCommand(command="blacklist", description="Чёрный список"),
        BotCommand(command="broadcast", description="Рассылка клиентам"),
        BotCommand(command="set_photo", description="Добавить фото велосипеда"),
        BotCommand(command="qr", description="QR-код велосипеда"),
        BotCommand(command="contract", description="Получить договор аренды (ID)"),
        BotCommand(command="add_client", description="Добавить клиента без Telegram"),
        BotCommand(command="manual_rent", description="Оформить аренду вручную"),
        BotCommand(command="reset_test", description="Очистить тестовые данные"),
    ]
    await bot.set_my_commands(admin_commands, scope=BotCommandScopeChat(chat_id=ADMIN_ID))

    if FRIEND_ID:
        friend_commands = [
            BotCommand(command="start", description="Начать"),
            BotCommand(command="report", description="Мой отчёт"),
        ]
        await bot.set_my_commands(friend_commands, scope=BotCommandScopeChat(chat_id=FRIEND_ID))


async def main():
    await setup_commands()
    scheduler = AsyncIOScheduler(timezone="Europe/Volgograd")
    scheduler.add_job(check_reminders, "cron", hour=10, minute=0)
    scheduler.add_job(check_maintenance, "cron", hour=10, minute=5)
    scheduler.start()
    log.info("Бот запущен")
    await dp.start_polling(bot)


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
