import os
import json
import base64
import logging
from datetime import datetime, timedelta

import gspread
from google.oauth2.service_account import Credentials
from apscheduler.schedulers.asyncio import AsyncIOScheduler

from contract import generate_contract_pdf, ru_date
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

BUSINESS_PHONE = "8-960-896-06-06"

creds_dict = json.loads(os.environ["GOOGLE_CREDS"])
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
gc = gspread.authorize(credentials)
sh = gc.open_by_key(SHEET_ID)

bikes_ws = sh.worksheet("Велосипеды")
clients_ws = sh.worksheet("Клиенты")
rentals_ws = sh.worksheet("Аренды")
finances_ws = sh.worksheet("Финансы")

PERIOD_DAYS = {"Неделя": 7, "Месяц": 30}
BIKE_ID_COL = "ID (номер велосипеда)"


def is_operator(user_id):
    return user_id == ADMIN_ID or (FRIEND_ID and user_id == FRIEND_ID)


# ------------------ ВЕЛОСИПЕДЫ ------------------

def get_free_bikes():
    rows = bikes_ws.get_all_records()
    return [r for r in rows if str(r.get("Статус", "")).strip() == "Свободен"]


def get_bike_by_id(bike_id):
    rows = bikes_ws.get_all_records()
    for i, r in enumerate(rows, start=2):
        if str(r.get(BIKE_ID_COL)) == str(bike_id):
            return r, i
    return None, None


def bike_display_name(bike):
    return f"{bike.get(BIKE_ID_COL)} — {bike.get('Модель')}"


def set_bike_status(bike_id, status):
    _, row_idx = get_bike_by_id(bike_id)
    if row_idx:
        col = bikes_ws.find("Статус").col
        bikes_ws.update_cell(row_idx, col, status)


# ------------------ КЛИЕНТЫ ------------------

def _normalize_phone(phone):
    digits = "".join(ch for ch in str(phone) if ch.isdigit())
    return digits[-10:] if len(digits) >= 10 else digits


def _next_client_id(rows):
    return f"К-{len(rows) + 1:03d}"


def get_client_by_phone(phone):
    target = _normalize_phone(phone)
    if not target:
        return None
    rows = clients_ws.get_all_records()
    for r in rows:
        if _normalize_phone(r.get("Телефон", "")) == target:
            return r
    return None


def get_client_by_id(client_id):
    rows = clients_ws.get_all_records()
    for r in rows:
        if str(r.get("ID клиента")) == str(client_id):
            return r
    return None


def create_client(name, phone, **fields):
    rows = clients_ws.get_all_records()
    new_id = _next_client_id(rows)
    clients_ws.append_row([
        new_id, "", name, phone,
        fields.get("dob", ""), fields.get("passport_series", ""),
        fields.get("passport_number", ""), fields.get("issued_by", ""),
        fields.get("issue_date", ""), fields.get("department_code", ""),
        fields.get("registration_address", ""), fields.get("actual_address", ""),
        fields.get("photo_main", ""), fields.get("photo_reg", ""),
        datetime.now().strftime("%d.%m.%Y"), "Нет", "Обычный", ""
    ])
    return new_id


# ------------------ АРЕНДЫ ------------------

def _next_rental_id(rows):
    return f"А-{len(rows) + 1:04d}"


def create_rental(bike_id, client_id, period, price):
    rows = rentals_ws.get_all_records()
    new_id = _next_rental_id(rows)
    start = datetime.now()
    end = start + timedelta(days=PERIOD_DAYS[period])

    rentals_ws.append_row([
        new_id, client_id, "", bike_id, "",
        period, start.strftime("%d.%m.%Y"), end.strftime("%d.%m.%Y"),
        "Арендован", price, "Оплачено", "Нет", "", ""
    ])
    new_row = len(rows) + 2
    rentals_ws.update_cell(new_row, 3, f'=IFERROR(INDEX(Клиенты!C:C,MATCH(B{new_row},Клиенты!A:A,0)),"")')
    rentals_ws.update_cell(new_row, 5, f'=IFERROR(INDEX(Велосипеды!C:C,MATCH(D{new_row},Велосипеды!A:A,0)),"")')

    set_bike_status(bike_id, "В аренде")
    add_finance_row("Доход", "аренда", price, get_bike_by_id(bike_id)[0]["Владелец"], f"аренда {new_id}", bike_id, new_id)
    return new_id, start, end


def get_rental_by_id(rental_id):
    rows = rentals_ws.get_all_records()
    for i, r in enumerate(rows, start=2):
        if str(r.get("ID аренды")) == str(rental_id):
            return r, i
    return None, None


def get_active_rentals():
    return rentals_ws.get_all_records()


def get_rental_column(name):
    return rentals_ws.find(name).col


def mark_returned(rental_id):
    rental, row_idx = get_rental_by_id(rental_id)
    if row_idx:
        rentals_ws.update_cell(row_idx, get_rental_column("Статус аренды"), "Возвращён")
        set_bike_status(rental.get("Номер велосипеда"), "Свободен")
        return True
    return False


def mark_extended(rental_id):
    _, row_idx = get_rental_by_id(rental_id)
    if row_idx:
        rentals_ws.update_cell(row_idx, get_rental_column("Статус аренды"), "Продлена")


# ------------------ ФИНАНСЫ ------------------

def add_finance_row(ftype, category, amount, owner, comment, bike_id="", rental_id=""):
    rows = finances_ws.get_all_records()
    new_id = len(rows) + 1
    finances_ws.append_row([new_id, datetime.now().strftime("%d.%m.%Y"), ftype, category, amount, owner, comment, bike_id, rental_id])


# ------------------ КЛАВИАТУРЫ ------------------

def main_menu_kb():
    return ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="🚲 Арендовать велосипед")]], resize_keyboard=True)


def period_keyboard(bike):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"Неделя — {bike['Тариф за неделю (₽)']}₽", callback_data="period_Неделя")],
        [InlineKeyboardButton(text=f"Месяц — {bike['Тариф за месяц (₽)']}₽", callback_data="period_Месяц")],
    ])


def reminder_buttons(rental_id):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔄 Продлил", callback_data=f"rem_extend_{rental_id}")],
        [InlineKeyboardButton(text="🕐 Просит отсрочку", callback_data=f"rem_delay_{rental_id}")],
        [InlineKeyboardButton(text="🚲 Сдал", callback_data=f"rem_return_{rental_id}")],
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


# ------------------ БОТ ------------------

storage = MemoryStorage()
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=storage)
router = Router()
dp.include_router(router)

FIELD_LABELS = [
    ("full_name", "ФИО"), ("dob", "Дата рождения"), ("passport_series", "Серия паспорта"),
    ("passport_number", "Номер паспорта"), ("issued_by", "Кем выдан"), ("issue_date", "Дата выдачи"),
    ("department_code", "Код подразделения"), ("registration_address", "Адрес регистрации"),
]


def recognition_summary_text(rec):
    lines = ["Распознано:\n"]
    for key, label in FIELD_LABELS:
        lines.append(f"{label}: {rec.get(key) or '—'}")
    lines.append("\nВсё верно?")
    return "\n".join(lines)


def recognition_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Всё верно", callback_data="recog_ok")],
        [InlineKeyboardButton(text="✏️ Исправить поле", callback_data="edit_menu")],
        [InlineKeyboardButton(text="📝 Заполнить вручную", callback_data="recog_manual")],
    ])


def edit_menu_kb(rec):
    buttons = [
        [InlineKeyboardButton(text=f"{label}: {rec.get(key) or '—'}", callback_data=f"editfield_{key}")]
        for key, label in FIELD_LABELS
    ]
    buttons.append([InlineKeyboardButton(text="⬅️ Назад", callback_data="edit_back")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


class EditField(StatesGroup):
    waiting_value = State()


class AdminRent(StatesGroup):
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
    choosing_period = State()


@router.callback_query(F.data == "edit_menu")
async def show_edit_menu(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    rec = data.get("recognized", {})
    await callback.message.edit_text("Какое поле исправить?", reply_markup=edit_menu_kb(rec))


@router.callback_query(F.data == "edit_back")
async def edit_back(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    rec = data.get("recognized", {})
    await callback.message.edit_text(recognition_summary_text(rec), reply_markup=recognition_kb())


@router.callback_query(F.data.startswith("editfield_"))
async def edit_field_prompt(callback: CallbackQuery, state: FSMContext):
    field_key = callback.data.split("editfield_", 1)[1]
    await state.update_data(edit_field=field_key)
    label = dict(FIELD_LABELS).get(field_key, field_key)
    await callback.message.edit_text(f"Введите новое значение для «{label}»:")
    await state.set_state(EditField.waiting_value)


@router.message(EditField.waiting_value)
async def edit_field_value(message: Message, state: FSMContext):
    data = await state.get_data()
    rec = data.get("recognized", {})
    rec[data.get("edit_field")] = message.text
    await state.update_data(recognized=rec)
    await state.set_state(AdminRent.confirming_recognition)
    await message.answer(recognition_summary_text(rec), reply_markup=recognition_kb())


# ------------------ СТАРТ / ГЛАВНОЕ МЕНЮ ------------------

@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    if not is_operator(message.from_user.id):
        return
    await message.answer("Elektro VLG — бот учёта аренды.", reply_markup=main_menu_kb())


@router.message(F.text == "🚲 Арендовать велосипед")
async def cmd_rent(message: Message, state: FSMContext):
    if not is_operator(message.from_user.id):
        return
    free_bikes = get_free_bikes()
    if not free_bikes:
        await message.answer("Свободных велосипедов сейчас нет.")
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=bike_display_name(b), callback_data=f"bike_{b[BIKE_ID_COL]}")]
        for b in free_bikes
    ])
    await message.answer("Выберите велосипед:", reply_markup=kb)


@router.callback_query(F.data.startswith("bike_"))
async def choose_bike(callback: CallbackQuery, state: FSMContext):
    if not is_operator(callback.from_user.id):
        return
    bike_id = callback.data.split("_", 1)[1]
    set_bike_status(bike_id, "В аренде")  # сразу занимаем, как договорились
    await state.update_data(bike_id=bike_id)
    await callback.message.edit_text(f"Велосипед {bike_id} закреплён за этой арендой.\n\nНомер телефона клиента:")
    await state.set_state(AdminRent.waiting_phone)


@router.message(AdminRent.waiting_phone)
async def rent_phone(message: Message, state: FSMContext):
    phone = message.text
    existing = get_client_by_phone(phone)
    if existing:
        await state.update_data(client_id=existing["ID клиента"], phone=phone)
        bike_id = (await state.get_data())["bike_id"]
        bike, _ = get_bike_by_id(bike_id)
        await message.answer(f"Клиент найден: {existing['ФИО']} ✅\n\nНа какой срок?", reply_markup=period_keyboard(bike))
        await state.set_state(AdminRent.choosing_period)
        return

    await state.update_data(phone=phone)
    await message.answer("Новый клиент. Пришлите фото главной страницы паспорта 📸")
    await state.set_state(AdminRent.waiting_photo_main)


@router.message(AdminRent.waiting_photo_main, F.photo)
async def rent_photo_main(message: Message, state: FSMContext):
    await state.update_data(photo_main=message.photo[-1].file_id)
    await message.answer("Теперь фото страницы с пропиской 📸")
    await state.set_state(AdminRent.waiting_photo_reg)


@router.message(AdminRent.waiting_photo_reg, F.photo)
async def rent_photo_reg(message: Message, state: FSMContext):
    data = await state.get_data()
    photo_reg_id = message.photo[-1].file_id
    await state.update_data(photo_reg=photo_reg_id)

    recognized = None
    if openai_client:
        wait_msg = await message.answer("Распознаю паспорт... ⏳")
        recognized = await recognize_passport(message.bot, data["photo_main"], photo_reg_id)
        try:
            await wait_msg.delete()
        except Exception:
            pass

    if recognized:
        await state.update_data(recognized=recognized)
        await message.answer(recognition_summary_text(recognized), reply_markup=recognition_kb())
        await state.set_state(AdminRent.confirming_recognition)
    else:
        await message.answer("Не распознано, заполним вручную. ФИО клиента:")
        await state.set_state(AdminRent.waiting_dob)  # ФИО спросим первым полем вручную ниже


@router.callback_query(F.data == "recog_ok", AdminRent.confirming_recognition)
async def recog_confirm(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    rec = data.get("recognized", {})
    await state.update_data(
        name=rec.get("full_name", ""), dob=rec.get("dob", ""),
        passport_series=rec.get("passport_series", ""), passport_number=rec.get("passport_number", ""),
        issued_by=rec.get("issued_by", ""), issue_date=rec.get("issue_date", ""),
        department_code=rec.get("department_code", ""), registration_address=rec.get("registration_address", "")
    )
    await callback.message.edit_text("Принято ✅")
    await ask_actual_address(callback.message, state)


@router.callback_query(F.data == "recog_manual", AdminRent.confirming_recognition)
async def recog_manual(callback: CallbackQuery, state: FSMContext):
    await callback.message.edit_text("Заполним вручную. ФИО клиента:")
    await state.set_state(AdminRent.waiting_dob)


@router.message(AdminRent.waiting_dob)
async def rent_name_or_dob(message: Message, state: FSMContext):
    data = await state.get_data()
    if "name" not in data:
        await state.update_data(name=message.text)
        await message.answer("Дата рождения (например 15.03.1990):")
        return
    await state.update_data(dob=message.text)
    await message.answer("Серия и номер паспорта одним сообщением (например 1234 567890):")
    await state.set_state(AdminRent.waiting_passport)


@router.message(AdminRent.waiting_passport)
async def rent_passport(message: Message, state: FSMContext):
    parts = message.text.split(maxsplit=1)
    series = parts[0] if parts else message.text
    number = parts[1] if len(parts) > 1 else ""
    await state.update_data(passport_series=series, passport_number=number)
    await message.answer("Кем выдан?")
    await state.set_state(AdminRent.waiting_issued_by)


@router.message(AdminRent.waiting_issued_by)
async def rent_issued_by(message: Message, state: FSMContext):
    await state.update_data(issued_by=message.text)
    await message.answer("Дата выдачи паспорта:")
    await state.set_state(AdminRent.waiting_issue_date)


@router.message(AdminRent.waiting_issue_date)
async def rent_issue_date(message: Message, state: FSMContext):
    await state.update_data(issue_date=message.text)
    await message.answer("Код подразделения:")
    await state.set_state(AdminRent.waiting_department_code)


@router.message(AdminRent.waiting_department_code)
async def rent_department_code(message: Message, state: FSMContext):
    await state.update_data(department_code=message.text)
    await message.answer("Адрес регистрации (прописка):")
    await state.set_state(AdminRent.waiting_reg_address)


@router.message(AdminRent.waiting_reg_address)
async def rent_reg_address(message: Message, state: FSMContext):
    await state.update_data(registration_address=message.text)
    await ask_actual_address(message, state)


async def ask_actual_address(target, state: FSMContext):
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Совпадает с пропиской", callback_data="actual_same")],
        [InlineKeyboardButton(text="✏️ Другой адрес", callback_data="actual_diff")],
    ])
    await target.answer("Фактический адрес такой же, как прописка?", reply_markup=kb)
    await state.set_state(AdminRent.waiting_actual_address_choice)


@router.callback_query(F.data == "actual_same", AdminRent.waiting_actual_address_choice)
async def actual_same(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    await state.update_data(actual_address=data.get("registration_address", ""))
    await callback.message.edit_text("Принято ✅")
    await proceed_to_period(callback.message, state)


@router.callback_query(F.data == "actual_diff", AdminRent.waiting_actual_address_choice)
async def actual_diff(callback: CallbackQuery, state: FSMContext):
    await callback.message.edit_text("Введите фактический адрес:")
    await state.set_state(AdminRent.waiting_actual_address_text)


@router.message(AdminRent.waiting_actual_address_text)
async def actual_address_text(message: Message, state: FSMContext):
    await state.update_data(actual_address=message.text)
    await proceed_to_period(message, state)


async def proceed_to_period(target, state: FSMContext):
    data = await state.get_data()
    client_id = create_client(
        data.get("name", ""), data.get("phone", ""),
        dob=data.get("dob", ""), passport_series=data.get("passport_series", ""),
        passport_number=data.get("passport_number", ""), issued_by=data.get("issued_by", ""),
        issue_date=data.get("issue_date", ""), department_code=data.get("department_code", ""),
        registration_address=data.get("registration_address", ""), actual_address=data.get("actual_address", ""),
        photo_main=data.get("photo_main", ""), photo_reg=data.get("photo_reg", "")
    )
    await state.update_data(client_id=client_id)
    bike, _ = get_bike_by_id(data["bike_id"])
    await target.answer(f"Клиент сохранён ({client_id}) ✅\n\nНа какой срок?", reply_markup=period_keyboard(bike))
    await state.set_state(AdminRent.choosing_period)


@router.callback_query(F.data.startswith("period_"), AdminRent.choosing_period)
async def choose_period(callback: CallbackQuery, state: FSMContext):
    period = callback.data.split("_", 1)[1]
    data = await state.get_data()
    bike_id, client_id = data["bike_id"], data["client_id"]
    bike, _ = get_bike_by_id(bike_id)
    price = int(bike["Тариф за неделю (₽)"] if period == "Неделя" else bike["Тариф за месяц (₽)"])
    client = get_client_by_id(client_id)

    rental_id, start_dt, end_dt = create_rental(bike_id, client_id, period, price)

    contract_data = {
        "rental_id": rental_id, "contract_date": ru_date(date_cls.today()), "business_phone": BUSINESS_PHONE,
        "full_name": client["ФИО"], "dob": client.get("Дата рождения", ""),
        "passport_series": client.get("Серия паспорта", ""), "passport_number": client.get("Номер паспорта", ""),
        "issued_by": client.get("Кем выдан", ""), "issue_date": client.get("Дата выдачи", ""),
        "department_code": client.get("Код подразделения", ""),
        "registration_address": client.get("Адрес регистрации", ""), "actual_address": client.get("Фактический адрес", ""),
        "phone": client.get("Телефон", ""), "bike_id": bike_id, "bike_name": bike["Модель"],
        "start_date": start_dt.strftime("%d.%m.%Y"), "end_date": end_dt.strftime("%d.%m.%Y"),
        "price_week": bike["Тариф за неделю (₽)"], "price_month": bike["Тариф за месяц (₽)"],
    }
    pdf_buf = generate_contract_pdf(contract_data)
    await callback.message.answer_document(
        BufferedInputFile(pdf_buf.read(), filename=f"dogovor_{rental_id}.pdf"),
        caption=f"📄 Договор {rental_id} — распечатайте для подписи"
    )
    await callback.message.edit_text(
        f"✅ Аренда {rental_id} оформлена\n\n"
        f"{bike_display_name(bike)} — {period}, {price}₽\n"
        f"Клиент: {client['ФИО']}, {client['Телефон']}\n"
        f"Оплата зачтена. Вернуть/оплатить до: {end_dt.strftime('%d.%m.%Y')}"
    )
    await state.clear()


# ------------------ НАПОМИНАНИЯ ------------------

@router.callback_query(F.data.startswith("rem_extend_"))
async def rem_extend(callback: CallbackQuery):
    if not is_operator(callback.from_user.id):
        return
    rental_id = callback.data.split("rem_extend_", 1)[1]
    rental, _ = get_rental_by_id(rental_id)
    if not rental or rental.get("Статус аренды") != "Арендован":
        await callback.answer("Эта аренда уже закрыта", show_alert=True)
        return
    bike, _ = get_bike_by_id(rental["Номер велосипеда"])
    period = rental["Тип аренды"]
    price = int(bike["Тариф за неделю (₽)"] if period == "Неделя" else bike["Тариф за месяц (₽)"])

    mark_extended(rental_id)
    new_id, start_dt, end_dt = create_rental(rental["Номер велосипеда"], rental["ID клиента"], period, price)

    await callback.message.edit_text(
        callback.message.text + f"\n\n✅ Продлено до {end_dt.strftime('%d.%m.%Y')} (новая аренда {new_id})"
    )


@router.callback_query(F.data.startswith("rem_delay_"))
async def rem_delay(callback: CallbackQuery):
    if not is_operator(callback.from_user.id):
        return
    rental_id = callback.data.split("rem_delay_", 1)[1]
    rental, row_idx = get_rental_by_id(rental_id)
    if not rental or rental.get("Статус аренды") != "Арендован":
        await callback.answer("Эта аренда уже закрыта", show_alert=True)
        return
    col = get_rental_column("Комментарий")
    rentals_ws.update_cell(row_idx, col, f"Отсрочка запрошена {datetime.now().strftime('%d.%m.%Y')}")
    await callback.message.edit_text(callback.message.text + "\n\n🕐 Отмечено: отсрочка. Напомню про эту аренду завтра снова.")


@router.callback_query(F.data.startswith("rem_return_"))
async def rem_return(callback: CallbackQuery):
    if not is_operator(callback.from_user.id):
        return
    rental_id = callback.data.split("rem_return_", 1)[1]
    rental, _ = get_rental_by_id(rental_id)
    if not rental:
        await callback.answer("Аренда не найдена", show_alert=True)
        return
    mark_returned(rental_id)
    await callback.message.edit_text(callback.message.text + "\n\n🚲 Сдал — велосипед снова свободен ✅")


async def check_reminders():
    rows = get_active_rentals()
    today = datetime.now().date()
    for r in rows:
        if r.get("Статус аренды") != "Арендован":
            continue
        try:
            end_date = datetime.strptime(r["Дата окончания (план)"], "%d.%m.%Y").date()
        except Exception:
            continue
        if (end_date - today).days > 0:
            continue

        client = get_client_by_id(r["ID клиента"])
        phone = client.get("Телефон", "—") if client else "—"
        text = f"{r['ФИО клиента']}\n{phone}\n\nСегодня оплата."
        for uid in filter(None, [ADMIN_ID, FRIEND_ID]):
            try:
                await bot.send_message(uid, text, reply_markup=reminder_buttons(r["ID аренды"]))
            except Exception as e:
                log.warning(f"Не удалось отправить напоминание {uid}: {e}")


async def setup_commands():
    # Перезаписываем общий список команд — убирает старые /rentals, /stats и т.д.
    await bot.set_my_commands([BotCommand(command="start", description="Открыть меню")], scope=BotCommandScopeDefault())

    # ВАЖНО: старая версия бота регистрировала расширенный список команд отдельно
    # для личных чатов админа и партнёра (BotCommandScopeChat) — такие команды
    # перекрывают общий список и не удаляются его обновлением. Стираем их явно.
    try:
        await bot.delete_my_commands(scope=BotCommandScopeChat(chat_id=ADMIN_ID))
    except Exception as e:
        log.warning(f"Не удалось очистить команды у ADMIN_ID: {e}")
    if FRIEND_ID:
        try:
            await bot.delete_my_commands(scope=BotCommandScopeChat(chat_id=FRIEND_ID))
        except Exception as e:
            log.warning(f"Не удалось очистить команды у FRIEND_ID: {e}")


async def main():
    await setup_commands()
    scheduler = AsyncIOScheduler(timezone="Europe/Volgograd")
    scheduler.add_job(check_reminders, "cron", hour=10, minute=0)
    scheduler.start()
    log.info("Бот запущен")
    await dp.start_polling(bot)


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
