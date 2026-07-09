import os
import json
import logging
from datetime import datetime, timedelta

import gspread
from google.oauth2.service_account import Credentials
from apscheduler.schedulers.asyncio import AsyncIOScheduler

from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton
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

BUSINESS_ADDRESS = "г. Волгоград, ул. Киргизская 2"
BUSINESS_PHONE = "8-960-896-06-06"
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


ensure_all_columns()
waiting_ws = ensure_worksheet("Ожидание", ["id", "telegram_id", "name", "date_added"])

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

def get_or_create_client(tg_id, name, phone):
    rows = clients_ws.get_all_records()
    for r in rows:
        if str(r.get("telegram_id")) == str(tg_id):
            return r.get("id")
    new_id = len(rows) + 1
    clients_ws.append_row([new_id, tg_id, name, phone, datetime.now().strftime("%d.%m.%Y"), "нет"])
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


# ------------------ ДОГОВОР ------------------

def generate_contract_text(rental_id, client_name, phone, bike_name, period, price,
                            start, end, delivery, delivery_address):
    pickup = (f"доставка по адресу: {delivery_address}" if delivery == "да"
              else f"самовывоз, {BUSINESS_ADDRESS}")
    return (
        f"📄 Договор аренды электровелосипеда №{rental_id}\n\n"
        f"Арендатор: {client_name}, тел. {phone or '—'}\n"
        f"Предмет аренды: {bike_name}\n"
        f"Срок аренды: {period} (с {start} по {end})\n"
        f"Стоимость: {price}₽\n"
        f"Получение: {pickup}\n"
        f"Контакт арендодателя: {BUSINESS_PHONE}\n\n"
        f"Арендатор обязуется бережно эксплуатировать велосипед и вернуть его в исправном "
        f"состоянии в указанный срок. В случае повреждения по вине арендатора стоимость "
        f"ремонта оплачивается арендатором.\n\n"
        f"Подтверждая аренду в боте, арендатор соглашается с условиями договора."
    )


# ------------------ КЛАВИАТУРЫ ------------------

def paid_button(rental_id):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Я оплатил(а)", callback_data=f"claim_paid_{rental_id}")]
    ])


def rental_action_buttons(rental_id, unpaid):
    buttons = []
    if unpaid:
        buttons.append([InlineKeyboardButton(text="✅ Я оплатил(а)", callback_data=f"claim_paid_{rental_id}")])
    buttons.append([InlineKeyboardButton(text="🔄 Продлить", callback_data=f"extend_{rental_id}")])
    buttons.append([InlineKeyboardButton(text="🚲 Вернуть велосипед", callback_data=f"claim_return_{rental_id}")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def waitlist_button():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔔 Уведомить когда освободится", callback_data="waitlist_join")]
    ])


# ------------------ БОТ ------------------

storage = MemoryStorage()
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=storage)
router = Router()
dp.include_router(router)


class Registration(StatesGroup):
    waiting_name = State()
    waiting_phone = State()


class Rent(StatesGroup):
    choosing_bike = State()
    choosing_period = State()
    choosing_delivery = State()
    entering_delivery_address = State()


class Extend(StatesGroup):
    choosing_period = State()


class SetPhoto(StatesGroup):
    waiting_photo = State()


def admin_only(user_id):
    return user_id == ADMIN_ID


def get_role(user_id):
    if user_id == ADMIN_ID:
        return "admin"
    if FRIEND_ID and user_id == FRIEND_ID:
        return "friend"
    return None


@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    await message.answer(
        "Привет! Это бот проката электровелосипедов.\n\n"
        "Давай зарегистрируемся. Как тебя зовут?"
    )
    await state.set_state(Registration.waiting_name)


@router.message(Registration.waiting_name)
async def reg_name(message: Message, state: FSMContext):
    await state.update_data(name=message.text)
    await message.answer("Отлично! Теперь отправь номер телефона (просто текстом).")
    await state.set_state(Registration.waiting_phone)


@router.message(Registration.waiting_phone)
async def reg_phone(message: Message, state: FSMContext):
    data = await state.get_data()
    get_or_create_client(message.from_user.id, data["name"], message.text)
    await state.clear()
    await message.answer(
        "Регистрация завершена! ✅\n\n"
        "Команды:\n"
        "/rent — арендовать велосипед\n"
        "/my — мои аренды"
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
        await message.answer(
            "Сейчас все велосипеды заняты, попробуйте позже 🙁",
            reply_markup=waitlist_button()
        )
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
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"Неделя — {bike['price_week']}₽", callback_data="period_неделя")],
        [InlineKeyboardButton(text=f"Месяц — {bike['price_month']}₽", callback_data="period_месяц")],
    ])
    if bike.get("photo_id"):
        await callback.message.answer_photo(bike["photo_id"], caption=bike["name_model"])
    await callback.message.edit_text(f"Выбран: {bike['name_model']}\nНа какой срок?", reply_markup=kb)
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

    await bot.send_message(chat, "Нажмите, когда переведёте оплату:", reply_markup=paid_button(rental_id))

    contract = generate_contract_text(
        rental_id, user.full_name, client["phone"] if client else "",
        bike["name_model"], period, price,
        start_dt.strftime("%d.%m.%Y"), end_dt.strftime("%d.%m.%Y"),
        delivery, delivery_address
    )
    await bot.send_message(chat, contract)

    await bot.send_message(
        ADMIN_ID,
        f"🆕 Новая аренда #{rental_id}\n"
        f"{bike['name_model']} — {period} ({price}₽)\n"
        f"Клиент: {user.full_name} (id {client_id})\n"
        f"Владелец велика: {bike['owner']}\n"
        + (f"Доставка: {delivery_address}" if delivery == "да" else "Самовывоз")
    )
    await state.clear()


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
    if tg_id:
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
    if tg_id:
        await bot.send_message(tg_id, f"По аренде #{rental_id} мы не нашли поступление оплаты. Свяжитесь с нами.")


# ------------------ ПРОДЛЕНИЕ ------------------

@router.callback_query(F.data.startswith("extend_"))
async def extend_start(callback: CallbackQuery, state: FSMContext):
    rental_id = callback.data.split("_")[-1]
    rental, _ = get_rental_by_id(rental_id)
    bike, _ = get_bike_by_id(rental["bike_id"])
    await state.update_data(old_rental_id=rental_id, bike_id=rental["bike_id"])
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"Неделя — {bike['price_week']}₽", callback_data="ext_period_неделя")],
        [InlineKeyboardButton(text=f"Месяц — {bike['price_month']}₽", callback_data="ext_period_месяц")],
    ])
    await callback.message.answer("На какой срок продлить?", reply_markup=kb)
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
    await callback.message.answer("Оплата за продление:", reply_markup=paid_button(new_id))
    await bot.send_message(
        ADMIN_ID,
        f"🔄 Продление: старая аренда #{old_rental_id} закрыта, новая #{new_id} "
        f"({bike['name_model']}, {period}, {price}₽)"
    )
    await state.clear()


# ------------------ ВОЗВРАТ ------------------

@router.callback_query(F.data.startswith("claim_return_"))
async def claim_return(callback: CallbackQuery):
    rental_id = callback.data.split("_")[-1]
    rental, _ = get_rental_by_id(rental_id)
    if not rental or rental.get("return_status") != "арендован":
        await callback.answer("Аренда уже закрыта", show_alert=True)
        return
    admin_kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Подтвердить возврат", callback_data=f"confirm_return_{rental_id}"),
    ]])
    await callback.answer("Заявка на возврат отправлена администратору ✅", show_alert=True)
    await bot.send_message(
        ADMIN_ID,
        f"🚲 Клиент хочет вернуть велосипед по аренде #{rental_id}\n"
        f"Проверьте велосипед и подтвердите возврат:",
        reply_markup=admin_kb
    )


@router.callback_query(F.data.startswith("confirm_return_"))
async def confirm_return(callback: CallbackQuery):
    if not admin_only(callback.from_user.id):
        await callback.answer("Только для администратора", show_alert=True)
        return
    rental_id = callback.data.split("_")[-1]
    rental, _ = get_rental_by_id(rental_id)
    mark_returned(rental_id)
    await callback.message.edit_text(callback.message.text + "\n\n✅ Возврат подтверждён, велосипед свободен")
    tg_id = get_client_telegram_id(rental["client_id"])
    if tg_id:
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
            f"📊 Ваш отчёт\n\n"
            f"Доход: {income_friend}₽\n"
            f"Расходы: {expense_friend}₽\n"
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
        if not tg_id:
            continue
        unpaid = r.get("payment_status") != "оплачено"
        if days_left == 1:
            await bot.send_message(
                tg_id,
                f"⏰ Завтра ({r['end_date']}) заканчивается аренда (сумма {r['amount']}₽).",
                reply_markup=rental_action_buttons(r["id"], unpaid)
            )
        elif days_left == 0:
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
            if unpaid:
                await bot.send_message(
                    tg_id,
                    f"🔴 У вас долг по аренде #{r['id']}: {r['amount']}₽, просрочено с {r['end_date']}. "
                    f"Пожалуйста, оплатите или свяжитесь с нами: {BUSINESS_PHONE}",
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


async def main():
    scheduler = AsyncIOScheduler(timezone="Europe/Volgograd")
    scheduler.add_job(check_reminders, "cron", hour=10, minute=0)
    scheduler.add_job(check_maintenance, "cron", hour=10, minute=5)
    scheduler.start()
    log.info("Бот запущен")
    await dp.start_polling(bot)


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
