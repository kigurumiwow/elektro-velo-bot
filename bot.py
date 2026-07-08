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

# ------------------ НАСТРОЙКИ (берутся из переменных окружения) ------------------
BOT_TOKEN = os.environ["BOT_TOKEN"]
ADMIN_ID = int(os.environ["ADMIN_ID"])  # ваш Telegram ID
SHEET_ID = os.environ["SHEET_ID"]  # ID таблицы из ссылки (между /d/ и /edit)

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

# ------------------ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ РАБОТЫ С ТАБЛИЦЕЙ ------------------

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


def get_or_create_client(tg_id, name, phone):
    rows = clients_ws.get_all_records()
    for r in rows:
        if str(r.get("telegram_id")) == str(tg_id):
            return r.get("id")
    new_id = len(rows) + 1
    clients_ws.append_row([new_id, tg_id, name, phone, datetime.now().strftime("%d.%m.%Y")])
    return new_id


def create_rental(bike_id, client_id, period, price, owner):
    rows = rentals_ws.get_all_records()
    new_id = len(rows) + 1
    start = datetime.now()
    end = start + timedelta(days=PERIOD_DAYS[period])
    rentals_ws.append_row([
        new_id, bike_id, client_id,
        start.strftime("%d.%m.%Y"), period, end.strftime("%d.%m.%Y"),
        price, owner, "не оплачено", "арендован"
    ])
    set_bike_status(bike_id, "в аренде")
    return new_id, end


def get_client_telegram_id(client_id):
    rows = clients_ws.get_all_records()
    for r in rows:
        if str(r.get("id")) == str(client_id):
            return r.get("telegram_id")
    return None


def get_active_rentals():
    return rentals_ws.get_all_records()


def mark_paid(rental_id):
    rows = rentals_ws.get_all_records()
    for i, r in enumerate(rows, start=2):
        if str(r.get("id")) == str(rental_id):
            col = rentals_ws.find("payment_status").col
            rentals_ws.update_cell(i, col, "оплачено")
            return True
    return False


def mark_returned(rental_id):
    rows = rentals_ws.get_all_records()
    for i, r in enumerate(rows, start=2):
        if str(r.get("id")) == str(rental_id):
            col = rentals_ws.find("return_status").col
            rentals_ws.update_cell(i, col, "возвращён")
            set_bike_status(r.get("bike_id"), "свободен")
            return True
    return False


def add_expense(category, amount, owner, comment):
    rows = finances_ws.get_all_records()
    new_id = len(rows) + 1
    finances_ws.append_row([
        new_id, datetime.now().strftime("%d.%m.%Y"), "расход",
        category, amount, owner, comment
    ])
    return new_id


def get_finance_rows():
    return finances_ws.get_all_records()


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
    free_bikes = get_free_bikes()
    if not free_bikes:
        await message.answer("Сейчас все велосипеды заняты, попробуйте позже 🙁")
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=b["name_model"], callback_data=f"bike_{b['id']}")]
        for b in free_bikes
    ])
    await message.answer("Выберите велосипед:", reply_markup=kb)
    await state.set_state(Rent.choosing_bike)


@router.callback_query(F.data.startswith("bike_"))
async def choose_bike(callback: CallbackQuery, state: FSMContext):
    bike_id = callback.data.split("_")[1]
    bike, _ = get_bike_by_id(bike_id)
    await state.update_data(bike_id=bike_id)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text=f"Неделя — {bike['price_week']}₽",
            callback_data="period_неделя")],
        [InlineKeyboardButton(
            text=f"Месяц — {bike['price_month']}₽",
            callback_data="period_месяц")],
    ])
    await callback.message.edit_text(f"Выбран: {bike['name_model']}\nНа какой срок?", reply_markup=kb)
    await state.set_state(Rent.choosing_period)


@router.callback_query(F.data.startswith("period_"))
async def choose_period(callback: CallbackQuery, state: FSMContext):
    period = callback.data.split("_")[1]
    data = await state.get_data()
    bike_id = data["bike_id"]
    bike, _ = get_bike_by_id(bike_id)
    price = bike["price_week"] if period == "неделя" else bike["price_month"]

    client_id = get_or_create_client(callback.from_user.id, callback.from_user.full_name, "")
    rental_id, end_date = create_rental(bike_id, client_id, period, price, bike["owner"])

    await callback.message.edit_text(
        f"Готово! ✅\n\n"
        f"Велосипед: {bike['name_model']}\n"
        f"Срок: {period}\n"
        f"Сумма: {price}₽\n"
        f"Оплата: наличными/переводом\n"
        f"Вернуть/продлить до: {end_date.strftime('%d.%m.%Y')}\n\n"
        f"Я напомню заранее об окончании срока 🙂"
    )
    await bot.send_message(
        ADMIN_ID,
        f"🆕 Новая аренда #{rental_id}\n"
        f"{bike['name_model']} — {period} ({price}₽)\n"
        f"Клиент: {callback.from_user.full_name} (id {client_id})\n"
        f"Владелец велика: {bike['owner']}"
    )
    await state.clear()


@router.message(Command("my"))
async def cmd_my(message: Message):
    rows = get_active_rentals()
    client_rows = clients_ws.get_all_records()
    my_client_id = None
    for r in client_rows:
        if str(r.get("telegram_id")) == str(message.from_user.id):
            my_client_id = r.get("id")
    if my_client_id is None:
        await message.answer("Вы ещё не регистрировались — введите /start")
        return
    my_rentals = [
        r for r in rows
        if str(r.get("client_id")) == str(my_client_id) and r.get("return_status") == "арендован"
    ]
    if not my_rentals:
        await message.answer("Активных аренд нет.")
        return
    text = "Ваши активные аренды:\n\n"
    for r in my_rentals:
        text += (
            f"Аренда #{r['id']}\n"
            f"До: {r['end_date']}\n"
            f"Сумма: {r['amount']}₽\n"
            f"Оплата: {r['payment_status']}\n\n"
        )
    await message.answer(text)


# ------------------ АДМИНСКИЕ КОМАНДЫ (только для вас) ------------------

def admin_only(user_id):
    return user_id == ADMIN_ID


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
        text += (
            f"#{r['id']} | {r['end_date']} | {r['amount']}₽ | "
            f"{r['payment_status']} | владелец: {r['owner']}\n"
        )
    await message.answer(text)


@router.message(Command("paid"))
async def cmd_paid(message: Message):
    if not admin_only(message.from_user.id):
        return
    parts = message.text.split()
    if len(parts) != 2:
        await message.answer("Использование: /paid ID_аренды")
        return
    if mark_paid(parts[1]):
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
    else:
        await message.answer("Не найдено такой аренды.")


@router.message(Command("expense"))
async def cmd_expense(message: Message):
    if not admin_only(message.from_user.id):
        return
    parts = message.text.split(maxsplit=4)
    if len(parts) < 4:
        await message.answer(
            "Использование: /expense <владелец: Я/Денц> <сумма> <категория> <комментарий>\n"
            "Пример: /expense Я 1500 ремонт замена камеры на велике 3"
        )
        return
    _, owner, amount, category = parts[:4]
    comment = parts[4] if len(parts) > 4 else ""
    if owner not in ("Я", "Денц"):
        await message.answer("Владелец должен быть 'Я' или 'Денц'")
        return
    try:
        amount = int(amount)
    except ValueError:
        await message.answer("Сумма должна быть числом")
        return
    exp_id = add_expense(category, amount, owner, comment)
    await message.answer(f"Расход #{exp_id} записан: {owner}, {amount}₽, {category}")


@router.message(Command("report"))
async def cmd_report(message: Message):
    if not admin_only(message.from_user.id):
        return
    rentals = get_active_rentals()
    paid = [r for r in rentals if r.get("payment_status") == "оплачено"]
    income_me = sum(int(r["amount"]) for r in paid if r.get("owner") == "Я")
    income_friend = sum(int(r["amount"]) for r in paid if r.get("owner") == "Денц")

    finances = get_finance_rows()
    expenses = [r for r in finances if r.get("type") == "расход"]
    expense_me = sum(int(r["sum"]) for r in expenses if r.get("owner") == "Я")
    expense_friend = sum(int(r["sum"]) for r in expenses if r.get("owner") == "Денц")

    await message.answer(
        f"📊 Отчёт\n\n"
        f"— Доходы (оплаченные аренды) —\n"
        f"Моя доля: {income_me}₽\n"
        f"Доля Денца: {income_friend}₽\n\n"
        f"— Расходы —\n"
        f"Мои: {expense_me}₽\n"
        f"Денца: {expense_friend}₽\n\n"
        f"— Итого чистыми —\n"
        f"Я: {income_me - expense_me}₽\n"
        f"Денц: {income_friend - expense_friend}₽"
    )


# ------------------ НАПОМИНАНИЯ ------------------

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
        if days_left == 1:
            await bot.send_message(
                tg_id,
                f"⏰ Напоминание: завтра ({r['end_date']}) заканчивается срок аренды "
                f"(сумма {r['amount']}₽). Продлить — /rent, вопросы — напишите нам."
            )
        elif days_left == 0:
            await bot.send_message(
                tg_id,
                f"⚠️ Сегодня последний день аренды. Нужно вернуть велосипед "
                f"или продлить аренду, иначе оплата продолжит копиться."
            )
            await bot.send_message(
                ADMIN_ID,
                f"⚠️ Сегодня истекает аренда #{r['id']} (клиент id {r['client_id']})"
            )
        elif days_left < 0:
            await bot.send_message(
                ADMIN_ID,
                f"🔴 Просрочка! Аренда #{r['id']} истекла {r['end_date']} "
                f"(клиент id {r['client_id']}), оплата: {r['payment_status']}"
            )


async def main():
    scheduler = AsyncIOScheduler(timezone="Europe/Volgograd")
    scheduler.add_job(check_reminders, "cron", hour=10, minute=0)
    scheduler.start()
    log.info("Бот запущен")
    await dp.start_polling(bot)


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())