import logging
from aiogram import Bot, Dispatcher, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InputFile
from aiogram.utils import executor
from openpyxl import load_workbook
from parser import parse_inventory, find_similar

TOKEN = "8464230833:AAHuVdH301Oh2vNEplUpYPHlWLYtlQEBZzk"

logging.basicConfig(level=logging.INFO)

bot = Bot(token=TOKEN)
dp = Dispatcher(bot)

state_users = {}
miss_data = {}

-

markup = ReplyKeyboardMarkup(resize_keyboard=True)
markup.add("Загрузить таблицу программы")
markup.add("Получить отчет")
markup.add("Показать несовпадения")



@dp.message_handler(commands=["start"])
async def start(message: types.Message):
    await message.answer(
        "Привет. Я бот учета.\nВыбери действие:",
        reply_markup=markup
    )



@dp.message_handler(lambda m: m.text == "Загрузить таблицу программы")
async def request_file(message: types.Message):
    state_users[message.from_user.id] = "waiting_program"
    await message.answer("Отправьте таблицу из программы (xlsx)")



@dp.message_handler(content_types=types.ContentType.DOCUMENT)
async def handle_excel(message: types.Message):

    user_state = state_users.get(message.from_user.id)

    if user_state != "waiting_program":
        return

    file = await bot.get_file(message.document.file_id)
    await bot.download_file(file.file_path, "program.xlsx")

    base_wb = load_workbook("tabl0.xlsx")
    base_ws = base_wb.active

    prog_wb = load_workbook("program.xlsx")
    prog_ws = prog_wb.active

    program_data = {}

    for row in range(2, prog_ws.max_row + 1):
        name = str(prog_ws[f"A{row}"].value).lower()
        qty = prog_ws[f"B{row}"].value or 0
        program_data[name] = qty

    for row in range(2, base_ws.max_row + 1):
        product_name = str(base_ws[f"A{row}"].value).lower()
        if product_name in program_data:
            base_ws[f"B{row}"] = program_data[product_name]

    base_wb.save("work.xlsx")

    state_users[message.from_user.id] = "waiting_fact"

    await message.answer("Таблица принята.\nТеперь отправь фактические остатки списком.")



@dp.message_handler(lambda message: state_users.get(message.from_user.id) == "waiting_fact")
async def handle_fact(message: types.Message):

    data = parse_inventory(message.text)

    wb = load_workbook("work.xlsx")
    ws = wb.active

    table_products = []

    for row in range(2, ws.max_row + 1):
        table_products.append(str(ws[f"A{row}"].value).lower())

    misses = []

    for name, qty in data.items():

        found_row = None

        for row in range(2, ws.max_row + 1):
            product_name = str(ws[f"A{row}"].value).lower()
            if name == product_name:
                found_row = row
                break

        if not found_row:
            similar = find_similar(name, table_products)
            if similar:
                for row in range(2, ws.max_row + 1):
                    if str(ws[f"A{row}"].value).lower() == similar:
                        found_row = row
                        break

        if found_row:
            ws[f"C{found_row}"] = qty
            plan = ws[f"B{found_row}"].value or 0
            ws[f"D{found_row}"] = plan - qty
        else:
            misses.append(name)

    wb.save("report.xlsx")

    miss_data[message.from_user.id] = misses
    state_users[message.from_user.id] = None

    await message.answer("Факт принят. Нажмите 'Получить отчет'")



@dp.message_handler(lambda message: message.text == "Получить отчет")
async def send_report(message: types.Message):
    try:
        await message.answer_document(InputFile("report.xlsx"))
    except:
        await message.answer("Сначала загрузите таблицу.")



@dp.message_handler(lambda message: message.text == "Показать несовпадения")
async def show_miss(message: types.Message):

    misses = miss_data.get(message.from_user.id)

    if not misses:
        await message.answer("Несовпадений нет.")
        return

    text = "Не найдено:\n\n" + "\n".join(misses)
    await message.answer(text)



if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)
