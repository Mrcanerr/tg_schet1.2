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

waiting_for_text = {}
miss_data = {}



btn_upload = KeyboardButton("Загрузить файл")
btn_report = KeyboardButton("Получить отчет")
btn_format = KeyboardButton("Я еблан")
btn_miss = KeyboardButton("Показать несовпадения")

markup = ReplyKeyboardMarkup(resize_keyboard=True)
markup.add(btn_upload)
markup.add(btn_report)
markup.add(btn_format)
markup.add(btn_miss)



@dp.message_handler(commands=["start"])
async def start(message: types.Message):
    await message.answer(
        "Привет. Я бот учета.\nВыбери действие:",
        reply_markup=markup
    )



@dp.message_handler(lambda message: message.text == "Я ебланище")
async def how_file(message: types.Message):
    await message.answer(
        "Excel файл должен содержать:\n\n"
        "Колонка A — Название товара\n"
        "Колонка B — Количество (план)\n\n"
        "Без пустых строк между товарами."
    )



@dp.message_handler(lambda message: message.text == "Загрузить файл")
async def upload_request(message: types.Message):
    await message.answer("Отправьте Excel файл (.xlsx)")

@dp.message_handler(content_types=types.ContentType.DOCUMENT)
async def handle_file(message: types.Message):
    if not message.document.file_name.endswith(".xlsx"):
        await message.answer("Нужен файл формата .xlsx")
        return

    file = await bot.get_file(message.document.file_id)
    await bot.download_file(file.file_path, "tabl0.xlsx")

    waiting_for_text[message.from_user.id] = True

    await message.answer("Файл получен.\nТеперь отправь фактические остатки списком.")



@dp.message_handler(content_types=types.ContentType.TEXT)
async def handle_inventory_text(message: types.Message):

    if message.from_user.id not in waiting_for_text:
        return

    data = parse_inventory(message.text)

    wb = load_workbook("tabl0.xlsx")
    ws = wb.active

    matches = 0
    misses = []

    table_products = []

    for row in range(2, ws.max_row + 1):
        product_name = str(ws[f"A{row}"].value).lower()
        table_products.append(product_name)

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
            matches += 1
        else:
            misses.append(name)

    wb.save("report.xlsx")

    miss_data[message.from_user.id] = misses
    waiting_for_text.pop(message.from_user.id)

    await message.answer(
        f"Готово.\n"
        f"Совпадений найдено: {matches}\n"
        f"Несовпадений: {len(misses)}\n\n"
        f"Нажмите 'Получить отчет'"
    )



@dp.message_handler(lambda message: message.text == "Получить отчет")
async def send_report(message: types.Message):
    try:
        file = InputFile("report.xlsx")
        await message.answer_document(file)
    except:
        await message.answer("Сначала загрузи файл и отправьте остатки.")



@dp.message_handler(lambda message: message.text == "Показать несовпадения")
async def show_miss(message: types.Message):

    misses = miss_data.get(message.from_user.id)

    if not misses:
        await message.answer("Несовпадений нет.")
        return

    text = "Не найдено в таблице:\n\n"
    text += "\n".join(misses)

    await message.answer(text)



if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)
