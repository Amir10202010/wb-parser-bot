import requests
import openpyxl
import time
import re
from background import keep_alive
from aiogram import Bot, Dispatcher
from aiogram.types import FSInputFile
from aiogram.filters import Command
from aiogram.types import Message
import asyncio
import os

# Ваш токен бота
API_TOKEN = os.environ.get('TOKEN')

# Инициализация бота и диспетчера
bot = Bot(token=API_TOKEN)
dp = Dispatcher()

# Функция для запроса данных по URL с повторными попытками
def fetch_data(url, retries=3, delay=2):
    for attempt in range(retries):
        try:
            response = requests.get(url)
            if response.status_code == 200:
                try:
                    return response.json()
                except requests.exceptions.JSONDecodeError:
                    print(f"Ошибка декодирования JSON по адресу: {url}")
                    return None
            else:
                print(f"Ошибка запроса. Код статуса: {response.status_code}. URL: {url}")
        except requests.RequestException as e:
            print(f"Ошибка соединения: {e}")
        
        print(f"Попытка {attempt + 1} не удалась. Повтор через {delay} секунд...")
        time.sleep(delay)
    
    print(f"Все попытки запроса к {url} исчерпаны.")
    return None

# Функция для парсинга продуктов из JSON
def parse_products(data):
    products = []
    if data and 'data' in data:
        for item in data['data']['products']:
            name = item.get('name', 'N/A')
            article = item.get('id', 'N/A')
            brand = item.get('brand', 'N/A')
            isbn = item.get('isbn', 'N/A')
            price = item.get('sizes', [{}])[0].get('price', {}).get('total', 0) / 100 // 1.0309945504
            products.append([name, article, brand, isbn, price])
    return products

# Функция для парсинга всех страниц последовательно
def fetch_all_pages(base_url):
    all_products = []
    page = 1
    
    while True:
        url = f"{base_url}&page={page}"
        data = fetch_data(url)  # Используем функцию с retry
        if not data or 'data' not in data or not data['data']['products']:
            break

        products = parse_products(data)
        all_products.extend(products)
        print(f"Обработана страница {page} с {len(products)} продуктами.")
        page += 1

    return all_products

# Функция для записи результатов в Excel
def save_to_excel(products, file_name):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Название", "Артикул", "Бренд", "ISBN", "Цена"])

    for product in products:
        sheet.append(product)

    workbook.save(file_name)

# Команда /start
@dp.message(Command("start"))
async def start_command(message: Message):
    await message.answer(
        "Привет! Я бот для парсинга данных с Wildberries.\n"
        "Отправь мне ссылку на продавца или бренд, и я соберу данные для тебя.\n"
        "Пример ссылки на продавца: https://www.wildberries.ru/seller/8969\n"
        "Пример ссылки на бренд: https://www.wildberries.ru/brands/eksmo/\n"
        "Используй команду /parse <ссылка>, чтобы начать парсинг."
    )

# Команда для обработки парсинга
@dp.message(Command("parse"))
async def parse_wb(message: Message):
    if 'https://www.wildberries.ru' not in message.text:
        await message.answer("Укажите нормальную ссылку на бренд или селлера")
    else:
        url = message.text.split(' ')[1]
        start_time = time.time()

        if "seller" in url:
            seller_id = url.split('/')[-1]
            base_url = f"https://catalog.wb.ru/sellers/v2/catalog?ab_testing=false&appType=1&curr=rub&dest=82&sort=priceup&spp=30&supplier={seller_id}"
            file_name = f"seller_{seller_id}.xlsx"
        elif "brands" in url:
            brand_name = re.search(r'brands/(.+?)/', url)
            if brand_name:
                brand_name = brand_name.group(1)
                # Запрос ID бренда
                brand_data = fetch_data(f'https://static-basket-01.wbbasket.ru/vol0/data/brands/{brand_name}.json')
                brand_id = brand_data['id'] if brand_data else None
                if brand_id:
                    base_url = f"https://catalog.wb.ru/brands/v2/catalog?ab_testing=false&appType=1&brand={brand_id}&curr=rub&dest=82&sort=priceup&spp=30"
                    file_name = f"brand_{brand_name}.xlsx"
                else:
                    await message.answer("Не удалось получить ID бренда.")
                    return
            else:
                await message.answer("Неправильная ссылка на бренд.")
                return
        else:
            await message.answer("Неправильная ссылка.")
            return

        await message.answer("Начало парсинга, подождите...")
        # Парсинг всех продуктов и запись в Excel
        products = fetch_all_pages(base_url)
        save_to_excel(products, file_name)

        end_time = time.time()
        elapsed_time = end_time - start_time

        # Отправка Excel файла пользователю
        file_input = FSInputFile(file_name)
        await bot.send_document(message.chat.id, file_input)
        os.remove(file_name)
        print(f"Файл {file_name} удален.")
        await message.answer(f"Парсинг завершен за {elapsed_time:.2f} секунд. Данные сохранены в {file_name}")

# Запуск бота
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    try:
        print('Запуск Бота')
        keep_alive()
        asyncio.run(main())
    except KeyboardInterrupt:
        print('Завершение Бота')
