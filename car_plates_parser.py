from bs4 import BeautifulSoup
import pandas as pd
import requests
import urllib3
import logging
import os
import pandas as pd
import xlsxwriter
from aiogram import Bot, Dispatcher, types
from aiogram.dispatcher.filters import Command
from aiogram.dispatcher.filters import Text
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import ParseMode
from aiogram.utils import executor
from aiogram.types import CallbackQuery
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.types import CallbackQuery
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup
import datetime

from auth_tg import token

def html_receive(region):

    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        # 'Cookie': '_ga=GA1.3.1940355900.1677262208; _gid=GA1.3.509915868.1677694972; _gat=1',
        'Origin': 'https://opendata.hsc.gov.ua',
        'Referer': 'https://opendata.hsc.gov.ua/check-leisure-license-plates/',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Chromium";v="110", "Not A(Brand";v="24", "Google Chrome";v="110"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    url = 'http://opendata.hsc.gov.ua/check-leisure-license-plates'
    csrf_response = requests.post(url, verify=False, timeout=10)
    csrf_soup = BeautifulSoup(csrf_response.text, 'html.parser')
    csrf_token = csrf_soup.find(
        'input', {'name': 'csrfmiddlewaretoken'})['value']

    data = {
        'region': region,
        'type_venichle': 'light_car_and_truck',
        'number': '',
        'csrfmiddlewaretoken': csrf_token,
    }

    response = requests.post(
        'http://opendata.hsc.gov.ua/check-leisure-license-plates/', headers=headers, data=data)

    # здесь вместо "html_code" нужно указать HTML код страницы
    decoded = response.content
    decoded_utf = decoded.decode('utf-8')

    return decoded_utf

def handler(region, name_of_tsc, price):
    datetime_start = datetime.datetime.now()
    decoded_utf = html_receive(region)
    soup = BeautifulSoup(decoded_utf, 'html.parser')

    # получаем таблицу
    table = soup.find('table')

    # создаем пустой список, в который будем добавлять данные
    data = []

    # проходимся по каждой строке таблицы
    try:
        for row in table.find_all('tr'):
            # получаем значения ячеек
            columns = row.find_all('td')
            # добавляем данные в список
            data.append([col.text.strip() for col in columns])

    except:
        data.append('Nothing')

    # создаем DataFrame с полученными данными
    df = pd.DataFrame(data[1:], columns=data[0])

    # фильтруем данные
    filtered_df = df[(df.iloc[:, 1] == f'{price}') & (
        df.iloc[:, 2].str.contains(f'{name_of_tsc}'))]

    # сортируем по первому столбцу в порядке от А до Я
    filtered_df = filtered_df.sort_values(
        by=filtered_df.columns[0], ascending=True)

    # сохраняем данные в Excel
    with pd.ExcelWriter(f'{name_of_tsc}.xlsx', engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # задаем ширину первого столбца в соответствии с его содержимым
        worksheet.set_column(0, 0, max(df.iloc[:, 0].str.len()) + 1)

    data_len = filtered_df.shape[0]

    # выводим количество номеров в первом столбце
    datetime_now = datetime.datetime.now()
    used_time = datetime_now - datetime_start
    print(
        f'Количество номеров: {data_len}, запрос на {name_of_tsc}, {region} регион, {datetime_now}, использовано {used_time}сек')

    return data_len

# TG Token
bot = Bot(token)
# устанавливаем уровень логирования
logging.basicConfig(level=logging.INFO)

# инициализируем бота и диспетчера
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)


# создаем состояние для региона
class Region(StatesGroup):
    region = State()

# создаем состояние для ТСЦ


class TSC(StatesGroup):
    tsc = State()


class PRICE(StatesGroup):
    price = State()


markup = InlineKeyboardMarkup().add(InlineKeyboardButton(
    "Перевірити наявність номерних знаків ✅", callback_data="start"))
markup_reload = InlineKeyboardMarkup().add(
    InlineKeyboardButton("Спробувати ще раз 🔂", callback_data="start"))


@dp.message_handler(commands=['start'])
async def cmd_start(message: types.Message):
    # выводим сообщение "Нажмите кнопку "начать""
    await message.answer('Привіт👋🏻 Я бот, який допоможе перевірити наявність безкоштовних номерів у твоєму сервісному центрі! Нажми на кнопку нижче 👇🏻', reply_markup=markup)


@dp.callback_query_handler(text="start")
async def process_start(callback_query: CallbackQuery):
    # выводим сообщение "Выберите регион"
    await bot.send_message(callback_query.from_user.id, '''😌 Оберіть регіон:

1 - АР Крим
2 - Вінницька
3 - Волинська
4 - Дніпропетровська
5 - Донецька
6 - Житомирська
7 - Закарпатська
8 - Запорізька
9 - Івано-Франківська
26 - м. Київ
10 - Київська
11 - Кіровоградська
12 - Луганська
13 - Львівська
14 - Миколаївська
15 - Одеська
16 - Полтавська
17 - Рівненська
18 - Сумська
19 - Тернопільська
20 - Харківська
21 - Херсонська
22 - Хмельницька
23 - Черкаська
25 - Чернігівська
24 - Чернівецька

❗️Надішліть число регіону, наприклад: 26''')

    # переходим в состояние Region.region
    await Region.region.set()


# обработчик сообщения со значением региона
@dp.message_handler(state=Region.region)
async def process_region(message: types.Message, state: FSMContext):

    # сохраняем значение региона в переменную
    region = message.text

    if not region.isdigit() or int(region) not in range(1,27):
        await message.answer(f"😢 Мабуть Ви обрали неправильний регіон. \n Спробуйте обрати ще раз!\n\n ❗️Надішліть число регіону, наприклад: 26",reply_markup=markup_reload)
        return

    # выводим сообщение "Напишите название или номер ТСЦ"
    await message.answer(f"Напишіть назву чи номер ТСЦ. \nЩоб дізнатись номер вашого ТСЦ, перейдіть за <a href=\"https://hsc.gov.ua/kontakti/kontakti-gsts-pidrozdiliv/\">посиланням</a>.\n❗️Надішліть номер ТСЦ, наприклад: 8046", parse_mode="HTML", disable_web_page_preview=True)

    # сохраняем значение региона в состояние Region.region
    await state.update_data(region=region)

    # переходим в состояние TSC.tsc
    await TSC.tsc.set()

# обработчик сообщения со значением названия ТСЦ


@dp.message_handler(state=TSC.tsc)
async def process_tsc(message: types.Message, state: FSMContext):
    # сохраняем значение названия ТСЦ в переменную
    tsc = message.text

    # получаем значение региона из состояния
    data = await state.get_data()
    region = data.get('region')

    # выводим сообщение "Укажите стоимость"
    await message.answer(
"""Оберiть вартiсть номерного знаку💸:

Комбінації чотирьох однакових цифр — 30 тис. грн;
Послідовні комбінації від 0001 до 0009 — 30 тис. грн;

Три однакові цифри поспіль (наприклад: 0777) — 10 тис. грн;
Дві послідовні пари однакових цифр (наприклад: 1122) — 10 тис. грн;
Комбінації 0123 та 1234 — 10 тис. грн;

Комбінації, що починаються з двох нулів, а наступні цифри неоднакові (наприклад: 0032) — 8 тис. грн;

Три однакові цифри не поспіль (наприклад: 7077) — 4 тис. грн;
Комбінації, де однакові перша і третя цифра, друга і четверта цифра; або перша й четверта цифра, друга і третя цифра (наприклад: 1221; 1212) — 4 тис. грн.""", reply_markup=InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="0 UAH", callback_data="0"),
             InlineKeyboardButton(text="4000 UAH", callback_data="4000"),
             InlineKeyboardButton(text="8000 UAH", callback_data="8000"),
             InlineKeyboardButton(text="10000 UAH", callback_data="10000"),
             InlineKeyboardButton(text="30000 UAH", callback_data="30000"),]
        ]
    ))

    # сохраняем значение названия ТСЦ и региона в состоянии
    await state.update_data(tsc=tsc, region=region)

    # переходим в состояние Cost.cost
    await PRICE.price.set()

# обработчик нажатия на кнопку со стоимостью


@dp.callback_query_handler(state=PRICE.price)
async def process_cost(callback_query: CallbackQuery, state: FSMContext):
    # сохраняем значение стоимости в переменную
    cost = callback_query.data

    # получаем значение названия ТСЦ и региона из состояния
    data = await state.get_data()
    tsc = data.get('tsc')
    region = data.get('region')

    # выводим сообщение с выбранными параметрами
    data_len = None
    try:
        # ТУТ ХЕНДЛЕР
        # ТУТ ХЕНДЛЕР
        # ТУТ ХЕНДЛЕР
        in_process_message = await bot.send_message(callback_query.from_user.id, '⏳ Перевіряю наявність…')
        data_len = handler(region=region, name_of_tsc=tsc, price=cost)
        await bot.send_message(callback_query.from_user.id, f"Обраний регіон: {region}\nОбраний ТСЦ: {tsc}.\nВсього вільних номерів: {data_len}")

        file_name = f"{tsc}.xlsx"  # имя файла
        # полный путь до файла
        file_path = os.path.join(os.getcwd(), file_name)

        await in_process_message.delete()
        with open(file_path, "rb") as file:
            # отправляем файл в чат
            await bot.send_document(callback_query.from_user.id, file, reply_markup=markup_reload)
            os.remove(file_name)
        # завершаем состояние и очищаем данные
        await state.finish()

    except:
        if data_len is not None:
            await bot.send_message(callback_query.from_user.id, f"Обраний регіон: {region}\nОбраний ТСЦ: {tsc}.\nВсього вільних номерів: {data_len}, але сталася помилка. Спробуйте ще раз!")
        else:
            await bot.send_message(callback_query.from_user.id, f"😢 Сталася помилка. Спробуйте ще раз 👇🏻", reply_markup=markup_reload)
        await state.finish()


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
