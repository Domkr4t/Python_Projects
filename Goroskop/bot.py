from aiogram import Bot, Dispatcher, executor, types
from aiogram.types.reply_keyboard import ReplyKeyboardMarkup
from aiogram.dispatcher.filters import Text
from main import result_today
from main import result_yesterday
from main import result_tomorrow
from main import result_week
from main import result_year
import os


bot = Bot(token="bot_token")
dp = Dispatcher(bot)

@dp.message_handler(commands="start")
async def start(message: types.Message):
    start_buttons = [
        'Гороскоп на вчера',
        'Гороскоп на сегодня',
        'Гороскоп на завтра',
        'Гороскоп на неделю',
        'Гороскоп на год'
    ]
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.add(*start_buttons)

    await message.answer("Гороскопы", reply_markup=keyboard)



#ГОРОСКОП НА СЕГОДНЯ
@dp.message_handler(Text(equals='Гороскоп на сегодня'))
async def get_goroskop(message: types.Message):
    await message.answer("Секундочку...")

    result_today()

    with open('result.txt', encoding = "utf-8") as file:
        data = file.read()

    await message.answer(data)

    os.remove('result.txt')


#ГОРОСКОП НА ВЧЕРА
@dp.message_handler(Text(equals='Гороскоп на вчера'))
async def get_goroskop(message: types.Message):
    await message.answer("Секундочку...")

    result_yesterday()

    with open('result.txt', encoding = "utf-8") as file:
        data = file.read()

    await message.answer(data)

    os.remove('result.txt')

#ГОРОСКОП НА ЗАВТРА
@dp.message_handler(Text(equals='Гороскоп на завтра'))
async def get_goroskop(message: types.Message):
    await message.answer("Секундочку...")

    result_tomorrow()

    with open('result.txt', encoding = "utf-8") as file:
        data = file.read()

    await message.answer(data)

    os.remove('result.txt')

#ГОРОСКОП НА НЕДЕЛЮ
@dp.message_handler(Text(equals='Гороскоп на неделю'))
async def get_goroskop(message: types.Message):
    await message.answer("Секундочку...")

    result_week()

    with open('result.txt', encoding = "utf-8") as file:
        data = file.read()

    await message.answer(data)

    os.remove('result.txt')

#ГОРОСКОП НА ГОД
@dp.message_handler(Text(equals='Гороскоп на год'))
async def get_goroskop(message: types.Message):
    await message.answer("Секундочку...")

    result_year()

    with open('result.txt', encoding = "utf-8") as file:
        data = file.read()

    await message.answer(data)

    os.remove('result.txt')

def main():
    executor.start_polling(dp)

if __name__ == "__main__":
    main()
