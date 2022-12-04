from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram import Bot, Dispatcher
from decouple import config

TOKEN = config('TOKEN')
host = config('POSTGRES_HOST')
user = config('POSTGRES_USER')
password = config('POSTGRES_PASSWORD')
db_name = config('POSTGRES_DB')

storage = MemoryStorage()
bot = Bot(token=TOKEN)
dp = Dispatcher(bot, storage=storage)

PG_URL = f'postgresql: //{user}:{password}@{host}/{db_name}'
