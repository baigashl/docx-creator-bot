import json
import os
from aiogram import types
from aiogram.dispatcher.filters import Text
from aiogram.utils import executor
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
import docx
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from config import dp, bot
from database.database import Order
from get_data import get_data_from_api, make_request
from docx.shared import Pt, Cm, Inches, Mm

button4 = KeyboardButton('🚚 Доставка')
markup3 = ReplyKeyboardMarkup(resize_keyboard=True).add(button4)


@dp.message_handler(commands=['start'])
async def process_start_command(message: types.Message):
    await bot.send_message(message.from_user.id, "👋🏼Добро пожаловать", reply_markup=markup3)


@dp.message_handler(Text(equals='🚚 Доставка'))
async def add_db(message: types.Message):
    id_user = message.from_user.id
    await bot.send_message(id_user, "Немного подождите")
    await get_data_from_api()
    with open('delivery.json') as json_file:
        delivery_data = json.load(json_file)

    # print(delivery_data[0]["orders"][0]["orderCode"])
    for order in delivery_data[0]["orders"]:
        code = Order.get_or_none(orderCode=order["orderCode"])
        if code is None:

            Order.create(
                orderCode=order["orderCode"],
            )
            with open(f'{order["orderCode"]}_prod_detail.json') as json_file:
                detail_data = json.load(json_file)


            # print(order["orderCode"])
            doc = docx.Document()
            sections = doc.sections
            section = sections[0]
            section.page_height = Mm(297)
            section.page_width = Mm(210)
            section.header_distance = Mm(12.7)
            section.footer_distance = Mm(12.7)
            for section in sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(3)
                section.right_margin = Inches(3)

            paragraph = doc.add_paragraph(f'Номер заказа\n')
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


            font_styles = doc.styles
            font_charstyle = font_styles.add_style('CommentsStyle', WD_STYLE_TYPE.CHARACTER)
            font_object = font_charstyle.font
            font_object.size = Pt(9)
            font_object.name = 'Helvetica'

            paragraph.add_run(f'{detail_data["orderId"]}\n', style='CommentsStyle')
            paragraph.add_run(f'Сумма\n', style='CommentsStyle').bold = True
            paragraph.add_run(f'{detail_data["localizedSum"]}\n', style='CommentsStyle')
            paragraph.add_run('\n', style='CommentsStyle')
            paragraph.add_run(f'Стоимость доставки для клиента\n', style='CommentsStyle')
            paragraph.add_run(f'{detail_data["deliveryDiscountFormatted"]}\n', style='CommentsStyle')
            paragraph.add_run('\n', style='CommentsStyle')
            paragraph.add_run(f'Имя\n', style='CommentsStyle')
            paragraph.add_run(f'{detail_data["purchaserFirstName"]} ', style='CommentsStyle')
            paragraph.add_run(f'{detail_data["purchaserLastName"]}\n', style='CommentsStyle')
            paragraph.add_run(f'Номер телефона\n', style='CommentsStyle')
            paragraph.add_run(f'+7 ({detail_data["purchaserPhoneNumber"][0:3]})-{detail_data["purchaserPhoneNumber"][3:6]}-{detail_data["purchaserPhoneNumber"][6:8]}-{detail_data["purchaserPhoneNumber"][8:10]}\n', style='CommentsStyle')
            paragraph.add_run('\n', style='CommentsStyle')
            #################################################################3
            for product in detail_data["products"]:
                paragraph.add_run(f'{product["name"]}\n', style='CommentsStyle')
                paragraph.add_run(f'{product["quantity"]}                 ', style='CommentsStyle')
                paragraph.add_run(f'{product["localizedActualPrice"]}\n', style='CommentsStyle')
                paragraph.add_run('\n', style='CommentsStyle')
            ########################################################
            # paragraph.add_run(f'{order["entries"][0]["name"]}\n', style='CommentsStyle')
            # paragraph.add_run(f'{order["entries"][0]["quantity"]}                 ', style='CommentsStyle')
            # paragraph.add_run(f'{order["entries"][0]["totalPrice"]}\n', style='CommentsStyle')
            ###############################################################################3
            paragraph.add_run(f'Адрес доставки\n', style='CommentsStyle').bold = True
            paragraph.add_run('\n', style='CommentsStyle')
            paragraph.add_run(f'{detail_data["deliveryAddress"]["formattedAddress"]}\n', style='CommentsStyle')
            # doc.center()

            doc.save(f'{order["orderCode"]}example.docx')
            # bot.send_file(f'{order["orderCode"]}example.docx')
            await bot.send_document(id_user, document=open(f'{order["orderCode"]}example.docx', 'rb'))
            os.remove(f'{order["orderCode"]}example.docx')
            os.remove(f'{order["orderCode"]}_prod_detail.json')
        else:
            os.remove(f'{order["orderCode"]}_prod_detail.json')
            print('exist')
    # await bot.send_file()
    await bot.send_message(id_user, "Все данные были добавлены Базу данных")

if __name__ == '__main__':
    executor.start_polling(dp)