from aiogram.dispatcher import FSMContext
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher.filters.state import StatesGroup, State
import logging
from aiogram import Bot, Dispatcher, executor, types
import cfg as c
from docx import Document
from docx_ed import docer as dc, file_reader as fl

bot = Bot(token=c.TOKEN)
logging.basicConfig(level=logging.INFO)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

m_i_key = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
m_i_key.add("1.0", "1.25", "1.5", "2")

alig_key = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
alig_key.add("по ширине", "по левому краю", "по правому краю", "по центру", "по умолчанию")

start_keys = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True, row_width=1)
start_keys.add("Проверить конкретный гост", "Проверить отдельно", "Прекратить взаимодействие")
other_keys = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
other_keys.add("Проверка выравнивания", "Проверка межстрочного интервала", "Проверка абзацных отступов",
               "Прекратить взаимодействие")

a_i_key = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
a_i_key.add("1.0", "1.25", "1.5", "2", "3")


def gost_keys():
    gost_keys = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    for ke in fl.file_reader.get_files():
        gost_keys.add(ke)
    return gost_keys


class Form(StatesGroup):
    file_ = State()
    prechoose = State()
    gost = State()
    choose1 = State()
    choose2 = State()


@dp.message_handler(commands=["start"])
async def aboba(message):
    await dp.bot.send_message(message.chat.id, f"Загружайте файл  для проверки на соответствие госту")
    await Form.file_.set()


@dp.message_handler(state=Form.file_, content_types=types.ContentType.DOCUMENT)
async def handle_docs(message: types.Message, state: FSMContext):
    if message.document.mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
        doc_name = await message.document.download(destination_dir="files/")
        async with state.proxy() as data:
            data['doc_obj'] = dc.FileManger(user_id=message.chat.id,
                                            docx_=Document(doc_name.name),
                                            name=doc_name.name)
        await message.answer(
            "Спасибо, ваш файл docx получен и обработан! Теперь отправьте, что вы хотите проверить",
            reply_markup=start_keys)
        await Form.prechoose.set()
    else:
        await message.answer("Пожалуйста, отправьте файл в формате docx.")


@dp.message_handler(state=Form.prechoose)
async def process(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        if "Проверить конкретный гост" == message.text:
            await message.answer(
                "Спасибо, за выбор теперь выберите гост на который вы хотите проверить",
                reply_markup=gost_keys())
            await Form.gost.set()
        elif "Проверить отдельно" == message.text:
            await message.answer(
                "Теперь отправьте, что вы хотите проверить",
                reply_markup=other_keys)
            await Form.choose1.set()

        elif "Прекратить взаимодействие" == message.text:
            await state.finish()
            data['doc_obj'].close()
            return


@dp.message_handler(state=Form.gost)
async def process_gost(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        if message.text in fl.file_reader.get_files():
            data['doc_obj'].gost = message.text
            respon = data['doc_obj'].full_check()
            await message.answer(respon, reply_markup=None)
            await Form.prechoose.set()
            await message.answer(
                "Спасибо, за использование нашего бота, вы можете выбрать другую функцию для вашего файла",
                reply_markup=start_keys)

        else:
            await message.answer('Данного госта нет в базе', reply_markup=gost_keys())

@dp.message_handler(state=Form.choose1)
async def process1(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        if "Проверка выравнивания" == message.text:
            data['doc_obj'].rej = 0
            await message.answer(
                "Спасибо, за выбор теперь выберите требование",
                reply_markup=alig_key)
        elif "Проверка межстрочного интервала" == message.text:
            data['doc_obj'].rej = 1
            await message.answer(
                "Спасибо, за выбор теперь выберите требование",
                reply_markup=m_i_key)
        elif "Проверка абзацных отступов" == message.text:
            data['doc_obj'].rej = 2
            await message.answer(
                "Спасибо, за выбор теперь выберите требование",
                reply_markup=a_i_key)

        elif "Прекратить взаимодействие" == message.text:
            await state.finish()
            data['doc_obj'].close()
            return
        await Form.choose2.set()


@dp.message_handler(state=Form.choose2)
async def process2(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        if data['doc_obj'].rej == 0:
            data['doc_obj'].alignment = message.text
        elif data['doc_obj'].rej == 1:
            data['doc_obj'].interval = float(message.text)
        elif data['doc_obj'].rej == 2:
            data['doc_obj'].indent = float(message.text)
    await message.answer(data['doc_obj'].checker(), reply_markup=None)
    await Form.choose1.set()
    await message.answer(
        "Спасибо, за использование нашего бота, вы можете выбрать другую функцию для вашего файла",
        reply_markup=start_keys)


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
