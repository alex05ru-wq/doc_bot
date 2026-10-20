"""
Telegram-бот, который задаёт вопросы и по ответам
готовит документ в формате Word (.docx) и PDF.

Запуск: python bot.py   (перед этим заполните .env)
"""
import os
from html import escape
import threading
from datetime import datetime
from pathlib import Path

import telebot
from dotenv import load_dotenv
from telebot import types

from config import BUSINESS_NAME, TEMPLATES
from doc_maker import make_documents

load_dotenv()
TOKEN = os.getenv('BOT_TOKEN')
ADMIN_ID = os.getenv('ADMIN_ID')          # необязательно: кому присылать копии
if not TOKEN:
    raise SystemExit('Не найден BOT_TOKEN. Создайте файл .env (см. .env.example)')

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / 'output'
CANCEL = '❌ Отмена'

bot = telebot.TeleBot(TOKEN)

# Кто сейчас какой документ заполняет: user_id -> {'tpl', 'step', 'answers'}
sessions = {}
lock = threading.Lock()


# ---------- Клавиатуры ----------

def templates_keyboard():
    kb = types.InlineKeyboardMarkup()
    for key, tpl in TEMPLATES.items():
        kb.add(types.InlineKeyboardButton(tpl['title'], callback_data=f'tpl:{key}'))
    return kb


def cancel_keyboard():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add(CANCEL)
    return kb


def confirm_keyboard():
    kb = types.InlineKeyboardMarkup()
    kb.add(types.InlineKeyboardButton('✅ Сформировать документ', callback_data='make'))
    kb.add(types.InlineKeyboardButton('🔄 Заполнить заново', callback_data='restart'),
           types.InlineKeyboardButton('❌ Отмена', callback_data='cancel'))
    return kb


# ---------- Проверка ответов ----------

def check_answer(kind, text):
    """Возвращает (значение, None) или (None, текст ошибки)."""
    text = text.strip()
    if not text:
        return None, 'Ответ не может быть пустым.'

    if kind == 'number':
        digits = text.replace(' ', '').replace(',', '.')
        try:
            number = float(digits)
        except ValueError:
            return None, 'Нужно число, например: 15000'
        if number <= 0:
            return None, 'Число должно быть больше нуля.'
        if number == int(number):
            return f'{int(number):,}'.replace(',', ' '), None
        return f'{number:,.2f}'.replace(',', ' ').replace('.', ','), None

    if kind == 'date':
        for fmt in ('%d.%m.%Y', '%d-%m-%Y', '%d/%m/%Y', '%d.%m.%y'):
            try:
                return datetime.strptime(text, fmt).strftime('%d.%m.%Y'), None
            except ValueError:
                pass
        return None, 'Нужна дата в формате ДД.ММ.ГГГГ, например: 25.10.2026'

    if kind == 'phone':
        digits = ''.join(ch for ch in text if ch.isdigit())
        if len(digits) not in (10, 11):
            return None, 'Похоже, в номере ошибка. Пример: +7 900 123-45-67'
        return text, None

    if len(text) > 300:
        return None, 'Слишком длинно, сократите до 300 символов.'
    return text, None


# ---------- Шаги диалога ----------

def show_menu(chat_id, text=None):
    bot.send_message(chat_id, text or 'Какой документ подготовить?',
                     reply_markup=templates_keyboard())


def ask_question(chat_id, user_id):
    with lock:
        s = sessions[user_id]
        tpl = TEMPLATES[s['tpl']]
        step = s['step']
    total = len(tpl['fields'])
    _, question, _ = tpl['fields'][step]
    bot.send_message(chat_id, f'Вопрос {step + 1} из {total}\n\n{question}',
                     reply_markup=cancel_keyboard())


def show_summary(chat_id, user_id):
    with lock:
        s = sessions[user_id]
        tpl = TEMPLATES[s['tpl']]
        answers = dict(s['answers'])
    lines = [f'<b>{tpl["title"]}</b>', '', 'Проверьте данные:']
    for key, question, _ in tpl['fields']:
        label = question.split('(')[0].rstrip(' :?')
        lines.append(f'• {label}: <b>{escape(answers[key])}</b>')
    bot.send_message(chat_id, '✅ Все ответы получены', reply_markup=types.ReplyKeyboardRemove())
    bot.send_message(chat_id, '\n'.join(lines), parse_mode='HTML',
                     reply_markup=confirm_keyboard())


def start_template(chat_id, user_id, key):
    with lock:
        sessions[user_id] = {'tpl': key, 'step': 0, 'answers': {}}
    bot.send_message(chat_id, f'Заполняем: {TEMPLATES[key]["title"]}\n'
                              f'Отвечайте на вопросы по очереди. '
                              f'Прервать можно кнопкой «{CANCEL}».')
    ask_question(chat_id, user_id)


# ---------- Обработчики ----------

@bot.message_handler(commands=['start'])
def cmd_start(message):
    with lock:
        sessions.pop(message.from_user.id, None)
    bot.send_message(message.chat.id,
                     f'Здравствуйте! Это {BUSINESS_NAME}.\n\n'
                     'Я задам несколько вопросов и подготовлю готовый документ '
                     'в форматах Word и PDF.',
                     reply_markup=types.ReplyKeyboardRemove())
    show_menu(message.chat.id)


@bot.message_handler(commands=['help'])
def cmd_help(message):
    bot.send_message(message.chat.id,
                     '/start — выбрать документ\n'
                     'Во время заполнения можно нажать «Отмена».')


@bot.message_handler(commands=['myid'])
def cmd_myid(message):
    bot.send_message(message.chat.id, f'Ваш ID: {message.from_user.id}')


@bot.callback_query_handler(func=lambda c: c.data.startswith('tpl:'))
def on_template(call):
    key = call.data.split(':', 1)[1]
    bot.answer_callback_query(call.id)
    if key not in TEMPLATES:
        show_menu(call.message.chat.id, 'Такого документа нет, выберите из списка:')
        return
    start_template(call.message.chat.id, call.from_user.id, key)


@bot.callback_query_handler(func=lambda c: c.data in ('make', 'restart', 'cancel'))
def on_confirm(call):
    bot.answer_callback_query(call.id)
    chat_id, user_id = call.message.chat.id, call.from_user.id
    with lock:
        s = sessions.get(user_id)
    # убираем кнопки у сообщения, чтобы не нажали дважды
    try:
        bot.edit_message_reply_markup(chat_id, call.message.message_id, reply_markup=None)
    except Exception:
        pass

    if not s or s['step'] < len(TEMPLATES[s['tpl']]['fields']):
        show_menu(chat_id, 'Сеанс устарел. Выберите документ заново:')
        return
    if call.data == 'cancel':
        with lock:
            sessions.pop(user_id, None)
        show_menu(chat_id, 'Отменено.')
        return
    if call.data == 'restart':
        start_template(chat_id, user_id, s['tpl'])
        return

    # call.data == 'make'
    tpl = TEMPLATES[s['tpl']]
    bot.send_message(chat_id, '⏳ Готовлю документ...')
    try:
        name = Path(tpl['file']).stem + '_' + datetime.now().strftime('%d-%m-%Y_%H-%M')
        docx_path, pdf_path = make_documents(BASE_DIR / tpl['file'], s['answers'],
                                             OUTPUT_DIR / str(user_id), name)
    except Exception as e:
        print('Ошибка при создании документа:', e)
        bot.send_message(chat_id, 'Не получилось создать документ 😔 Попробуйте ещё раз позже.')
        return

    for path in (docx_path, pdf_path):
        with open(path, 'rb') as f:
            bot.send_document(chat_id, f, visible_file_name=path.name)
    bot.send_message(chat_id, '🎉 Готово! Документ в двух форматах: Word — чтобы '
                              'отредактировать, PDF — чтобы распечатать или отправить.')

    if ADMIN_ID and str(user_id) != str(ADMIN_ID):
        try:
            who = call.from_user.full_name
            bot.send_message(ADMIN_ID, f'📥 Новый документ: {tpl["title"]}\nОт: {who}')
            with open(pdf_path, 'rb') as f:
                bot.send_document(ADMIN_ID, f, visible_file_name=pdf_path.name)
        except Exception as e:
            print('Не удалось отправить копию владельцу:', e)

    with lock:
        sessions.pop(user_id, None)
    show_menu(chat_id, 'Подготовить ещё один документ?')


@bot.message_handler(content_types=['text'])
def on_text(message):
    chat_id, user_id = message.chat.id, message.from_user.id
    with lock:
        s = sessions.get(user_id)

    if message.text == CANCEL:
        with lock:
            sessions.pop(user_id, None)
        bot.send_message(chat_id, 'Отменено.', reply_markup=types.ReplyKeyboardRemove())
        show_menu(chat_id)
        return

    if not s:
        show_menu(chat_id, 'Выберите документ кнопкой ниже 👇')
        return

    fields = TEMPLATES[s['tpl']]['fields']
    if s['step'] >= len(fields):
        bot.send_message(chat_id, 'Все ответы уже получены — нажмите кнопку под сводкой.')
        return

    key, _, kind = fields[s['step']]
    value, error = check_answer(kind, message.text)
    if error:
        bot.send_message(chat_id, f'⚠️ {error}\nПопробуйте ещё раз:')
        return

    with lock:
        s['answers'][key] = value
        s['step'] += 1
        done = s['step'] >= len(fields)

    if done:
        show_summary(chat_id, user_id)
    else:
        ask_question(chat_id, user_id)


@bot.message_handler(content_types=['photo', 'document', 'voice', 'sticker', 'video', 'audio'])
def on_other(message):
    bot.send_message(message.chat.id, 'Пожалуйста, отвечайте текстом 🙂')


if __name__ == '__main__':
    print('Бот запущен. Остановить: Ctrl+C')
    bot.infinity_polling(skip_pending=True)
