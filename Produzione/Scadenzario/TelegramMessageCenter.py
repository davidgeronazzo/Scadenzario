import telebot

token = "6007635672:AAF_kA2nV4mrscssVRHW0Fgzsx0DjeZQIHU"
bot = telebot.TeleBot(token)
test_id = "-672088289"
run_id = "-1001995962404"


def segnalaErrore(Mode, text):

    if Mode == "TEST":
        chat_id = test_id
    else:
        chat_id = run_id

    if text == "[Errno 13] Permission denied: 'Scadenzario polizze_2024.07.02.xlsx.sb-e2d485dc-6vi1nV'":
        text = "Errore: permesso negato!"

    bot.send_message(chat_id, text=text, parse_mode='Markdown')


def mandaScadenze(Mode, text):
    if Mode == "TEST":
        chat_id = test_id
    else:
        chat_id = run_id

    bot.send_message(chat_id, text=text, parse_mode='Markdown')
