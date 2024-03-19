import pandas as pd
import time
from datetime import datetime, timedelta
import os
import telebot


def calcola_dt():

    Now = datetime.now()
    currMonth = Now.month
    currYear = Now.year
    nextMonth = currMonth % 12 + 1

    if nextMonth != 1:
        nextMonthYear = currYear
    else:
        nextMonthYear = currYear + 1

    starting_scantime = datetime(nextMonthYear, nextMonth, 1, 8, 30)
    weekDay = starting_scantime.weekday()

    while weekDay != 0:
        starting_scantime = starting_scantime + timedelta(days=1)
        weekDay = starting_scantime.weekday()

    stopping_scantime = starting_scantime

    delta_t = stopping_scantime - Now
    delta_t = delta_t.seconds + delta_t.days * 3600 * 24

    return delta_t


def main():

    # trova il nome del file
    os.chdir("Q:")
    Dir = os.listdir()

    found = 0

    i = 0
    FileName = []

    while found == 0:

        if Dir[i][0:26] == "TEST_Scadenzario polizze_2":
            FileName = Dir[i]
            found = 1
        else:
            i = i + 1

    Tab = pd.read_excel(FileName)
    Tab["SCADENZA"] = pd.to_datetime(Tab["SCADENZA"])
    Now = datetime.now()
    dt_2mesi = timedelta(days=60)

    deadline = Now + dt_2mesi

    # seleziono le scadenze che mi interessano
    SelTab = Tab[Tab["SCADENZA"] <= deadline]

    # riordino la tabella in base alla data
    SortTab = SelTab.sort_values(by=["SCADENZA"], ascending=False)

    text = "*SCADENZE ENTRO I PROSSIMI GIORNI:*\n \n"
    for i in range(len(SortTab)):

        Scadenza = SortTab["SCADENZA"].iloc[i]
        Polizza = SortTab["Tipologia polizza"].iloc[i]

        timeLeft = Now - Scadenza
        timeLeft = timeLeft.days

        if timeLeft <= 0:
            symbol = "⚠️"
            sign = ""
        else:
            symbol = "❌"
            sign = "+"

        ScadenzaString = datetime.strftime(Scadenza, "%d/%m/%Y")
        Compagnia = SortTab["Compagnia assicurativa"].iloc[i]
        string = (sign+str(timeLeft) + " " + symbol + " Polizza "+Polizza + " " + Compagnia + " in scadenza il " +
                  ScadenzaString + "\n \n")
        text = text + string

    token = "6007635672:AAF_kA2nV4mrscssVRHW0Fgzsx0DjeZQIHU"
    bot = telebot.TeleBot(token)
    test_id = "-672088289"
    run_id = "-4038294282"
    bot.send_message(test_id, text=text, parse_mode='Markdown')


while True:
    main()
    dt = calcola_dt()
    print("Prossimo alert tra: "+str(round(dt/3600/24)) + " giorni.")
    time.sleep(dt)
