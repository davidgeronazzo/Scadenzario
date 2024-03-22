import pandas as pd
import time
from datetime import datetime, timedelta
import os
import telebot


def calcola_dt():

    Now = datetime.now()

    starting_scantime = datetime(Now.year, Now.month, Now.day, 8, 30)

    if Now.weekday() == 0:
        starting_scantime = Now + timedelta(days=1)
    weekDay = starting_scantime.weekday()

    while weekDay != 0:
        starting_scantime = starting_scantime + timedelta(days=1)
        weekDay = starting_scantime.weekday()

    stopping_scantime = starting_scantime

    delta_t = stopping_scantime - Now
    delta_t = delta_t.seconds + delta_t.days * 3600 * 24

    month = stopping_scantime.month

    if month == 1:
        Mese = "gennaio"
    elif month == 2:
        Mese = "febbraio"
    elif month == 3:
        Mese = "marzo"
    elif month == 4:
        Mese = "aprile"
    elif month == 5:
        Mese = "maggio"
    elif month == 6:
        Mese = "giugno"
    elif month == 7:
        Mese = "luglio"
    elif month == 8:
        Mese = "agosto"
    elif month == 9:
        Mese = "settembre"
    elif month == 10:
        Mese = "ottobre"
    elif month == 11:
        Mese = "novembre"
    elif month == 12:
        Mese = "dicembre"
    else:
        Mese = ""

    dataString = ("Il prossimo alert arriverà lunedì " + str(stopping_scantime.day) + " " + Mese + " "
                  + str(stopping_scantime.year))

    return stopping_scantime, dataString, delta_t


def trovaFile():
    # trova il nome del file
    os.chdir("Q:")
    Dir = os.listdir()

    found = 0

    i = 0
    ScadenzarioFileName = []

    while found == 0:

        if Dir[i][0:21] == "Scadenzario polizze_2":
            ScadenzarioFileName = Dir[i]
            found = 1
        else:
            i = i + 1

    return ScadenzarioFileName


def main():

    # trova il nome del file

    os.chdir("Q:")
    Dir = os.listdir()

    found = 0

    i = 0
    ScadenzarioFileName = []

    while found == 0:

        if Dir[i][0:21] == "Scadenzario polizze_2":
            ScadenzarioFileName = Dir[i]
            found = 1

        else:
            i = i + 1

    Tab = pd.read_excel(ScadenzarioFileName)
    Tab["SCADENZA"] = pd.to_datetime(Tab["SCADENZA"], dayfirst=True)
    Now = datetime.now()
    dt_2mesi = timedelta(days=60)

    deadline = Now + dt_2mesi

    # seleziono le scadenze che mi interessano
    SelTab = Tab[Tab["SCADENZA"] <= deadline]

    # riordino la tabella in base alla data

    Scadenze = SelTab["SCADENZA"]
    timesLeft = Now - Scadenze
    timesLeft = timesLeft.apply(lambda x: x.days)
    SelTab["Tempo residuo"] = timesLeft

    firstTab = SelTab[SelTab["Tempo residuo"] <= 0]
    firstTab = firstTab.sort_values(by=["SCADENZA"])

    SecondTab = SelTab[SelTab["Tempo residuo"] > 0]
    SecondTab = SecondTab.sort_values(by=["SCADENZA"], ascending=False)

    text = "*OCCHIO alle scadenze!*\n \n"

    for i in range(len(firstTab)):

        Scadenza = firstTab["SCADENZA"].iloc[i]
        Polizza = firstTab["Tipologia polizza"].iloc[i]
        Riferimento = firstTab["Riferimento"].iloc[i]

        timeLeft = Now - Scadenza
        timeLeft = timeLeft.days

        if timeLeft <= 0:
            symbol = "⚠️"
            sign = ""
        else:
            symbol = "❌"
            sign = "+"

        ScadenzaString = datetime.strftime(Scadenza, "%d/%m/%Y")
        Compagnia = firstTab["Compagnia assicurativa"].iloc[i]
        string = (sign+str(timeLeft) + " " + symbol + " Polizza "+Polizza + " " + Compagnia + " in scadenza il " +
                  ScadenzaString + " - " + str(Riferimento) + "\n \n")
        text = text + string

    for i in range(len(SecondTab)):

        Scadenza = SecondTab["SCADENZA"].iloc[i]
        Polizza = SecondTab["Tipologia polizza"].iloc[i]
        Nota = str(SecondTab["NOTE"].iloc[i])
        Riferimento = firstTab["Riferimento"].iloc[i]

        if Nota == "nan":
            Nota = ""

        timeLeft = Now - Scadenza
        timeLeft = timeLeft.days

        if timeLeft <= 0:
            symbol = "⚠️"
            sign = ""
        else:
            symbol = "❌"
            sign = "+"

        ScadenzaString = datetime.strftime(Scadenza, "%d/%m/%Y")
        Compagnia = SecondTab["Compagnia assicurativa"].iloc[i]
        string = (sign+str(timeLeft) + " " + str(Nota) + " " + symbol + " Polizza "+Polizza + " " + Compagnia +
                  " scaduta il " + ScadenzaString + " - " + Riferimento + "\n \n")
        text = text + string

    t, dataString, tLeft = calcola_dt()
    text = text + dataString

    token = "6007635672:AAF_kA2nV4mrscssVRHW0Fgzsx0DjeZQIHU"
    bot = telebot.TeleBot(token)
    test_id = "-672088289"
    run_id = "-1001995962404"
    bot.send_message(test_id, text=text, parse_mode='Markdown')

    return tLeft, t


now = datetime.now()
NextSend = now - timedelta(hours=1)

while True:

    curr_dir = os.getcwd()
    now = datetime.now()
    last_modify = pd.read_csv("last_modify.csv")
    last_modify = last_modify["last_modify"][0]
    last_modify = pd.to_datetime(last_modify)
    # last_modify = datetime.hrtime(last_modify)
    FileName = trovaFile()
    modify_time = os.path.getmtime("Q:"+FileName)
    modify_time = datetime.fromtimestamp(modify_time)

    # modify_time = last_modify[0]

    if modify_time > last_modify or now > NextSend:
        dtLeft, NextSend = main()
        modify_time_dict = {"last_modify": modify_time}
        modify_time_df = pd.DataFrame(modify_time_dict, index=[0])
        os.chdir(curr_dir)
        modify_time_df.to_csv("last_modify.csv", index=False)
        print("Prossimo alert tra: "+str(round(dtLeft/3600/24)) + " giorni.")

    os.chdir(curr_dir)

    dt = 60
    time.sleep(dt)
