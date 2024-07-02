import pandas as pd
import time
from datetime import datetime, timedelta
import os
from TelegramMessageCenter import segnalaErrore, mandaScadenze
import csv
from ftplib import FTP
def salvaWatchdog():
    now = datetime.now()
    ultimo_ciclo = {"t": now}

    with open("lastRunScadenzario.csv", 'w') as csvfile:

        writer = csv.DictWriter(csvfile, ultimo_ciclo.keys())
        writer.writeheader()
        writer.writerow(ultimo_ciclo)

    ftp = FTP("192.168.10.211", timeout=120)
    ftp.login('ftpdaticentzilio', 'Sd2PqAS.We8zBK')

    ftp.cwd('/dati/Database_Produzione')

    file_name = "lastRunScadenzario.csv"
    file = open(file_name, "rb")
    ftp.storbinary(f"STOR " + file_name, file)
    ftp.close()

def preparaResoconto(tab, Mode):
    now_main = datetime.now()
    dt_2mesi = timedelta(days=60)

    deadline = now_main + dt_2mesi

    # seleziono le scadenze che mi interessano
    sel_tab = tab[tab["SCADENZA"] <= deadline]

    # riordino la tabella in base alla data

    scadenze = sel_tab["SCADENZA"]
    times_left = now_main - scadenze
    times_left = times_left.apply(lambda x: x.days)
    sel_tab["Tempo residuo"] = times_left

    first_tab = sel_tab[sel_tab["Tempo residuo"] <= 0]
    first_tab = first_tab.sort_values(by=["SCADENZA"])
    second_tab = sel_tab[sel_tab["Tempo residuo"] > 0]
    second_tab = second_tab.sort_values(by=["SCADENZA"], ascending=False)

    text = "*OCCHIO alle scadenze!*\n \n"

    symbol = "⚠️"
    sign = ""

    for i in range(len(first_tab)):

        scadenza = first_tab["SCADENZA"].iloc[i]
        polizza = first_tab["Tipologia polizza"].iloc[i]
        riferimento = first_tab["Riferimento"].iloc[i]

        time_left = now_main - scadenza
        time_left = time_left.days

        scadenza_string = datetime.strftime(scadenza, "%d/%m/%Y")
        compagnia = first_tab["Compagnia assicurativa"].iloc[i]
        string = (sign + str(time_left) + " " + symbol + " polizza " + polizza + " " + compagnia + " in scadenza il " +
                  scadenza_string + " - " + str(riferimento) + "\n \n")
        text = text + string

    symbol = "❌"
    sign = "+"

    for i in range(len(second_tab)):

        scadenza = second_tab["SCADENZA"].iloc[i]
        polizza = second_tab["Tipologia polizza"].iloc[i]
        nota = str(second_tab["NOTE"].iloc[i])
        riferimento = second_tab["Riferimento"].iloc[i]

        if nota == "nan":
            nota = ""

        time_left = now_main - scadenza
        time_left = time_left.days

        scadenza_string = datetime.strftime(scadenza, "%d/%m/%Y")
        compagnia = second_tab["Compagnia assicurativa"].iloc[i]
        string = (sign + str(time_left) + " " + str(nota) + " " + symbol + " polizza " + polizza + " " + compagnia +
                  " scaduta il " + scadenza_string + " - " + riferimento + "\n \n")
        text = text + string

    t, data_string, t_left = calcola_dt()
    text = text + data_string

    mandaScadenze(Mode, text)

    t_left = "Prossimo alert tra "
    return t_left, t


def calcola_dt():

    now_dt = datetime.now()

    starting_scantime = datetime(now_dt.year, now_dt.month, now_dt.day, 8, 30)

    if now_dt.weekday() == 0:
        starting_scantime = now_dt + timedelta(days=1)
    week_day = starting_scantime.weekday()

    while week_day != 0:
        starting_scantime = starting_scantime + timedelta(days=1)
        week_day = starting_scantime.weekday()

    stopping_scantime = starting_scantime

    delta_t = stopping_scantime - now_dt
    delta_t = delta_t.seconds + delta_t.days * 3600 * 24

    month = stopping_scantime.month

    if month == 1:
        mese = "gennaio"
    elif month == 2:
        mese = "febbraio"
    elif month == 3:
        mese = "marzo"
    elif month == 4:
        mese = "aprile"
    elif month == 5:
        mese = "maggio"
    elif month == 6:
        mese = "giugno"
    elif month == 7:
        mese = "luglio"
    elif month == 8:
        mese = "agosto"
    elif month == 9:
        mese = "settembre"
    elif month == 10:
        mese = "ottobre"
    elif month == 11:
        mese = "novembre"
    elif month == 12:
        mese = "dicembre"
    else:
        mese = ""

    data_string = ("Il prossimo alert arriverà lunedì " + str(stopping_scantime.day) + " " + mese + " "
                   + str(stopping_scantime.year))
    print(data_string)
    return stopping_scantime, data_string, delta_t


def trova_file():
    # trova il nome del file

    # os.chdir("Z:") # run PC
    os.chdir("Q:")  # my PC

    server_dir = os.listdir()

    found = 0

    i = 0
    scadenzario_file_name = []

    while found == 0:

        if server_dir[i][0:21] == "Scadenzario polizze_2":
            scadenzario_file_name = server_dir[i]
            found = 1
        else:
            i = i + 1

    return scadenzario_file_name


def main():

    Mode = "TEST"
    t_left = []
    t = []
    scadenzario_file_name = trova_file()

    tab = pd.read_excel(scadenzario_file_name)

    try:
        tab["SCADENZA"] = pd.to_datetime(tab["SCADENZA"], dayfirst=True)
        t_left, t = preparaResoconto(tab, Mode)
        t_left = "Prossimo alert tra: "+str(round(dtLeft/3600/24)) + " giorni."

    except Exception as err:
        errString = str(err)
        if errString[:4] == "year":
            annoSbagliato = errString[5:11]
            text = "E' stato inserito l'anno anomalo "+annoSbagliato+"\nProssimo alert tra 2 ore."
            segnalaErrore(Mode, text)
            print(err)
            t = datetime.now() + timedelta(seconds=15)
            t_left = "Prossimo alert tra 2 ore"

    return t_left, t


now = datetime.now()
NextSend = now - timedelta(hours=1)

while True:

    RunDisk = "Z"
    TestDisk = "Q"

    currDisk = TestDisk

    curr_dir = os.getcwd()
    now = datetime.now()
    os.chdir(curr_dir)

    last_modify = pd.read_csv("last_modify.csv")
    last_modify = last_modify["last_modify"][0]
    last_modify = pd.to_datetime(last_modify)
    # last_modify = datetime.hrtime(last_modify)
    FileName = trova_file()
    modify_time = os.path.getmtime(currDisk+":"+FileName)
    modify_time = datetime.fromtimestamp(modify_time)

    # modify_time = last_modify[0]

    if modify_time > last_modify or now > NextSend:
        dtLeft, NextSend = main()
        modify_time_dict = {"last_modify": modify_time}
        modify_time_df = pd.DataFrame(modify_time_dict, index=[0])
        modify_time_df.to_csv("last_modify.csv", index=False)

        os.chdir(curr_dir)


    salvaWatchdog()

    dt = 15
    time.sleep(dt)
