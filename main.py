import time

import pywhatkit
import pandas as pd
from datetime import datetime


def send_message_person(number, name):
    try:
        #time.sleep(5)
        pywhatkit.sendwhatmsg_instantly("+39" + str(int(number)), "Tanti auguri " + name + " 🎂🥳\t", 15, True, 5)
        #pywhatkit.sendwhatmsg("+39" + str(int(number)), "Tanti auguri " + name + " 🎂\t", 15, True, 5)
        #time.sleep(5)
    except:
        print("Errore nell'invio del messaggio")


def send_message_group(id_group, name):
    try:
        #time.sleep(5)
        #pywhatkit.sendwhatmsg_to_group(id_group, "Tanti auguri " + name + " 🎂\t", int(str(datetime.now().hour)), int(str(datetime.now().minute))+1, 15, True, 5)
        pywhatkit.sendwhatmsg_to_group_instantly(id_group, "Tanti auguri" + name + " :cake\t", 15, True, 5)
        #time.sleep(5)
    except:
        print("Errore nell'invio del messaggio")


def send_happy_birthday(path):
    df = pd.read_excel(path)
    #print(df)
    day = int(datetime.today().day)
    month = int(datetime.today().month)

    df_loc_ready = df.loc[(df['mese'] == month) & (df['giorno'] == day)]
    str_out = "Oggi compiono gli anni: \n"
    #print(df_loc_ready)
    for i in df_loc_ready.index:
        nickname = df_loc_ready['nickname'][i]

        if df_loc_ready['gruppo'][i] == "none":
            number = df_loc_ready['numero'][i]
            send_message_person(number, nickname)
            str_out = str_out + str(df_loc_ready['nome'][i]) + "\n"
        else:
            id_group = df_loc_ready['gruppo'][i]
            send_message_group(id_group, nickname)
            str_out = str_out + str(df_loc_ready['nome'][i]) +"\n"

    return str_out


print(send_happy_birthday("birthdays.xlsx"))
time.sleep(5)

