import numpy as np
import pandas as pd
import os
import re
import openpyxl
import tkinter as tk
propusk = []
name_list = ""
colls = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10","11", "12", "13", "14", "15","16", "17", "18", "19", "20","21", "22", "23", "24", "25","26", "27", "28", "29", "30", "31"]

def on_button_click():
    global name_list
    name_list = entry.get()
    root.destroy()

def tkin():
    global root
    root = tk.Tk()
    root.title("Ввод данных")

    lbl = tk.Label(root, text="Введите название нужного листа:")
    lbl.pack()

    # Поле ввода
    global entry
    entry = tk.Entry(root)
    entry.pack(padx=20, pady=10)

    # Кнопка
    button = tk.Button(root, text="Ввести", command=on_button_click)
    button.pack(pady=10)

    # Глобальная переменная для хранения данных
    user_input = None

    # Запускаем цикл обработки событий
    root.mainloop()


def transform_fio(full_name):
    parts = full_name.split()
    last_name = parts[0]
    initials = ''.join([f"{name[0]}." for name in parts[1:]])
    if full_name == "Литвинов Михаил Николаевич":
        last_name = 'литвинов'
        initials = "М.Н."
    return f"{last_name} {initials}"


def conv_str(stroka:str):
    x1 = np.nan
    x2 = np.nan
    if "<>" in stroka:
        x1 = "ВП"
        x2 = np.nan
    elif "--" in stroka:
        x1 = np.nan
        x2 = np.nan
    else:
        if "Б" in stroka:
            x1 = "Б"
            x2 = np.nan
        elif "К" in stroka:
            x1 = "К"
            x2 = np.nan
        elif "О" in stroka:
            x1 = "О"
            x2 = np.nan
        else:
            if "Н" in stroka:
                x1 ="нв"
            match = re.search(r'Я(\(вр\))?\s*([0-9.]+)', stroka)
            if match:
                value = match.group(2)
                if value[-1] == "0":
                    x2 = str(int(float(value)))
                else:
                    x2 = value
            else:
                x2 = np.nan

            if "ВП" in stroka:
                match = re.search(r'ВП(\d+\.\d+)', stroka)
                value = match.group(1)
                if value[-1] == "0":
                    x2 = str(int(float(value)))
                else:
                    x2 = value
                x1, x2 = x2, x1
    return str(x1)+str(x2)




def tabel1(sheet_name):
    files_in_directory = os.listdir("./")


    for file in files_in_directory:
        if ".xlsx" in file and "НОВЫЙ ТАБЕЛЬ" in file:
            file_name = "./" + file
            break
    if file_name != '':
        df = pd.read_excel(file_name, sheet_name=sheet_name, skiprows=3, usecols="E, H:AL")
        df.columns = ["ФИО", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20","21", "22", "23", "24", "25","26", "27", "28", "29", "30", "31"]
    return df

def tabel2():
    files_in_directory = os.listdir("./")
    sheet_name = "Лист_1"
    qw = []

    for file in files_in_directory:
        if ".xlsx" in file and "Табель учета" in file:
            file_name = "./" + file

            df = pd.read_excel(file_name, sheet_name=sheet_name, skiprows=13, usecols="B, D:S")
            df.columns = ["ФИО", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"]
            df = df.loc[df["1"] != "17"]
            df = df.loc[df['1'] != "1"]
            df = df.loc[df['1'] != "<1>"]
            df = df.loc[df['1'] != "<17>"]
            df = df.loc[df['1'] != "Отметка о явках и неявках на работу по числам месяца"]
            df = df.reset_index(drop=True)
            df = df[:-13]
            df1 = pd.DataFrame([], columns=["ФИО", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"])
            df2 = pd.DataFrame([], columns=["ФИО", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"])
            for i in range(0, len(df), 2):
                df1.loc[i] = df.loc[i]
                df2.loc[i] = df.loc[i+1]
            df1 = df1.reset_index(drop=True)
            df2 = df2.reset_index(drop=True)
            df2.columns = ["ФИО", "17", "18", "19", "20","21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32"]
            df2 = df2.drop(["ФИО", "32"], axis=1)
            df = pd.concat([df1, df2], axis=1)
            qw.append(df)
    df = pd.concat(qw)
    df = df.reset_index(drop=True)
    return df


tkin()
df_Tabel = tabel1(name_list)
df_tab = tabel2()
oshibki = []
propusk_people = []
for index, row in df_tab.iterrows():
    surname = transform_fio(row['ФИО'])
    matching_row = df_Tabel[df_Tabel['ФИО'] == surname]
    if not matching_row.empty:
        # print(index, matching_row.index)
        # print(index, matching_row.index[0])

        for column in colls:  # Проход по остальным столбцам, начиная со второго
            try:
                if conv_str(row[column]) != str(df_Tabel.loc[matching_row.index[0],column])+str(df_Tabel.loc[matching_row.index[1],column]):
                    if str(df_Tabel.loc[matching_row.index[0],column]) == "х":
                        continue
                    else:
                        oshibki.append(f"У человека {row['ФИО']} есть различия в {column} дне.")
            except:
                propusk.append(surname)
    else:
        propusk_people.append(surname)
print(df_tab)
print(len(oshibki))
print(*oshibki)
print(propusk)
print(propusk_people)
oshibki.append("")
oshibki.append("")
oshibki.append("Фамилии не обработанные по причине не соответсвия в документах:")
oshibki += list(set(propusk)) + propusk_people
workbook = openpyxl.Workbook()
sheet = workbook.active


# Записываем данные в столбик
for i in range(len(oshibki)):
    sheet.cell(row=i+1, column=1, value=oshibki[i])
workbook.save('Отличия в данных.xlsx')