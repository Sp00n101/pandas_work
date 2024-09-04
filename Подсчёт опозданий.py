from datetime import time
import pandas as pd
import os
import tkinter as tk

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


def opozd(sob, start, starn_min, finish, finish_min):
    resalt.append([])

    df1 = df[df['событие'] == sob]
    df1['время'] = pd.to_datetime(df1['время']).dt.time
    df1 = df1[ df1['время'].between(time(start, starn_min, 0), time(finish, finish_min, 0))]
    df1 = df1[['дата', 'время', 'событие', 'субъект']]
    for index, row in df1.iterrows():
        surname = transform_fio(row['субъект'])
        matching_row = tabel[tabel['ФИО'] == surname]
        if not matching_row.empty:
            try:
                if conv_str(row[column]) != str(df_Tabel.loc[matching_row.index[0], column]) + str(
                        df_Tabel.loc[matching_row.index[1], column]):
                    if str(df_Tabel.loc[matching_row.index[0], column]) == "х":
                        continue
                    else:
                        oshibki.append(f"У человека {row['ФИО']} есть различия в {column} дне.")
            except:
                print(f'{row["субъект"]}')
        else:
            print(f'{row["субъект"]}')

name_list = ""
tkin()
tabel = tabel1(name_list)
files_in_directory = os.listdir("./")
print(files_in_directory)
sheet_name = "Page1"
qw = []
resalt = []
for file in files_in_directory:
    if ".xlsx" in file in file:
        file_name = "./" + file

        df = pd.read_excel(file_name, sheet_name=sheet_name, skiprows=0)
        df.columns = ["дата", "время", "событие", "Источник", "субъект", "сводка"]
        qw.append(df)
df = pd.concat(qw)
df = df.loc[df['дата'] != "Дата"]

opozd("Нормальный вход по ключу", 7,0,7,20)
opozd("Нормальный вход по ключу", 19,0,19,20)
opozd("Нормальный выход по ключу", 6,30,7,0)
opozd("Нормальный выход по ключу", 18,30,19,0)

with pd.ExcelWriter("./Опоздания.xlsx") as writer:
    for i in range(4):
        q = -1
        match i:
            case 0:
                q = "Опоздания утром"
            case 1:
                q = "Опоздания вечером"
            case 2:
                q = "Ранний уход утром"
            case 3:
                q = "Ранний уход вечером"
        resalt[i].to_excel(writer, sheet_name=str(q), index=True)

