import pandas as pd
import os
import re
from openpyxl.workbook.child import INVALID_TITLE_REGEX


def nedodelki():
    files_in_directory = os.listdir("./")
    file_name = ""
    for file in files_in_directory:
        if ".xlsx" in file and "Недоделки" in file:
            file_name = "./" + file
            break
    if file_name != '':
        df = pd.read_excel(file_name, sheet_name=0, usecols="A, M:P", skiprows=3)
        df.columns = ["Заказ", "Брак", "Повреждение", "Качество", "Другое"]
        df = df.fillna('')
        df['Комментарий'] = (df['Брак'].astype(str) + " " + df['Повреждение'].astype(str) + " " + df['Качество'].astype(
            str) + " " + df['Другое'].astype(str))
        df.reset_index(drop=True)
        return df


def otif():
    files_in_directory = os.listdir("./")
    file_name = ""
    for file in files_in_directory:
        if ".xlsx" in file and "OTIF" in file:
            file_name = "./" + file
            break
    if file_name != '':
        df = pd.read_excel(file_name, sheet_name="IF", skiprows=0, usecols="A,N,U")
        df.columns = ["Заказ", "Проверка", "Коментарий"]
        df = df.fillna('')
        df.reset_index(drop=True)
        return df


df1 = pd.DataFrame()
ot = otif()
nedo = nedodelki()
for index, row in ot.iterrows():
    surname = row['Заказ']
    matching_row = nedo[nedo['Заказ'] == surname]
    if row['Проверка'] == "":
        if not matching_row.empty:
            df1 = df1._append(matching_row, ignore_index=True)
        else:
            new_row = {'Заказ': surname, 'Комментарий': ''}
            df1 = df1._append(new_row, ignore_index=True)
df1 = df1[["Заказ", "Комментарий"]]
df1 = df1.groupby(['Заказ'], sort=False)["Комментарий"].apply(lambda x: '; '.join(x)).reset_index()
for index, row in df1.iterrows():
    df1.iloc[index, 1] = str(df1.iloc[index, 1]).strip()
with pd.ExcelWriter("./Данные_OTIF.xlsx") as writer:
    df1.to_excel(writer, sheet_name=re.sub(INVALID_TITLE_REGEX, '_', "OTIF"))
