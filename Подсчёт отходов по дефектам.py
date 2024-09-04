import numpy as np
import pandas as pd
import os
import re
from openpyxl.workbook.child import INVALID_TITLE_REGEX


mash = ["Asahi/Gopfert 2100", "Martine 924 Msk", "Martine 1232 Msk", "Miniline 616 Msk", "RDC Goepfert 1624",
        "Bobst/Flexo 1600 Msk", "Tanabe JD BoxR 1450"]
mass_w5bc = ['делам', 'коро', 'кро', 'мор', 'переп', 'Пров', 'разм', 'толщ', 'трещ']

gof = ['BHS', "WIP", "-"]

files_in_directory = os.listdir("./")
print(files_in_directory)
file_name = "нет пути"
sheet_name = "Контроль отходов"


for file in files_in_directory:
    if ".xlsx" in file and "Контроль отходов" in file:
        print(file)
        file_name = "./" + file
        break
print(file_name)
if file_name != '':
    df = pd.read_excel(file_name, sheet_name=sheet_name, skiprows=2)

    df = df.dropna(axis=0, how='all')
    # Удаление столбцов, в которых отсутствуют данные
    df = df.dropna(axis=1, how='all')
    df.columns = df.iloc[0]
    df = df.drop(index=0)

    w = df[df['Машина'] == "BHS 2500"]
    w1 = w[w['Тип отхода'] == "После машины"]
    w1 = w1[w1['Дефект'].apply(lambda x: any(sub in x for sub in mass_w5bc))]
    w1 = w1.groupby(['Машина', 'Дефект']).agg({'Вес кг': ['sum']}).reset_index()

    q = df[df.Машина.isin(mash)]
    w2 = q[q['Тип отхода'] == "До машины"]
    w2 = w2[w2['Дефект'].apply(lambda x: any(sub in x for sub in mass_w5bc))]
    w2 = w2[w2['Источник отхода'].apply(lambda x: any(sub in x for sub in gof))]
    w2 = w2.groupby(['Дефект']).agg({'Вес кг': ['sum']}).reset_index()

    w12 = [w1, w2]
    w22 = pd.concat(w12, axis=0)
    w22 = w22.reset_index(drop=True)
    w22.columns = w22.columns.droplevel(1)
    w22 = w22.groupby(['Дефект']).agg({'Вес кг': ['sum']}).reset_index()

    ma = df[df['Тип отхода'] == "После машины"]
    ma = ma.loc[ma['N возврат'].isnull() | (ma['N возврат'] == '') | (ma['N возврат'] == "0")]
    ma = ma.loc[ma['N акт'].isnull() | (ma['N акт'] == '') | (ma['N акт'] == "0")]
    ma = ma.loc[ma['N офсет'].isnull() | (ma['N офсет'] == '') | (ma['N офсет'] == "0")]
    ma = ma.groupby(['Машина', 'Дефект']).agg({'Вес кг': ['sum']}).reset_index()
    ma = ma[ma.Машина.isin(mash)]

    ma.columns = ma.columns.droplevel(1)
    ma1 = ma.groupby(['Дефект']).agg({'Вес кг': ['sum']}).reset_index()
    print(ma1)

    data = []
    with pd.ExcelWriter("./Данные_за_месяц.xlsx") as writer:
        w1.to_excel(writer, sheet_name=re.sub(INVALID_TITLE_REGEX, '_', "BHS 2500_1"))
        w2.to_excel(writer, sheet_name=re.sub(INVALID_TITLE_REGEX, '_', "BHS 2500_2"))
        for i in mash:
            asd = ma[ma['Машина'] == i]
            asd.to_excel(writer, sheet_name=re.sub(INVALID_TITLE_REGEX, '_', i), index=True)

        w22.to_excel(writer, sheet_name=re.sub(INVALID_TITLE_REGEX, '_', " Сумма отходов по BHS 2500"))
        ma1.to_excel(writer, sheet_name=re.sub(INVALID_TITLE_REGEX, '_', " Сумма отходов по машинам"))
