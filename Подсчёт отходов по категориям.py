import pandas as pd
import os
from os import path
summ_w23 = 0
summ_w5a = 0
summ_w5bc = 0
summ_w678 = 0
summ_w9 = 0
summ_w9a = 0

# current_directory = path.join("" ,'\skrussia.com','dfs-rus','MOS','Production','Сменное задание')
# print(current_directory)
files_in_directory = os.listdir("./")
print(files_in_directory)
file_name = "нет пути"
sheet_name = "Контроль отходов"

mass_w23 = ['Зачист']
mass_w5a = ['Двух', 'отход']
mass_w5bc = ['делам','коро','кро','мор','переп','Пров','разм','толщ','трещ']
mass_w678 = ['высеч','делам','печ','коро','мех','отсут','пере','склад','склей','толщ','трещ']
mass_w9 = ['']
mass_w9a = ['мех', 'загр', 'намок']

mass_mash =['Mini','924','123','Bobst','RDC','Asahi','Tana']

prokladki = 0

for file in files_in_directory:
    if ".xlsx" in file and "Контроль отходов" in file:
        print(file)
        file_name = "./" + file
        break
print(file_name)


if file_name != '':
    df = pd.read_excel(file_name, sheet_name=sheet_name, skiprows = 2)

    df = df.dropna(axis=0, how='all')
    # Удаление столбцов, в которых отсутствуют данные
    df = df.dropna(axis=1, how='all')
    df.columns = df.iloc[0]
    df['N возврат'] = df['N возврат'].fillna(0)
    df['N акт'] = df['N акт'].fillna(0)
    df['N офсет'] = df['N офсет'].fillna(0)
    q =0
    # Вывод полученного DataFrame
    for row in df.itertuples():
        if row.Index != 0:
            try:
                df.iloc[row.Index, 11] = int(df.iloc[row.Index, 11])
            except:
                df.iloc[row.Index, 11] = 0
                Massag = "В столбцах M, N, O были комментарии"

            try:
                df.iloc[row.Index, 12] = int(df.iloc[row.Index, 12])
            except:
                df.iloc[row.Index, 12] = 0
                Massag = "В столбцах M, N, O были комментарии"

            try:
                df.iloc[row.Index, 13] = int(df.iloc[row.Index, 13])
            except:
                df.iloc[row.Index, 13] = 0
                Massag = "В столбцах M, N, O были комментарии"
            if df.iloc[row.Index, 17] == "Прокладки":
                prokladki += df.iloc[row.Index, 15]
            mashin = 0
            vozvrat = 0
            df.iloc[row.Index, 14] = str(df.iloc[row.Index, 14])

            if not("BHS" in df.iloc[row.Index, 3]) and  not("WIP" in df.iloc[row.Index, 3]) and not(any(ext in df.iloc[row.Index, 3] for ext in mass_mash)) :
                vozvrat = 1


            if (("M100" in df.iloc[row.Index, 10] or "WIP" in df.iloc[row.Index, 10]) and vozvrat == 0):
                mashin = -1

            if any(ext in df.iloc[row.Index, 10] for ext in mass_mash) and vozvrat == 0:
                if "загр" in df.iloc[row.Index, 14] and "До" in df.iloc[row.Index, 8] and "Tana" in df.iloc[row.Index, 3]:
                    mashin = -1
                else:
                    mashin = 1

            if df.iloc[row.Index, 10] == "-" and vozvrat==0:
                if "До" in df.iloc[row.Index, 8]:
                    mashin = -1
                else:
                    if "BHS" in df.iloc[row.Index, 3] or "WIP" in df.iloc[row.Index, 3]:
                        mashin = -1
                    if any(ext in df.iloc[row.Index, 3] for ext in mass_mash):
                        mashin = 1



            z = 0
            if any(ext in df.iloc[row.Index, 14] for ext in mass_w23) and int(df.iloc[row.Index, 11]) <= 0 and int(df.iloc[row.Index, 12]) <= 0 and int(df.iloc[row.Index, 13]) <= 0:
                summ_w23 += df.iloc[row.Index, 15]
                z = z + 1


            if any(ext in df.iloc[row.Index, 14] for ext in mass_w5a) and int(df.iloc[row.Index, 11]) <= 0 and int(df.iloc[row.Index, 12]) <= 0 and int(df.iloc[row.Index, 13]) <= 0:
                summ_w5a += df.iloc[row.Index, 15]
                z = z + 1

            if any(ext in df.iloc[row.Index, 14] for ext in mass_w5bc) and int(df.iloc[row.Index, 11]) <= 0 and int(df.iloc[row.Index, 12]) <= 0 and int(df.iloc[row.Index, 13]) <= 0 and  mashin == -1:
                summ_w5bc += df.iloc[row.Index, 15]
                z = z + 1

            if any(ext in df.iloc[row.Index, 14] for ext in mass_w678) and int(df.iloc[row.Index, 11]) <= 0 and int(df.iloc[row.Index, 12]) <= 0 and int(df.iloc[row.Index, 13]) <= 0 and  mashin == 1:
                summ_w678 += df.iloc[row.Index, 15]
                z = z + 1


            if vozvrat == 1 or int(df.iloc[row.Index, 11]) > 0 or int(df.iloc[row.Index, 12]) > 0 or int(df.iloc[row.Index, 13]) > 0:
                summ_w9 += df.iloc[row.Index, 15]
                z = z + 1


            if any(ext in df.iloc[row.Index, 14] for ext in mass_w9a) and int(df.iloc[row.Index, 11]) <= 0 and int(df.iloc[row.Index, 12]) <= 0 and int(df.iloc[row.Index, 13]) <= 0 and  mashin == -1:
                summ_w9a += df.iloc[row.Index, 15]
                z = z + 1



            # if z == 0 and df.iloc[row.Index, 14] != "Side run" and df.iloc[row.Index, 14] != "Side trim":
            #     print(row)
            #     q = q + 1


    print("w23 ", summ_w23/1000)
    print("w5a ", summ_w5a/1000)
    print("w5bc ", summ_w5bc/1000)
    print("w678 ", summ_w678/1000)
    print("w9 ", summ_w9/1000)
    print("w9a ", summ_w9a/1000)
    print('Сумма ', summ_w23 + summ_w5a + summ_w5bc + summ_w678 + summ_w9 + summ_w9a)

    df1 = pd.DataFrame({'A': ['W1       Общий вес отходов  [TON]', '', 'W2,3    Зачистка рулонов и бумага на втулке [TON]',
                                'W4       Трим  [TON]', 'W5a     Отход с обрубочного ножа  [TON]', 'W5b,c  Отходы гофроагрегата (качество) [TON]',
                                'W6,7,8 Отходы станков конвертации  [TON]', 'W9       Прочие отходы  [TON]','W9a     WIP',
                                'W15a   Контролируемые отходы [TON]','W14a   Не контролируемые отходы [TON]','','Сумма по прокладкам'],
                       'B': ['', '', summ_w23/1000, '', summ_w5a/1000, summ_w5bc/1000, summ_w678/1000, summ_w9/1000,summ_w9a/1000, '', '', '', prokladki/1000]})

    with pd.ExcelWriter("./Данные_за_день.xlsx") as writer:
        df1.to_excel(writer, sheet_name='Данные')







