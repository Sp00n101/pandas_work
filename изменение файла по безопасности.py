import pandas as pd
import os
import win32com.client
global df2

ver = {
    '1 - Почти невозможно': "Не вероятно",
    '2 - Возможно': "Маловероятно",
    '3 - Уверены, что событие может произойти': "Вероятно",
    "4 - Опасное событие происходило ранее": "Скорее всего",
    '5 - Опасное событие несколько раз происходило': "Почти точно",
}

tyj = {
    '1 - Первая помощь или никаких травм': "Отсутствие",
    '2 - Случай медицинского лечения': "Лёгкая",
    '3 - Потеря рабочего времени': "Средняя",
    '4 - Инвалидность, необратимая травма': "Тяжёлая",
    '5 - Летальный исход': "Летальный",
}


files_in_directory = os.listdir("./")
file_df = "нет пути"
sheet_df = "Sheet"
file_df1 = "нет пути"
sheet_df1 = "REGISTRY"


for file in files_in_directory:
    if ".xlsx" in file and "Soobshchenie" in file:
        file_df = "./" + file
    if ".xlsx" in file and "NearMiss" in file:
        file_df1 = "./" + file


if file_df1 != '' and file_df != '':
    df = pd.read_excel(file_df, sheet_name=sheet_df,  converters={'ID': str})
    df1 = pd.read_excel(file_df1, sheet_name=sheet_df1, usecols="A", skiprows=1, converters={'ID': str})

    df2 = df[df["О чём Вы хотите сообщить"] == "Безопасность"]
    df2 = df2[df2['Площадка'] == "Москва"]
    df2['Время создания'] = pd.to_datetime(df['Время создания'])

    df2['Время'] = df2['Время создания'].astype(str)
    df2['Область повышенного риска'] = df2['Область повышенного риска'].astype(str)
    df2['Тип сообщения'] = df2['Тип сообщения'].astype(str)

    df2 = df2[~df2['ID'].isin(df1["ID"])]


    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = True

    workbook = excel.Workbooks.Open(os.path.abspath(file_df1))
    sheet = workbook.Sheets(sheet_df1)

    i = 3

    while True:
        if sheet.Cells(i, 1).Value is None and sheet.Cells(i, 2).Value is None:
            for index, row in df2.iterrows():
                sheet.Cells(i, 1).Value = row['ID']
                sheet.Cells(i, 3).Value = row['Подробно опишите ситуацию']
                sheet.Cells(i, 4).Value = row['Приложите фотографию места']
                if row['Тип сообщения'] != "nan":
                    sheet.Cells(i, 5).Value = row['Тип сообщения']
                sheet.Cells(i, 6).Value = ver[row['Оцените вероятность опасного события']]
                sheet.Cells(i, 7).Value = tyj[row['Оцените возможные последствия']]
                if row['Область повышенного риска'] != "nan":
                    sheet.Cells(i, 9).Value = row['Область повышенного риска']
                sheet.Cells(i, 14).Value = row['Определите тип воздействия']
                sheet.Cells(i, 15).Value = row['Время']
                sheet.Cells(i, 16).Value = row['Место (участок/оборудование)-Москва']
                sheet.Cells(i, 17).Value = row['ФИО сообщившего']

                i += 1
            break
        i += 1
    workbook.Save()
    workbook.Close()
    excel.Quit()