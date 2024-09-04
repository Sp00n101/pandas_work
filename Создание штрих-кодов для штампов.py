from PIL import Image
import barcode
from barcode.writer import ImageWriter
import pandas as pd
import os


def ad_list(png):
    combined_image = Image.new('RGB', (1000, 1414))  # Размеры листа подбираются под ваши нужды
    x_offset = 0
    y_offset = 0
    for png_file in png:
        img = Image.open(png_file)
        print(img.width,img.height)
        combined_image.paste(img, (x_offset, y_offset))
        x_offset += img.width  # Измените смещение по горизонтали для следующего штрихкода
        if x_offset >= 1001 - 1.15*img.width:
            x_offset = 0
            y_offset += img.height

    combined_image.save('объединенные_штрихкоды_' + str(list) + '.pdf', save_all=True)

files_in_directory = os.listdir("./shtamp//")
sheet_name = "Реестр"
qw = []

for file in files_in_directory:
    if ".xlsm" in file:
        file_name = "./shtamp//" + file
        print(file_name)
        df = pd.read_excel(file_name, sheet_name=sheet_name, skiprows=0, usecols="A")
        qw.append(df)
df = pd.concat(qw, axis=0)
df = df.reset_index(drop=True)
df.columns = ["Штамп"]
codes = list(df["Штамп"])
print(codes)
print(len(codes))


for i in codes:
    code128 = barcode.get_barcode_class('code128')
    code = code128(str(i), writer=ImageWriter())

    # Сохраняем штрих-код в формате PNG
    filename = code.save(".\\direct\\"+str(i)+".png", {"module_width":0.22, "module_height": 4, "font_size": 5, "text_distance": 4, "quiet_zone": 1})



path = ".\\direct\\"
dir_list = os.listdir(path)


png_files = list(map(lambda num: ".\\direct\\" + num, dir_list))

list = 1

for i in range(0, len(png_files), 40):
    if i+40 > len(png_files):
        ad_list(png_files[i:len(png_files)])
    else:
        ad_list(png_files[i:i+40])
    list += 1



