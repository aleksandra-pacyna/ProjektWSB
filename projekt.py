import datetime
import os
import regex as re
import shutil
from openpyxl import load_workbook


print('Podaj lokalizację.')
path = input()
location_files = path

shutil.rmtree('data - nowy folder')

os.mkdir('data - nowy folder')
print('Utworzono nowy folder o nazwie: "data - nowy folder\n"')

list_of_files_in_folder = os.listdir(path)

org_path = path
new_path = r'C:\Users\48690\PycharmProjects\python\ZUS-przelewy-pythonProject\przelewy ZUS\data - nowy folder'


for element in list_of_files_in_folder:
    file_name = org_path + '\\' + element
    new_file_name = new_path + '\\' + element
    shutil.copy(file_name, new_file_name, follow_symlinks=True)
print('Pliki zostały skopiowane.\n')


month_before_name = '0'+ str(datetime.date.today().month)
month_current_name = '0'+ str(datetime.date.today().month + 1 )

folder = new_path
objs = os.listdir(folder)

for src in objs:
   full_src = os.path.join(folder, src)
   if os.path.isfile(full_src):
       dst = src.replace(month_before_name, month_current_name)
       if src != dst:
           full_dst = os.path.join(folder, dst)
           os.rename(full_src, full_dst)
print('Nazwy plików zostały zmienione.\n')

file_list_in_new_folder = os.listdir(new_path)
print(file_list_in_new_folder)

myregex = '^[0-9]+,[0-9]{8},[0-9]+,[0-9]+,[0-9],"[0-9]+","[0-9]+","[^"]+","[^"]+",'
regex_amount = r'^[0-9]{3},[0-9]{8},\K[0-9]+'
regex_amount_empty = '^[0-9]{3},[0-9]{8},'
success = []
fail = []


wb = load_workbook("MATRYCA DO WYSYŁKI PRZELEWÓW.xlsx")
sheet = wb.active

max_row = sheet.max_row
max_column = sheet.max_column


for element in file_list_in_new_folder:
    new_path = os.path.join(new_path)
    print(new_path + '\\' + element)


    open_file = open(new_path + '\\' + element, encoding="ANSI")
    file_txt_content = open_file.read()

    print('\n')
    print(file_txt_content)
    x = re.search(myregex, file_txt_content)
    if x:
        success.append(element[0:4])
        print('\n Udało się, plik: ' + element)
        for i in range(1, max_row + 1):
            if str(sheet.cell(row=i, column=1).value) == element[0:4]:
                print('\n\n')
                y = re.search(regex_amount, file_txt_content)
                if y:
                    f = open(new_path + '\\' + element, "w", encoding="ANSI")
                    nowa_zawartosc = re.sub(regex_amount, str(sheet.cell(row=i, column=6).value), file_txt_content)
                    f.write(nowa_zawartosc)
                    f.close()
    else:
        fail.append(element[0:4])
        print("NIE UDAŁO SIĘ \n" + element)
print(success)
print(fail)









