import os
import xlrd
folder_path = r'Excel'

def get_frequency_dict():
    print('Hello world')

for file in os.listdir(folder_path):
    ext = file.split('.')[1]
    if ext in ('xls', 'xlsx'):
        workbook = xlrd.open_workbook(f"Excel/{file}")
        worksheet = workbook.sheet_by_index(0)
        get_frequency_dict()
        os.remove(f"Excel/{file}")
        quit()
print('\tВ папке нет excel-файлов')
quit()