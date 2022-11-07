import os
import pprint
import openpyxl
import pandas as pd


# file = pd.read_csv('csv/IIT_mag_2kurs_osen_2022_2023.csv')
# dic = {}
# print(
#     'Шошников И.К.\n\nЧучаева С.М.'.split('\n\n')
# )
#
# pprint.pprint(dic)
# for files in os.listdir('input'):
#     print(files)
#
# print(file.columns[5])

# def get_sheets_title(workbook):
#     return workbook.title
#
# none_sort_wb = openpyxl.load_workbook('output/prepodi_result.xlsx')
# print(none_sort_wb._sheets[0].title)
# none_sort_wb._sheets.sort(key=get_sheets_title)
# none_sort_wb.save('output/prepodi_result.xlsx')

# d = {'a': 21, 'b': 2135}
#
# del d['b']
#
# print(d)

def find_first_upper_symbol(string: str):
    for i, char in enumerate(string):
        if char.isupper():
            return i
    return -1

s = 'ASdsdАрхитектура интеграции и развертывания'
print(s.split('\n'))
print(s[find_first_upper_symbol(s):])

aud = 'ауд. G-234 (B-78)'
aud1 = 'комп. G-234 (B-78)'
aud0 = 'G-234'

def get_clear_auditory(auditory: str):
    if auditory.find('(') != -1:
        auditory = auditory[:auditory.find('(')-1]
    if auditory.find('ауд. ') != -1:
        auditory = auditory[auditory.find('ауд. ')+5:]
    if auditory.find('комп. ') != -1:
        auditory = auditory[auditory.find('комп. ')+6:]
    return auditory

print(get_clear_auditory(aud))
print(get_clear_auditory(aud0))
print(get_clear_auditory(aud1))
# def edit_d(d):
#     d['a'] = 124
#
# edit_d(d)
# print(d)

# fios = 'Глухов А.В.' \
# 'Прилипко В.А.' \
# 'Прилипко В.А.'

# file = pd.read_csv('csv/IIT_mag_2kurs_osen_2022_2023.csv')
# fios = file.iloc[59, 72]
#
# print(fios.split('\n'))












