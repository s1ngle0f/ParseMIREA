import os
import pprint
import openpyxl
import pandas as pd


# file = pd.read_csv('csv/IIT_3-kurs_22_23_osen_07.10.2022.csv')
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

none_sort_wb = openpyxl.load_workbook('output/prepodi_result.xlsx')
print(none_sort_wb._sheets[0].title)
none_sort_wb._sheets.sort()
