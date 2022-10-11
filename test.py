import pprint

import pandas as pd


file = pd.read_csv('csv/IIT_3-kurs_22_23_osen_07.10.2022.csv')
dic = {}
print(
    'Шошников И.К.\n\nЧучаева С.М.'.split('\n\n')
)

pprint.pprint(dic)

print(file.columns[5])