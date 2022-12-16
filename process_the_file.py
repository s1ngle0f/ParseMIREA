import pprint

import pandas as pd
import to_csv
import re

ONE_PART = 15
COUNT_GROUPS_IN_ONE_PART = 2
COLS_BEETWEN_GROUPS = 5
FIO_COL = 7
CHETNOST_COL = 4
PREDMET_COL = 5
TYPE_OF_LESSON_COL = 6
AUDITORIUM_COL = 8
GROUP_COL = 5

def read_csv(name):
    if name.find('.xlsx') != -1:
        path = f'csv/{name.replace(".xlsx", ".csv")}'
        to_csv.create_csv_from_excel(name)
    else:
        path = f'csv/{name}'
    print(path)
    return pd.read_csv(path)
# print(read_csv('IIT_3-kurs_22_23_osen_07.10.2022_non_processed.csv').head(10))

def get_by(sort_name, arr):
    res = {}
    # for i, el in enumerate(arr):
    #     if res.get(el.get(sort_name)) == None:
    #         res[el.get(sort_name)] = []
    #     res[el.get(sort_name)].append(el)
    #     # del res[el.get(sort_name)][-1][sort_name]
    ban_symbols = [',', '\n\n', '\n', ', ']
    for i, el in enumerate(arr):
        for ban_symbol in ban_symbols:
            if ban_symbol in el.get(sort_name):
                # names = el.get(sort_name).split(ban_symbol)
                names = re.split(",|\n\n|\n|, ", el.get(sort_name))
                for name in names:
                    if res.get(get_clear_fio(name)) == None:
                        res[get_clear_fio(name)] = []
                    res[get_clear_fio(name)].append(el)
        if not any(ban_symbol in el.get(sort_name) for ban_symbol in ban_symbols):
            if res.get(get_clear_fio(el.get(sort_name))) == None:
                res[get_clear_fio(el.get(sort_name))] = []
            res[get_clear_fio(el.get(sort_name))].append(el)
    return res

def add_to_proccess_result(res, fio, weekday, num_of_lesson, chetnost, lesson, vid_zanyatiy, auditory, group):
    res.append({
        'ФИО': fio,
        'День недели': weekday,
        'Номер пары': num_of_lesson,
        'Четность недели': chetnost,
        'Предмет': lesson,
        'Вид занятий': vid_zanyatiy,
        'Аудитория': auditory,
        'Группа': group
    })

def process(name):
    file = read_csv(name)
    res = []
    ban_symbols = [',', '\n\n', '\n', ', ']
    for part in range(0, len(file.columns), ONE_PART):
        for CFNG in range(0, (COUNT_GROUPS_IN_ONE_PART-1)*COLS_BEETWEN_GROUPS+1, COLS_BEETWEN_GROUPS): #Корректор для сдвига для следующих учебных групп в одном паттерне расписания
            fio_col = part + CFNG + FIO_COL
            if fio_col < len(file.columns):
                for index, fio in enumerate(file.iloc[:, fio_col]): # Проверить на наличие бага с совпадением расписания
                    if type(fio) == str and index > 0:
                        for ban_symbol in ban_symbols:
                            if ban_symbol in fio:
                                fios = fio.split(ban_symbol)
                                for local_fio in fios:
                                    add_to_proccess_result(res=res,
                                       fio=get_clear_fio(local_fio),
                                       # fio=fio,
                                       weekday=str(file.iloc[index, 0]),
                                       num_of_lesson=str(file.iloc[index, 1]),
                                       chetnost=str(file.iloc[index, CHETNOST_COL]),
                                       lesson=file.iloc[index, part + CFNG + PREDMET_COL],
                                       vid_zanyatiy=file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
                                       auditory=get_clear_auditory(file.iloc[index, part + CFNG + AUDITORIUM_COL]),
                                       group=file.columns[part + CFNG + GROUP_COL]
                                   )
                        if not any(ban_symbol in fio for ban_symbol in ban_symbols):
                            # print(f'Одна фамилия! {fio}')
                            add_to_proccess_result(res = res,
                                fio = get_clear_fio(fio),
                                weekday = str(file.iloc[index, 0]),
                                num_of_lesson = str(file.iloc[index, 1]),
                                chetnost = str(file.iloc[index, CHETNOST_COL]),
                                lesson = file.iloc[index, part + CFNG + PREDMET_COL],
                                vid_zanyatiy = file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
                                auditory = get_clear_auditory(file.iloc[index, part + CFNG + AUDITORIUM_COL]),
                                group = file.columns[part + CFNG + GROUP_COL]
                            )
                        #     while get_n_upper_symbol(fio) > 0:
                        #         add_to_proccess_result(res=res,
                        #                                fio=get_clear_fio(fio[:find_n_upper_symbol(fio, 2) + 2]),
                        #                                weekday=str(file.iloc[index, 0]),
                        #                                num_of_lesson=str(file.iloc[index, 1]),
                        #                                chetnost=str(file.iloc[index, CHETNOST_COL]),
                        #                                lesson=file.iloc[index, part + CFNG + PREDMET_COL],
                        #                                vid_zanyatiy=file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
                        #                                auditory=get_clear_auditory(file.iloc[index, part + CFNG + AUDITORIUM_COL]),
                        #                                group=file.columns[part + CFNG + GROUP_COL]
                        #                                )
                        #         fio = fio[find_n_upper_symbol(fio, 2) + 2:]
                        # else:
                        #     print(f'Много фамилий! {fio}')
    return res
# pprint.pprint(process('IIT_3-kurs_22_23_osen_07.10.2022.xlsx'))

def get_clear_fio(fio: str):
    fio = fio.replace(' ', '')
    return f'{fio[:find_n_upper_symbol(fio, 1)]} {fio[find_n_upper_symbol(fio, 1):]}'

def get_clear_auditory(auditory: str):
    # if auditory.find('(') != -1:
    #     auditory = auditory[:auditory.find('(')-1]
    ban_symbols = ['ауд. ', 'лаб. ', 'комп. ', '(В-78)']
    for ban_symbol in ban_symbols:
        if ban_symbol in auditory:
            auditory.replace(ban_symbol, '')
    return auditory

def get_n_upper_symbol(string: str):
    counter = 0
    for i, char in enumerate(string):
        if char.isupper():
            counter += 1
    return counter

def find_n_upper_symbol(string: str, n: int = 0):
    counter = 0
    for i, char in enumerate(string):
        if char.isupper():
            if n == counter:
                return i
            else:
                counter += 1
    return 0

def process_dirty(name):
    file = read_csv(name)
    res = []
    ban_symbols = [',', '\n\n', '\n', ', ']
    for part in range(0, len(file.columns), ONE_PART):
        for CFNG in range(0, (COUNT_GROUPS_IN_ONE_PART-1)*COLS_BEETWEN_GROUPS+1, COLS_BEETWEN_GROUPS): #Корректор для сдвига для следующих учебных групп в одном паттерне расписания
            fio_col = part + CFNG + FIO_COL
            if fio_col < len(file.columns):
                for index, fio in enumerate(file.iloc[:, fio_col]): # Проверить на наличие бага с совпадением расписания
                    if type(fio) == str and index > 0:
                        add_to_proccess_result(res=res,
                           # fio=get_clear_fio(local_fio),
                           fio=fio,
                           weekday=str(file.iloc[index, 0]),
                           num_of_lesson=str(file.iloc[index, 1]),
                           chetnost=str(file.iloc[index, CHETNOST_COL]),
                           lesson=file.iloc[index, part + CFNG + PREDMET_COL],
                           vid_zanyatiy=file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
                           auditory=get_clear_auditory(file.iloc[index, part + CFNG + AUDITORIUM_COL]),
                           group=file.columns[part + CFNG + GROUP_COL]
                        )
                        if not any(ban_symbol in fio for ban_symbol in ban_symbols):
                            # print(f'Одна фамилия! {fio}')
                            add_to_proccess_result(res = res,
                                fio = get_clear_fio(fio),
                                weekday = str(file.iloc[index, 0]),
                                num_of_lesson = str(file.iloc[index, 1]),
                                chetnost = str(file.iloc[index, CHETNOST_COL]),
                                lesson = file.iloc[index, part + CFNG + PREDMET_COL],
                                vid_zanyatiy = file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
                                auditory = get_clear_auditory(file.iloc[index, part + CFNG + AUDITORIUM_COL]),
                                group = file.columns[part + CFNG + GROUP_COL]
                            )
                        #     while get_n_upper_symbol(fio) > 0:
                        #         add_to_proccess_result(res=res,
                        #                                fio=get_clear_fio(fio[:find_n_upper_symbol(fio, 2) + 2]),
                        #                                weekday=str(file.iloc[index, 0]),
                        #                                num_of_lesson=str(file.iloc[index, 1]),
                        #                                chetnost=str(file.iloc[index, CHETNOST_COL]),
                        #                                lesson=file.iloc[index, part + CFNG + PREDMET_COL],
                        #                                vid_zanyatiy=file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
                        #                                auditory=get_clear_auditory(file.iloc[index, part + CFNG + AUDITORIUM_COL]),
                        #                                group=file.columns[part + CFNG + GROUP_COL]
                        #                                )
                        #         fio = fio[find_n_upper_symbol(fio, 2) + 2:]
                        # else:
                        #     print(f'Много фамилий! {fio}')
    return res
# pprint.pprint(process('IIT_3-kurs_22_23_osen_07.10.2022.xlsx'))
