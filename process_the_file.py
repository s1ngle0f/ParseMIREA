import pprint

import pandas as pd
import to_csv


def read_csv(name):
    to_csv.create_csv_from_excel(name)
    path = f'csv/{name.replace(".xlsx", ".csv")}' if name.find('.xlsx') else f'csv/{name}'
    return pd.read_csv(path)
# print(read_csv('IIT_3-kurs_22_23_osen_07.10.2022_non_processed.csv').head(10))


# Пример одного преподавателя из массива
# { 'Зубков М. В.': {
#   'ПОНЕДЕЛЬНИК': {
#       '1': {
#           '|': {'Предмет': 'Моделирование', 'Вид занятий': 'пр', 'Аудитория': 'Г-227-2', 'Группа': 'ИКБО-30-20'},
#           '||': {'Предмет': 'Моделирование', 'Вид занятий': 'пр', 'Аудитория': 'Г-227-2', 'Группа': 'ИКБО-30-20'} #Не обязательно, тк может быть либо одно, либо другое
#       }
#   }
# }}


def add_lesson_to_prepod(prepod: dict, weekday, num_of_lesson, chetnost, predmet, type_of_lesson, auditorium, group) -> dict:
    if prepod.get(weekday) != None:
        if prepod[weekday].get(num_of_lesson) != None: #ТУТ СДЕЛАТЬ ПРОВЕРКУ НА СОВПАДЕНИЕ ГРУПП!!!
            prepod[weekday][num_of_lesson][chetnost] = {
                    'Предмет': predmet,
                    'Вид занятий': type_of_lesson,
                    'Аудитория': auditorium,
                    'Группа': group
                }
        else:
            prepod[weekday][num_of_lesson] = {
                chetnost:
                    {
                        'Предмет': predmet,
                        'Вид занятий': type_of_lesson,
                        'Аудитория': auditorium,
                        'Группа': group
                    }
            }
    else:
        prepod[weekday] = {
            num_of_lesson: {
                chetnost:
                    {
                        'Предмет': predmet,
                        'Вид занятий': type_of_lesson,
                        'Аудитория': auditorium,
                        'Группа': group
                    }
            }
        }
    return prepod


def preparate_data(file) -> dict:
    res = {}
    for part in range(0, len(file.columns), 15):
        for CFNG in range(0, 6, 5): #Корректор для сдвига для следующих учебных групп в одном паттерне расписания
            fio_col = part + CFNG + 7
            if fio_col < len(file.columns):
                for index, fio in enumerate(file.iloc[:, fio_col]): # Проверить на наличие бага с совпадением расписания
                    if type(fio) == str and index > 0:
                        if fio.find('\n\n') == -1 and fio.find(',') == -1:
                            if res.get(fio) != None:
                                res[fio] = add_lesson_to_prepod(res[fio], file.iloc[index, 0], str(file.iloc[index, 1]), str(file.iloc[index, 4]),
                                                    file.iloc[index, part + CFNG + 5], file.iloc[index, part + CFNG + 6], file.iloc[index, part + CFNG + 8], file.columns[part + CFNG + 5])
                            else:
                                res[fio] = { # ФИО
                                    file.iloc[index, 0]: { # День недели
                                        str(file.iloc[index, 1]): { # Номер пары
                                            str(file.iloc[index, 4]): # Четность недель
                                                {
                                                    'Предмет': file.iloc[index, part + CFNG + 5],
                                                    'Вид занятий': file.iloc[index, part + CFNG + 6],
                                                    'Аудитория': file.iloc[index, part + CFNG + 8],
                                                    'Группа': file.columns[part + CFNG + 5]
                                                }
                                        }
                                    }
                                }
                        elif fio.find(',') != -1:
                            fios = fio.split(',')
                            # print(fio, fios)
                            for local_fio in fios:
                                if res.get(local_fio) != None:
                                    res[local_fio] = add_lesson_to_prepod(res[local_fio], file.iloc[index, 0],
                                                                          str(file.iloc[index, 1]),
                                                                          str(file.iloc[index, 4]),
                                                                          file.iloc[index, part + CFNG + 5],
                                                                          file.iloc[index, part + CFNG + 6],
                                                                          file.iloc[index, part + CFNG + 8],
                                                                          file.columns[part + CFNG + 5])
                                else:
                                    res[local_fio] = {  # ФИО
                                        file.iloc[index, 0]: {  # День недели
                                            str(file.iloc[index, 1]): {  # Номер пары
                                                str(file.iloc[index, 4]):  # Четность недель
                                                    {
                                                        'Предмет': file.iloc[index, part + CFNG + 5],
                                                        'Вид занятий': file.iloc[index, part + CFNG + 6],
                                                        'Аудитория': file.iloc[index, part + CFNG + 8],
                                                        'Группа': file.columns[part + CFNG + 5]
                                                    }
                                            }
                                        }
                                    }
                        else:
                            fios = fio.split('\n\n')
                            for local_fio in fios:
                                if res.get(local_fio) != None:
                                    res[local_fio] = add_lesson_to_prepod(res[local_fio], file.iloc[index, 0],
                                                                    str(file.iloc[index, 1]), str(file.iloc[index, 4]),
                                                                    file.iloc[index, part + CFNG + 5],
                                                                    file.iloc[index, part + CFNG + 6],
                                                                    file.iloc[index, part + CFNG + 8],
                                                                    file.columns[part + CFNG + 5])
                                else:
                                    res[local_fio] = {  # ФИО
                                        file.iloc[index, 0]: {  # День недели
                                            str(file.iloc[index, 1]): {  # Номер пары
                                                str(file.iloc[index, 4]):  # Четность недель
                                                    {
                                                        'Предмет': file.iloc[index, part + CFNG + 5],
                                                        'Вид занятий': file.iloc[index, part + CFNG + 6],
                                                        'Аудитория': file.iloc[index, part + CFNG + 8],
                                                        'Группа': file.columns[part + CFNG + 5]
                                                    }
                                            }
                                        }
                                    }
    return res

def process(name):
    file = read_csv(name)
    return preparate_data(file)
# pprint.pprint(process('IIT_3-kurs_22_23_osen_07.10.2022.xlsx'))