import pprint

import pandas as pd
import to_csv


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
            if prepod[weekday][num_of_lesson].get(chetnost) == None:
                prepod[weekday][num_of_lesson][chetnost] = {
                        'Предмет': predmet,
                        'Вид занятий': type_of_lesson,
                        'Аудитория': auditorium,
                        'Группа': group
                    }
            else:
                prepod[weekday][num_of_lesson][chetnost]['Группа'] += f'\n{group}'
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


def preparate_one_prepod(file, res, fio, index, part, CFNG):
    if res.get(fio) != None:
        res[fio] = add_lesson_to_prepod(res[fio], file.iloc[index, 0], str(file.iloc[index, 1]),
                                        str(file.iloc[index, CHETNOST_COL]),
                                        file.iloc[index, part + CFNG + PREDMET_COL], file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
                                        file.iloc[index, part + CFNG + AUDITORIUM_COL], file.columns[part + CFNG + GROUP_COL])
    else:
        res[fio] = {  # ФИО
            file.iloc[index, 0]: {  # День недели
                str(file.iloc[index, 1]): {  # Номер пары
                    str(file.iloc[index, CHETNOST_COL]):  # Четность недель
                        {
                            'Предмет': file.iloc[index, part + CFNG + PREDMET_COL],
                            'Вид занятий': file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
                            'Аудитория': file.iloc[index, part + CFNG + AUDITORIUM_COL],
                            'Группа': file.columns[part + CFNG + GROUP_COL]
                        }
                }
            }
        }


def preparate_data_prepods(file) -> dict:
    res = {}
    for part in range(0, len(file.columns), ONE_PART):
        for CFNG in range(0, (COUNT_GROUPS_IN_ONE_PART-1)*COLS_BEETWEN_GROUPS+1, COLS_BEETWEN_GROUPS): #Корректор для сдвига для следующих учебных групп в одном паттерне расписания
            fio_col = part + CFNG + FIO_COL
            if fio_col < len(file.columns):
                for index, fio in enumerate(file.iloc[:, fio_col]): # Проверить на наличие бага с совпадением расписания
                    if type(fio) == str and index > 0:
                        # if fio.find('\n\n') == -1 and fio.find(',') == -1:
                        #     preparate_one_prepod(file, res, fio, index, part, CFNG)
                        # elif fio.find(',') != -1:
                        #     fios = fio.split(',')
                        #     # print(fio, fios)
                        #     for local_fio in fios:
                        #         preparate_one_prepod(file, res, local_fio, index, part, CFNG)
                        # else:
                        #     fios = fio.split('\n\n')
                        #     for local_fio in fios:
                        #         preparate_one_prepod(file, res, local_fio, index, part, CFNG)
                        if fio.find(',') != -1:
                            fios = fio.split(',')
                            # print(fio, fios)
                            for local_fio in fios:
                                preparate_one_prepod(file, res, local_fio, index, part, CFNG)
                        elif fio.find('\n\n') != -1:
                            fios = fio.split('\n\n')
                            for local_fio in fios:
                                preparate_one_prepod(file, res, local_fio, index, part, CFNG)
                        elif fio.find('\n') != -1:
                            fios = fio.split('\n')
                            for local_fio in fios:
                                preparate_one_prepod(file, res, local_fio, index, part, CFNG)
                        else:
                            preparate_one_prepod(file, res, fio, index, part, CFNG)
    return res

def process_prepods(name):
    file = read_csv(name)
    return preparate_data_prepods(file)
# pprint.pprint(process('IIT_3-kurs_22_23_osen_07.10.2022.xlsx'))

#
def convert_prepods_to_auditory(prepods: dict):
    res = {}
    for name, weekdays in prepods.items():
        for weekday, lessons in weekdays.items():
            for num_lesson, chetnosti in lessons.items():
                for chet_nechet, content in chetnosti.items():
                    if type(content.get('Аудитория')) != float:
                        if content.get('Аудитория').find(',') != -1:
                            audis = content.get('Аудитория').split(',')
                            for auditory in audis:
                                if res.get(auditory) == None:
                                    res[auditory] = []
                                res[auditory].append({name: {weekday: {num_lesson: {chet_nechet: content}}}})
                        elif content.get('Аудитория').find('\n\n') != -1:
                            audis = content.get('Аудитория').split('\n\n')
                            for auditory in audis:
                                if res.get(auditory) == None:
                                    res[auditory] = []
                                res[auditory].append({name: {weekday: {num_lesson: {chet_nechet: content}}}})
                        elif content.get('Аудитория').find('\n') != -1:
                            audis = content.get('Аудитория').split('\n')
                            for auditory in audis:
                                if res.get(auditory) == None:
                                    res[auditory] = []
                                res[auditory].append({name: {weekday: {num_lesson: {chet_nechet: content}}}})
                        else:
                            if res.get(content.get('Аудитория')) == None:
                                res[content.get('Аудитория')] = []
                            res[content.get('Аудитория')].append({name: {weekday: {num_lesson: {chet_nechet: content}}}})
    return res










