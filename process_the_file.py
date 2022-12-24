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

########EXAMS
STATIC_COLUMNS = [0, 1]
ALTERATION_COLUMN = 1
END_LINE = 64
conf = {
    'height': 3,
    'pattern': {
        'Группа': [0, 0]
    },
    'block': {
        'Вид занятий': [0, 0],  # row col
        'Предмет': [1, 0],
        'ФИО': [2, 0],
        'Время': [0, 1],
        'Аудитория': [0, 2]
    },
    'notnull': [0, 0]
}
COL_NAMES = [re.compile("[А-Я]{4}-[0-9]{2}-[0-9]{2}"), 'время', '№                      ауд', 'Ссылка']

def read_csv(name, is_fill_free_space = True):
    if name.find('.xlsx') != -1:
        path = f'csv/{name.replace(".xlsx", ".csv")}'
        to_csv.create_csv_from_excel(name, is_fill_free_space)
    else:
        path = f'csv/{name}'
    print(path)
    # return pd.read_csv(path) #До парса экзаменов
    return pd.read_csv(path, header=None)
# print(read_csv('IIT_3-kurs_22_23_osen_07.10.2022_non_processed.csv').head(10))

def get_by(sort_name, arr):
    res = {}
    ban_symbols = [',', '\n\n', '\n', ', ']
    for i, el in enumerate(arr):
        # print(f'{el.get(sort_name)} {type(el.get(sort_name))}')
        if type(el.get(sort_name)) != float:
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

def get_first_pattern(file, col_names, line, cut = True):
    res = []
    for col_index in range(0, len(file.columns) - len(col_names) + 1):
        for i, name in enumerate(col_names):
            # print(file.iloc[line, col_index + i])
            if type(name) == re.Pattern:
                if name.search(str(file.iloc[line, col_index + i])) == None:
                    break
            else:
                if file.iloc[line, col_index + i] != name:
                    break
            if i == len(col_names) - 1:
                # print(file.iloc[line, col_index: col_index + i + 1])
                res.append(file.iloc[:, col_index: col_index + i + 1])
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

def process_exams(name):
    # file = read_csv(name, False)
    file = read_csv(name, False).iloc[:END_LINE, :]
    for col in STATIC_COLUMNS:
        file[col] = file[col].ffill()
    # print(file.iloc[:20, :2])
    patterns = get_first_pattern(file, COL_NAMES, 0)
    res = []

    for pattern in patterns:
        # print(pattern.iloc[:, 0])
        # print(pattern.iloc[conf.get('pattern').get('group')[0], conf.get('pattern').get('group')[1]])
        for index, alteration in enumerate(file.iloc[1:, ALTERATION_COLUMN]):
            if file.iloc[index, ALTERATION_COLUMN] != alteration:
                block = pattern.iloc[index+1:index+1+conf.get('height'), :]
                # print(block.iloc[:, :])
                # print(type(block.iloc[conf.get('notnull')[0], conf.get('notnull')[1]]))
                # print(block.iloc[conf.get('notnull')[0], conf.get('notnull')[1]])
                if type(block.iloc[conf.get('notnull')[0], conf.get('notnull')[1]]) != float:
                    block_res = {}
                    group = pattern.iloc[conf.get('pattern').get('Группа')[0], conf.get('pattern').get('Группа')[1]]
                    block_res['Группа'] = group
                    block_res['День'] = alteration[:2]
                    for k, v in conf.get('block').items():
                        # print(block.iloc[v[0], v[1]])
                        block_res[k] = block.iloc[v[0], v[1]]
                    res.append(block_res)

    # for part in range(0, len(file.columns), ONE_PART):
    #     for CFNG in range(0, (COUNT_GROUPS_IN_ONE_PART-1)*COLS_BEETWEN_GROUPS+1, COLS_BEETWEN_GROUPS): #Корректор для сдвига для следующих учебных групп в одном паттерне расписания
    #         fio_col = part + CFNG + FIO_COL
    #         if fio_col < len(file.columns):
    #             for index, fio in enumerate(file.iloc[:, fio_col]): # Проверить на наличие бага с совпадением расписания
    #                 if type(fio) == str and index > 0:
    #                     for ban_symbol in ban_symbols:
    #                         if ban_symbol in fio:
    #                             fios = fio.split(ban_symbol)
    #                             for local_fio in fios:
    #                                 add_to_proccess_result(res=res,
    #                                    fio=get_clear_fio(local_fio),
    #                                    # fio=fio,
    #                                    weekday=str(file.iloc[index, 0]),
    #                                    num_of_lesson=str(file.iloc[index, 1]),
    #                                    chetnost=str(file.iloc[index, CHETNOST_COL]),
    #                                    lesson=file.iloc[index, part + CFNG + PREDMET_COL],
    #                                    vid_zanyatiy=file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
    #                                    auditory=get_clear_auditory(file.iloc[index, part + CFNG + AUDITORIUM_COL]),
    #                                    group=file.columns[part + CFNG + GROUP_COL]
    #                                )
    #                     if not any(ban_symbol in fio for ban_symbol in ban_symbols):
    #                         # print(f'Одна фамилия! {fio}')
    #                         add_to_proccess_result(res = res,
    #                             fio = get_clear_fio(fio),
    #                             weekday = str(file.iloc[index, 0]),
    #                             num_of_lesson = str(file.iloc[index, 1]),
    #                             chetnost = str(file.iloc[index, CHETNOST_COL]),
    #                             lesson = file.iloc[index, part + CFNG + PREDMET_COL],
    #                             vid_zanyatiy = file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
    #                             auditory = get_clear_auditory(file.iloc[index, part + CFNG + AUDITORIUM_COL]),
    #                             group = file.columns[part + CFNG + GROUP_COL]
    #                         )
    #                     #     while get_n_upper_symbol(fio) > 0:
    #                     #         add_to_proccess_result(res=res,
    #                     #                                fio=get_clear_fio(fio[:find_n_upper_symbol(fio, 2) + 2]),
    #                     #                                weekday=str(file.iloc[index, 0]),
    #                     #                                num_of_lesson=str(file.iloc[index, 1]),
    #                     #                                chetnost=str(file.iloc[index, CHETNOST_COL]),
    #                     #                                lesson=file.iloc[index, part + CFNG + PREDMET_COL],
    #                     #                                vid_zanyatiy=file.iloc[index, part + CFNG + TYPE_OF_LESSON_COL],
    #                     #                                auditory=get_clear_auditory(file.iloc[index, part + CFNG + AUDITORIUM_COL]),
    #                     #                                group=file.columns[part + CFNG + GROUP_COL]
    #                     #                                )
    #                     #         fio = fio[find_n_upper_symbol(fio, 2) + 2:]
    #                     # else:
    #                     #     print(f'Много фамилий! {fio}')
    return res
# pprint.pprint(process_exams('IIT_1-kurs_2022_2023_zima (1).xlsx'))
# process_exams('IIT_1-kurs_2022_2023_zima (1).xlsx')

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
