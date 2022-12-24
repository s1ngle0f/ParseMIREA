import os
import pprint

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import process_the_file
from process_the_file import process, get_by, process_dirty, process_exams
import copy_sheet


weekdays = {
    'ПОНЕДЕЛЬНИК': 0,
    'ВТОРНИК': 1,
    'СРЕДА': 2,
    'ЧЕТВЕРГ': 3,
    'ПЯТНИЦА': 4,
    'СУББОТА': 5,
    'ВОСКРЕСЕНЬЕ': 6
}

parity = {
    'I': 0,
    'II': 1
}

ban_symbols = [',', '\n\n', '\n', ', ']

def write(sort_by, name, result_filename):
    # preparated_data = get_by(sort_by, process(name))
    preparated_data = get_by(sort_by, process_dirty(name))
    try:
        wb = openpyxl.load_workbook(f'output/{result_filename}.xlsx')
    except:
        wb = openpyxl.Workbook()

        # Удаление листа, создаваемого по умолчанию, при создании документа
        for sheet_name in wb.sheetnames:
            sheet = wb.get_sheet_by_name(sheet_name)
            wb.remove_sheet(sheet)

    template = openpyxl.load_workbook('templates/pattern_prepod.xlsx')['Шаблон']

    for key, val in preparated_data.items():
        k = key.replace('/', ' ')
        try:
            ws = wb.get_sheet_by_name(k)
        except:
            ws = wb.create_sheet(k)
            copy_sheet.copy_sheet(template, ws)

        for info in val:
            # print(k, weekday, num_lesson, num_chetnost, info)
            line_num = 4+weekdays.get(info.get('День недели'))*16+(int(info.get('Номер пары'))-1)*2+parity.get(info.get('Четность недели'))
            ws[f'E{line_num}'] = info.get('Предмет')
            ws[f'F{line_num}'] = info.get('Вид занятий')
            ws[f'G{line_num}'] = info.get('ФИО')
            ws[f'H{line_num}'] = info.get('Аудитория')
            ws[f'I{line_num}'] = info.get('Группа')

            if any(ban_symbol in info.get('ФИО') for ban_symbol in ban_symbols):#Только для определения коллизий ФИО
                ws[f'J{line_num}'].fill = PatternFill(start_color="ff0000", end_color="ff0000", fill_type="solid")

    wb.save(f'output/{result_filename}.xlsx')

def write_exams(sort_by, name, result_filename):
    # preparated_data = get_by(sort_by, process(name))
    preparated_data = get_by(sort_by, process_exams(name))
    try:
        wb = openpyxl.load_workbook(f'output/{result_filename}.xlsx')
    except:
        wb = openpyxl.Workbook()

        # Удаление листа, создаваемого по умолчанию, при создании документа
        for sheet_name in wb.sheetnames:
            sheet = wb.get_sheet_by_name(sheet_name)
            wb.remove_sheet(sheet)

    template = openpyxl.load_workbook('templates/pattern_prepod_exams.xlsx')['Шаблон']

    for key, val in preparated_data.items():
        k = key.replace('/', ' ')
        try:
            ws = wb.get_sheet_by_name(k)
        except:
            ws = wb.create_sheet(k)
            copy_sheet.copy_sheet(template, ws)

        for info in val:
            # print(k, weekday, num_lesson, num_chetnost, info)
            line_num = 4+weekdays.get(info.get('День недели'))*16+(int(info.get('Номер пары'))-1)*2+parity.get(info.get('Четность недели'))
            ws[f'E{line_num}'] = info.get('Предмет')
            ws[f'F{line_num}'] = info.get('Вид занятий')
            ws[f'G{line_num}'] = info.get('ФИО')
            ws[f'H{line_num}'] = info.get('Аудитория')
            ws[f'I{line_num}'] = info.get('Группа')

            # if any(ban_symbol in info.get('ФИО') for ban_symbol in ban_symbols):#Только для определения коллизий ФИО
            #     ws[f'J{line_num}'].fill = PatternFill(start_color="ff0000", end_color="ff0000", fill_type="solid")

    wb.save(f'output/{result_filename}.xlsx')

# write('IIT_3-kurs_22_23_osen_07.10.2022.xlsx')


def write_all(sort_by, result_filename):
    for files in os.listdir('input'):
        if files.find('ITU_mag') != -1:
            process_the_file.COUNT_GROUPS_IN_ONE_PART = 3
            write(sort_by, files, result_filename)
        else:
            process_the_file.COUNT_GROUPS_IN_ONE_PART = 2
            write(sort_by, files, result_filename)

    def get_sheets_title(sheet):
        return sheet.title

    none_sort_wb = openpyxl.load_workbook(f'output/{result_filename}.xlsx')
    none_sort_wb._sheets.sort(key=get_sheets_title)
    none_sort_wb.save(f'output/{result_filename}.xlsx')

########################################################################

def write_auditory(name):
    preparated_data_prepods = process_prepods(name)
    preparated_data = process_the_file.convert_prepods_to_auditory(preparated_data_prepods)
    # pprint.pprint(preparated_data)
    try:
        wb = openpyxl.load_workbook('output/auditory_result.xlsx')
    except:
        wb = openpyxl.Workbook()

        # Удаление листа, создаваемого по умолчанию, при создании документа
        for sheet_name in wb.sheetnames:
            sheet = wb.get_sheet_by_name(sheet_name)
            wb.remove_sheet(sheet)

    template = openpyxl.load_workbook('templates/pattern_prepod.xlsx')['Шаблон']

    for auditory_name, prepods in preparated_data.items():
        # print(auditory_name, prepods)
        if type(auditory_name) != float:
            k = auditory_name.replace('/', ' ')
            try:
                ws = wb.get_sheet_by_name(k)
            except:
                ws = wb.create_sheet(k)
                copy_sheet.copy_sheet(template, ws)
            for prepod_content in prepods:
                for fio, v in prepod_content.items():
                    for weekday, lessons in v.items():
                        for num_lesson, chetnost in lessons.items():
                            for num_chetnost, info in chetnost.items():
                                # print(k, weekday, num_lesson, num_chetnost, info)
                                line_num = 4+weekdays.get(weekday)*16+(int(num_lesson)-1)*2+parity.get(num_chetnost)
                                ws[f'E{line_num}'] = info.get('Предмет')
                                ws[f'F{line_num}'] = info.get('Вид занятий')
                                ws[f'G{line_num}'] = fio
                                ws[f'H{line_num}'] = info.get('Аудитория')
                                ws[f'I{line_num}'] = info.get('Группа')
    wb.save('output/auditory_result.xlsx')

# write('IIT_3-kurs_22_23_osen_07.10.2022.xlsx')


def write_all_auditories():
    for files in os.listdir('input'):
        if files.find('ITU_mag') != -1:
            process_the_file.COUNT_GROUPS_IN_ONE_PART = 3
            write_auditory(files)
        else:
            process_the_file.COUNT_GROUPS_IN_ONE_PART = 2
            write_auditory(files)

    def get_sheets_title(sheet):
        return sheet.title

    none_sort_wb = openpyxl.load_workbook('output/auditory_result.xlsx')
    none_sort_wb._sheets.sort(key=get_sheets_title)
    none_sort_wb.save('output/auditory_result.xlsx')

########################################################################

def write_lesson(name):
    preparated_data_prepods = process_prepods(name)
    preparated_data = process_the_file.convert_prepods_to(preparated_data_prepods, 'Предмет')
    # pprint.pprint(preparated_data)
    try:
        wb = openpyxl.load_workbook('output/lesson_result.xlsx')
    except:
        wb = openpyxl.Workbook()

        # Удаление листа, создаваемого по умолчанию, при создании документа
        for sheet_name in wb.sheetnames:
            sheet = wb.get_sheet_by_name(sheet_name)
            wb.remove_sheet(sheet)

    template = openpyxl.load_workbook('templates/pattern_prepod.xlsx')['Шаблон']

    for auditory_name, prepods in preparated_data.items():
        # print(auditory_name, prepods)
        if type(auditory_name) != float:
            k = auditory_name.replace('/', ' ')
            try:
                ws = wb.get_sheet_by_name(k)
            except:
                ws = wb.create_sheet(k)
                copy_sheet.copy_sheet(template, ws)
            for prepod_content in prepods:
                for fio, v in prepod_content.items():
                    for weekday, lessons in v.items():
                        for num_lesson, chetnost in lessons.items():
                            for num_chetnost, info in chetnost.items():
                                # print(k, weekday, num_lesson, num_chetnost, info)
                                line_num = 4+weekdays.get(weekday)*16+(int(num_lesson)-1)*2+parity.get(num_chetnost)
                                ws[f'E{line_num}'] = info.get('Предмет')
                                ws[f'F{line_num}'] = info.get('Вид занятий')
                                ws[f'G{line_num}'] = fio
                                ws[f'H{line_num}'] = info.get('Аудитория')
                                ws[f'I{line_num}'] = info.get('Группа')
    wb.save('output/lesson_result.xlsx')

# write('IIT_3-kurs_22_23_osen_07.10.2022.xlsx')


def write_all_lessons():
    for files in os.listdir('input'):
        if files.find('ITU_mag') != -1:
            process_the_file.COUNT_GROUPS_IN_ONE_PART = 3
            write_lesson(files)
        else:
            process_the_file.COUNT_GROUPS_IN_ONE_PART = 2
            write_lesson(files)

    def get_sheets_title(sheet):
        return sheet.title

    none_sort_wb = openpyxl.load_workbook('output/lesson_result.xlsx')
    none_sort_wb._sheets.sort(key=get_sheets_title)
    none_sort_wb.save('output/lesson_result.xlsx')
