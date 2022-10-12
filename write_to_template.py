import os

import pandas as pd
import openpyxl

import process_the_file
from process_the_file import process
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


def write(name):
    # wb_target = openpyxl.Workbook()
    # target_sheet = wb_target.create_sheet(..sheet_name..)
    #
    # wb_source = openpyxl.load_workbook(..path\\+\\file_name.., data_only = True)
    # source_sheet = wb_source[..sheet_name..]
    #
    # copy_sheet(source_sheet, target_sheet)
    #
    # if 'Sheet' in wb_target.sheetnames:  # remove default sheet
    #     wb_target.remove(wb_target['Sheet'])
    #
    # wb_target.save('out.xlsx')

    preparated_data = process(name)
    try:
        wb = openpyxl.load_workbook('output/prepodi_result.xlsx')
    except:
        wb = openpyxl.Workbook()

        # Удаление листа, создаваемого по умолчанию, при создании документа
        for sheet_name in wb.sheetnames:
            sheet = wb.get_sheet_by_name(sheet_name)
            wb.remove_sheet(sheet)

    template = openpyxl.load_workbook('templates/pattern_prepod.xlsx')['Шаблон']

    for fio, v in preparated_data.items():
        k = fio.replace('/', ' ')
        try:
            ws = wb.get_sheet_by_name(k)
        except:
            ws = wb.create_sheet(k)
            copy_sheet.copy_sheet(template, ws)

        for weekday, lessons in v.items():
            for num_lesson, chetnost in lessons.items():
                for num_chetnost, info in chetnost.items():
                    # print(k, weekday, num_lesson, num_chetnost, info)
                    line_num = 4+weekdays.get(weekday)*16+(int(num_lesson)-1)*2+parity.get(num_chetnost)
                    ws[f'E{line_num}'] = info.get('Предмет')
                    ws[f'F{line_num}'] = info.get('Вид занятий')
                    ws[f'G{line_num}'] = k
                    ws[f'H{line_num}'] = info.get('Аудитория')
                    ws[f'I{line_num}'] = info.get('Группа')
    wb.save('output/prepodi_result.xlsx')

# write('IIT_3-kurs_22_23_osen_07.10.2022.xlsx')


def write_all():
    for files in os.listdir('input'):
        if files.find('ITU_mag') != -1:
            process_the_file.COUNT_GROUPS_IN_ONE_PART = 3
            write(files)
        else:
            process_the_file.COUNT_GROUPS_IN_ONE_PART = 2
            write(files)

    def get_sheets_title(sheet):
        return sheet.title

    none_sort_wb = openpyxl.load_workbook('output/prepodi_result.xlsx')
    none_sort_wb._sheets.sort(key=get_sheets_title)
    none_sort_wb.save('output/prepodi_result.xlsx')

write_all()

