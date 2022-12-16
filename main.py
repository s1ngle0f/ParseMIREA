import write_to_template
import json
import time
import threading

#  'Аудитория'
#  'Вид занятий'
#  'Группа'
#  'День недели'
#  'Номер пары'
#  'Предмет'
#  'ФИО'
#  'Четность недели'

if __name__ == '__main__':
    start_time = time.time()

    with open('settings.json') as f:
        settings = json.load(f)
        if settings['prepods']:
            write_to_template.write_all('ФИО', 'prepods_result')
        if settings['auditories']:
            write_to_template.write_all('Аудитория', 'auditory_result')
        if settings['lessons']:
            write_to_template.write_all('Предмет', 'lessons_result')

    print("--- %s seconds ---" % (round(time.time() - start_time, 2)))
