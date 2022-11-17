import write_to_template
import json

if __name__ == '__main__':
    with open('settings.json') as f:
        settings = json.load(f)
        if settings['prepods']:
            write_to_template.write_all_prepods()
            # print('prepods')
        if settings['auditories']:
            write_to_template.write_all_auditories()
            # print('auditories')
        if settings['lessons']:
            write_to_template.write_all_lessons()
            # print('lessons')

