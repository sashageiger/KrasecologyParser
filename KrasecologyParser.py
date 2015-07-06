# -*- coding: utf-8 -*-


import os, sys, json, re
import datetime
from time import sleep
import urllib.request as ulb

import xlsxwriter


#Часовой пояс Красноярск (GMT +7)
TIME_ZONE = 7

#Текущее время
CURRENT_DATE_TIME = datetime.datetime.now()

#Ширина колонки в экселе
COL_WIDTH = 20

#Путь к рабочему столу
DESKTOP_PATH = os.path.expanduser("~/Desktop/")

#Кодировка на сайте
CODING = 'utf-8'

#Интервал между запросами к сайту
DELAY = 0.5

#Временные интервалы день или неделя
INTERVAL = {
    1: 'day',
    2: 'week'
}

#Названия постов
PLACES = {
    1: 'Ачинск-Юго-Восточный',
    2: 'Красноярск-Северный',
    3: 'Красноярск-Березовка',
    4: 'Красноярск-Солнечный',
    5: 'Красноярск-Черемушки',
    6: 'Красноярск-Кубеково',
}

PLACE_URL = 'http://krasecology.ru/Main/GetAirSensorList/%s'

SENSOR_URL = 'http://krasecology.ru/Main/GetAirSensorData/%s?timelap=%s'


def parse_sensors():
    '''
    Проходит по ссылкам вида 'http://krasecology.ru/Main/GetAirSensorList/1'
    и получает список датчиков по каждому посту из заданных в PLACES.
    Возвращает структуру, содержащую распределение датчиков по постам:
    sp = {
        1: {
            1: {
                'N': 'Углерода оксид',
                'C': '337'
                'U': 'мг/м3'
            }
            2: ...
        ...
        }
        2: ...
    }
    '''

    sp = {i:{} for i in PLACES.keys()}

    print('\nПолучение списка датчиков по постам')
    for p in sp.keys():

        try:
            req = ulb.urlopen(PLACE_URL % p).read().decode(CODING)
            data_ = json.loads(req)
        except Exception as e:
            print('Не удалось загрузить данные от %s' % PLACE_URL % p)
            continue

        print('Пост "%s": %s датчиков' % (PLACES[p], len(data_)))
        i = 1
        for s in data_:
            if s['Name'] == 'Роза ветров':
                continue
            sp[p][i] = {'N':s['Name'], 'C':s['Code'], 'U':s['Unit']}
            i += 1
        sleep(DELAY)

    return sp


def parse_sensor_request(jsondata):
    '''
    На вход получает данные от датчика в формате json и возвращает
    список точек вида [ [ дата, значение], [ дата, значение], ... ].
    '''

    j = json.loads(jsondata)
    data_ = []
    for i in j['Data']:
        # перевод из формата UNIX в формат Excel
        time_ = (int(str(i['x'])[:-3]) + 3600 * TIME_ZONE) / 86400 + 25569
        data_.append([time_, i['y']])
    return data_


def parse_data(pl, sp, interval):
    '''
    На вход получает:
    1. перечеть постов в виде списка [1,2,3 ...];
    2. перечеть датчиков в виде структуры от parse_sensors;
    3. временной интервал для сканирования.
    Возвращает запись вида:
    {
        'I': 'Неделя',
        'D': '02.07.2015'
        'P': {
            'Ачинск-Юго-Восточный': {
                'Углерода оксид': [ 'мг/м3', [ дата, значение], ... ]
                'Серы диоксид': [ 'мг/м3', [ дата, значение], ... ]
                ...
            }
            ...
        }
    }
    '''

    data_ = {
        'I': dict(enumerate(['День', 'Неделя'], 1))[interval],
        'D': '',
        'P': {},
    }

    print('\nПолучение данных от датчиков.')
    for p in pl:
        cp = data_['P'][PLACES[p]] = {}
        for s in sorted(sp[p].keys()):
            cp[sp[p][s]['N']] = []
            try:
                url = SENSOR_URL % (sp[p][s]['C'], INTERVAL[interval])
                print('Пост "%s" датчик "%s"' % (PLACES[p], sp[p][s]['N']))
                req = ulb.urlopen(url).read().decode(CODING)
                cp[sp[p][s]['N']] = [sp[p][s]['U']] + parse_sensor_request(req)
            except:
                print('Не удалось получить данные от %s' % url)
                continue
            sleep(DELAY)

    return data_


def sup_sub(str_):
    '''
    Удаляет теги sup и sub из строки.
    '''

    rgx = re.compile(r'<su[pb]>(.*)</su[pb]>')
    for i in rgx.finditer(str_):
        s = i.start()
        e = i.end()
        txt = i.group(1)
        str_ = ''.join([str_[:s], txt, str_[e:]])
    return str_


def make_xlsx(data_):
    '''
    На вход получает данные от функции parse_data,
    создает отчет в формате эксель в папке DESKTOP_PATH.
    В каждой вкладке отчета помещаются данные по одному из постов.
    '''

    full_path = (
        os.path.join(DESKTOP_PATH, 'krasecology.ru %s.xlsx') % \
        CURRENT_DATE_TIME.strftime('%d.%m.%Y (%H.%M)')).replace('/', '\\')

    print('\nСохранение данных в книгу %s.' % full_path)
    workbook = xlsxwriter.Workbook(full_path)

    header_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#d3d3d3',
            'border': 1,
        })

    val_format = workbook.add_format(
        {
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
        })

    date_format = workbook.add_format(
        {
            'align': 'center',
            'valign': 'vcenter',
            'num_format': 'dd.mm.yyyy hh:mm',
            'border': 1,
        })

    for p in sorted(data_['P'].keys()):
        worksheet = workbook.add_worksheet(p)
        sens = data_['P'][p]

        row, col = 0, 0
        worksheet.write(
            row, col,
            'Дата сканирования: ' + CURRENT_DATE_TIME.strftime('%d.%m.%Y (%H.%M)')
        ); row += 2

        for s in sorted(sens.keys()):
            worksheet.merge_range(row, col, row, col + 1, s, header_format)
            row += 1

            if not sens[s]:
                worksheet.merge_range(
                    row, col, row, col + 1, 'Не удалось получить данные', header_format)
                worksheet.set_column(col, col + 1, COL_WIDTH)
                col += 3
                row = 2
                continue

            worksheet.write(row, col, 'Время измерения', header_format)
            worksheet.write_rich_string(row, col + 1, sup_sub(sens[s][0]), header_format)
            row += 1

            for val in sens[s][1:]:
                worksheet.write(row, col, val[0], date_format)
                worksheet.write(row, col + 1, val[1], val_format)
                row += 1

            worksheet.set_column(col, col + 1, COL_WIDTH)
            col += 3
            row = 2

    workbook.close()


if __name__ == '__main__':
    print(''.join(['\n', '='*44, '\n',
                   ' Парсер http://krasecology.ru/operative/air\n',
                   '='*44, '\n']))
    while 1:
        i = input('Выберите временной интервал для получения данных.\n'
                  'Введите:\n 1 - для получения данных за день,\n 2 - для '
                  'получения данных за неделю.\nНажмите Enter для продолжения.'
                  ' Для выхода введите q.\n>> ')
        if i in ['1', '2']:
            break
        elif i == 'q':
            sys.exit()

    sp = parse_sensors()
    dt = parse_data(sorted(PLACES.keys()), sp, int(i))
    make_xlsx(dt)

    print('\nЗавершение.')
    input('\nСканирование завершено.\nНажмите любую клавишу для выхода.')