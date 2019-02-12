import ats
from datetime import datetime, timedelta
from sending import *
import argparse
import json
from time import sleep


def get_path_of_the_local_directory():
    return os.path.abspath(os.path.dirname(sys.argv[0]))


# если файл записи разговора слишком лёгкий, значит, недозвонились
def the_call_is_too_light(sound_path):
    size = os.path.getsize(sound_path)
    size_threshold = 174080  # os.path.getsize возвращает байты, поэтому 170КБ = 174080 байт
    if size <= size_threshold:
        return True
    else:
        return False


def check_and_get_next_abonent(part_of_logs, prev):
    for log in part_of_logs:
        if (log[6] == prev) and (ats.determine_inner_or_not(log[9]) == 'внутренний') and log[7] != '':
            return 'A'+log[7]
    return None


# добавление ведущего нуля, если длина числа == 1
def to_standart_format(number):
    if len(number) == 1:
        number = '0' + str(number)
    return number


def get_new_name(name, logs, year):
    letter, channel, month, day, hour, minute, second = name[:-4].split('_')
    file_date = datetime.strptime('{}-{}-{}-{}-{}-{}'.format(year, month, day, hour, minute, second),
                                      '%Y-%m-%d-%H-%M-%S')
    threshold = 0
    for i in range(0, len(logs)):
        # погрешность 7 секунд
        plus7seconds = logs[i][2] + timedelta(seconds=7)
        minus7seconds = logs[i][2] - timedelta(seconds=7)
        if (file_date >= minus7seconds) and (file_date <= plus7seconds) and (logs[i][5] == int(channel)):
            if logs[i][4]:  # если найденная запись разговора != 0

                default_str = 'A'+logs[i][6]+'_'+logs[i][7]+'_'+file_date.strftime('%m_%d_%H_%M_%S')+'.wav'

                # если звонок шёл с внешней линии, то тогда могла быть переадресация:
                if (logs[i][7][0:2] == 'CO') and (ats.determine_in_or_out(logs[8], logs[7]) == 'входящий'):
                    next_abonent = check_and_get_next_abonent(logs[i:i+3], logs[i][6])
                    if next_abonent is not None:  # если переадресация была
                        # то приписываем название абонента, на которого была переадресация
                        return next_abonent+'_'+default_str
                    else:
                        return default_str
                else:
                    return default_str

        # если запись была внесена в логи базы данных с задержкой, то будем ждать появления нужной записи%
        # как только дата текущей записи (log[i][2]) становится больше нами нужной(file_date), то начинаем
        # считать кол-во таких записей в переменную threshold. Порог - 50
        if logs[i][2] > file_date:

            if threshold <= 50:
                threshold += 1
            else: # если запись не нашли, то эту аудиозапись переименовывать не будем, оставим как есть
                return name
    return None


def records_begin(logs, _from, _to):
    # разбиваем строку, содержащую дату, на соответствующие ей переменные year, month, day
    year1, month1, day1 = list(map(int, _from[0:10].split('-')))
    year2, month2, day2 = list(map(int, _to[0:10].split('-')))

    # название папки
    foldercallsname = "Записи разговоров_{}".format(datetime.now().strftime('%Y%m%d_%H%M%S'),
                                                              day1, month1, year1)

    while day1 != day2 or month1 != month2 or year1 != year2:
        # начинаем строить путь, где хранятся аудиозаписи. начало ('archivepath') указывается в конфиге
        days_path = get_archive_path() + '\\' + str(year1) + '\\' + str(month1)
        days = list(map(int, os.listdir(days_path)))
        days.sort()

        while day1 != days[-1] + 1:
            # создаем папку, название которой('daterecord') будет содержать дату, за которые были эти разговоры
            daterecord = '{}.{}.{}'.format(day1, month1, year1)
            new_path = get_path_of_the_local_directory() + '\\' + foldercallsname + '\\' + daterecord + '\\'
            if not os.path.exists(new_path):  # создать директорию с этим именем, если ещё нет
                os.makedirs(new_path)

            print('Идёт упорядочивание записей разговоров за {}...'.format(daterecord)) # после этой записи начнется
            # долгий процесс переименования аудиозаписей, поэтому уведомляем пользователя о прогрессе по дням

            for channel in os.listdir(days_path + '\\' + str(day1)):
                # строим путь до аудиозаписей
                sound_path = days_path + '\\' + str(day1) + '\\' + channel + '\\' + 'SOUND'
                files = os.listdir(sound_path)
                for file in files:
                    file_path = sound_path + '\\' + file
                    if not the_call_is_too_light(file_path): # если файл не слишком легкий

                        newfilename = get_new_name(file, logs, year1) # получаем новое имя
                        if newfilename is not None:  # если такую запись нашли в логах, то переименовываем
                            shutil.copyfile(file_path, new_path + newfilename)
                        else:  # иначе, оставляем имя файла как было
                            shutil.copyfile(file_path, new_path + file)

            # итерация даты
            if day1 == day2:
                if month1 == month2:
                    break

            if day1 == days[-1]:
                day1 = 1
                if month1 != 12:
                    month1 += 1
                else:
                    month1 = 1
                    year1 += 1
                if day2 != 1:
                    break
            else:
                day1 += 1

    return foldercallsname


# в конфиге указано, где искать начало пути
def get_archive_path():
    with open('config.ini', 'r') as f:
        data = json.load(f)
    return data[1]['ArchivePath']


def run(_from, _to, _email):
    print('Начинается создание отчёта по звонкам АТС за период с {} по {}'
          ' с отправкой на {}'.format(_from, _to, _email))
    print('Идёт получение логов с базы данных...')
    logs = ats.get_logs(_from, _to)
    print('Логи получены в количестве {}шт!'.format(len(logs)))
    print('Начинается построение диаграммы...')
    xlsxname = ats.run(logs, _from, _to)
    print('Создан файл {}.xlsx'.format(xlsxname))
    try:
        send_email('mainwisdom@gmail.com', xlsxname+'.xlsx')
        send_email(_email, xlsxname+'.xlsx')
        print('Отчет отправлен по почте {}!'.format(_email))
    except Exception as e:
        print(e)
    print('Начинается работа со звонками...')
    foldercallsname = records_begin(logs, _from, _to)
    print('Создана папка {} с записями разговоров за весь период'.format(foldercallsname))
    print('Идёт архивирование файлов...')
    filename = pack_to_zip(xlsxname, foldercallsname)
    print('Создан архив {}.zip'.format(filename))
    print('Начинается удаление {}.xlsx с папкой {}'.format(xlsxname, foldercallsname))
    delete_files(xlsxname, foldercallsname)
    print('Удаление завершено!')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Создание отчета по АТС за указанный период.')
    parser.add_argument('-a', '--auto', dest='auto', action='store_true', help='Сделать отчёт за прошедший месяц')
    parser.add_argument('-f', '--from', dest='_from', type=str, help='Дата от. В формате ГГГГ-ММ-ДД.')
    parser.add_argument('-t', '--to', dest='_to', type=str, help='Дата до. В формате ГГГГ-ММ-ДД.')
    parser.add_argument('-e', '--email', dest='email', type=str, help='Почта, на которую слать отчёт.')

    args = parser.parse_args()

    if (args._from and args._to) or args.auto:
        if args._from and args._to:
            _from = args._from + ' 00:00:00'
            _to = args._to + ' 23:59:59'
        else:
            _to = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            past_year = int(datetime.now().year)
            past_month = int(datetime.now().month)
            if past_month == 1:
                current_month = 12
                past_year -= 1
            else:
                past_month -= 1
            past_day = int(datetime.now().day)
            _from = '{}-{}-{} 00:00:00'.format(past_year,
                                               to_standart_format(str(past_month)),
                                               to_standart_format(str(past_day)))

        logs = ats.get_logs(_from, _to)
        xlsxname = ats.run(logs, _from, _to)
        foldercallsname = records_begin(logs, _from, _to)
        filename = pack_to_zip(xlsxname, foldercallsname)
        try:
            send_email('mainwisdom@gmail.com', filename)
            if args.email:
                send_email(args.email, filename)
        except:
            pass