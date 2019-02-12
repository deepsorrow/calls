import firebirdsql
import xlsxwriter
from datetime import datetime, timedelta
import json


class Formats: # текстовые форматы для записей на сводной таблице
    def __init__(self, workbook):
        wb = workbook
        self.border_center = wb.add_format({'align': 'center', 'valign': 'center', 'border': 1})
        self.bold_center = wb.add_format({'bold': True, 'align': 'center', 'valign': 'center'})
        self.center = wb.add_format({'align': 'center', 'valign': 'center'})


def get_logs(_from='2018-10-10', _to='2018-10-11'):
    conn = firebirdsql.connect(
        host='******************',
        port='***************',
        database='***************',
        user='***************',
        password='***************'
    )
    cur = conn.cursor()
    cur.execute("select * from calllog where CALL_BEGIN > '{}' and CALL_BEGIN < '{}'".format(_from, _to))

    logs = cur.fetchall()
    conn.close()

    return logs


def create_workbook(xlsxname):
    workbook = xlsxwriter.Workbook('{}.xlsx'.format(xlsxname))
    return workbook


def determine_numbers_category(number, line):
    if not number:
        return 0

    if (len(number) == 3 or number == 'PE0001') and number[0] != 'C':
        return 1  # внутренний
    elif line != 0:
        if len(number) == 6 or number[0] == 'C':
            return 2  # город
        elif len(number) == 11:
            if (number[1] == '9' or number [1] == '8') and number[1:4] != '800':
                return 4  # сотовый
            elif number[1:5] == '3843':
                return 2  # город
            else:
                return 3  # межгород
    else:
        return 0


# для сводной таблицы считаем количество проговоренных минут для каждого абонента + считаем отдельно громкоговоритель
def count_each_number_and_get_pe0001(logs, summary_data):
    pe0001 = 0
    for entry in logs:
        if entry[4]:  # длит. звонка != 0
            category = determine_numbers_category(entry[7], entry[5])

            if category:
                index = '{}.{}'.format(entry[6], category)
                try:
                    summary_data[index] += entry[4]
                except:
                    summary_data[index] = entry[4]

            if entry[7] == 'PE0001':
                pe0001 += entry[4]
    return pe0001


def determine_in_or_out(number, to):
    if not to:
        return ''
    elif number:
        return 'входящий'
    else:
        return 'исходящий'


def determine_inner_or_not(number):
    if number:
        return 'внутренний'
    else:
        return 'внешний'


def determine_was_raised(number):
    datetime1 = datetime.strptime(str(timedelta(seconds=number)), '%H:%M:%S')
    return datetime1


# создание логов в своем формате
def export_to_xlsx_raw_data(logs, workbook):
    ws_logs = workbook.add_worksheet('Логи')

############################################################
    ws_logs.write(0, 0, "Дата", Formats(workbook).center)  # 2
    ws_logs.write(0, 1, "Время", Formats(workbook).center)  # 2
    ws_logs.write(0, 2, "Абонент", Formats(workbook).center)  # 6
    ws_logs.write(0, 3, "Номер", Formats(workbook).center)  # 7
    ws_logs.write(0, 4, "Направление", Formats(workbook).center)  # 8 and 7
    ws_logs.write(0, 5, "Звонок", Formats(workbook).center)  # 3
    ws_logs.write(0, 6, "Линия", Formats(workbook).center)  # 5
    ws_logs.write(0, 7, "Тип звонка", Formats(workbook).center)  # 9
    ws_logs.write(0, 8, "Длительность разговора", Formats(workbook).center)  # 4

    ws_logs.set_column(0, 1, 10)  # дата время
    ws_logs.set_column(2, 2, 8)   # абонент
    ws_logs.set_column(3, 3, 13)  # номер
    ws_logs.set_column(4, 4, 15)  # направление
    ws_logs.set_column(5, 5, 10)  # звонок
    ws_logs.set_column(6, 6, 7)   # линия
    ws_logs.set_column(7, 7, 15)  # тип звонка
    ws_logs.set_column(8, 8, 24)  # длительность разг.
############################################################

    row = 1
    col = 0
    cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    date_format = workbook.add_format({'num_format': 'hh:mm:ss'})
    for log in logs:
        if not (determine_numbers_category(log[7], log[5]) != 1 and log[5] == 0):
            ws_logs.write(row, col, log[2].strftime("%d.%m.%Y"), cell_format)  # дата
            col += 1
            ws_logs.write(row, col, log[2].strftime("%H:%M:%S"), cell_format)  # время
            col += 1
            ws_logs.write(row, col, log[6], cell_format) # абонент
            col += 1
            ws_logs.write(row, col, log[7], cell_format) # номер
            col += 1
            ws_logs.write(row, col, determine_in_or_out(log[8], log[7]), cell_format) # направление
            col += 1
            ws_logs.write(row, col, str(timedelta(seconds=log[3])), cell_format) # звонок
            col += 1
            ws_logs.write(row, col, log[5], cell_format) # линия
            col += 1
            ws_logs.write(row, col, determine_inner_or_not(log[9]), cell_format) # тип звонка
            col += 1
            if log[4]:
                ws_logs.write_datetime(row, col, determine_was_raised(log[4]), date_format) # длительность разговора
            else:
                ws_logs.write(row, col, 'Неотвеченный', cell_format)  # длительность разговора
            col = 0
            row += 1


# считаем проговоренных минут по вертикали
def summary_calls_vertically(workbook, worksheet, catalog):
    row, col = 3, 5
    sum = 0

    prev = catalog[0]
    for entry in catalog:
        if prev[0][0:3] == entry[0][0:3]:
            sum += entry[1]
        else:
            worksheet.write(row, col, round(sum/60), Formats(workbook).center)
            row += 1
            sum = entry[1]
        prev = entry
    worksheet.write(row, col, round(sum/60), Formats(workbook).center)


# считаем проговоренных минут по горизонтали
def summary_calls_horizontally(workbook, worksheet, catalog, size):
    row = size
    sum = 0

    for col in range(1,5):
        for entry in catalog:
            if int(entry[0][4:]) == col:
                sum += entry[1]
        worksheet.write(row, col, round(sum/60), Formats(workbook).center)
        sum = 0


def substitute_number_with_name(number_str, data):
    try:
        return data[number_str]
    except:
        return number_str


# создание сводной таблицы
def create_summary_table(workbook, catalog, _from='2018-10-10', _to='2018-10-11'):
    ws_table = workbook.add_worksheet('Таблица')

###############################################################
    ws_table.write(0, 0, "АТС за период {} {}".format(_from, _to), Formats(workbook).bold_center)  # 2

    # объединение ячеек
    ws_table.merge_range('A1:F1', 'Звонки по АТС за период с {} по {}'.format(_from, _to),
                         Formats(workbook).bold_center)
    ws_table.merge_range('A2:A3', 'Абонент', Formats(workbook).border_center)
    ws_table.merge_range('B2:E2', 'Вид звонка', Formats(workbook).border_center)
    ws_table.merge_range('F2:F3', 'Общий итог', Formats(workbook).border_center)

    ws_table.write(2, 1, "Внутренний")
    ws_table.write(2, 2, "Городской")
    ws_table.write(2, 3, "Межгородний")
    ws_table.write(2, 4, "Сотовый")

    ws_table.set_column(0, 0, 34)  # абонент
    ws_table.set_column(1, 5, 13)  # виды звонка + итог
#################################################################
    row, col = 2, 0

    prev = '000.0'
    size = 3
    with open('config.ini', 'r') as fp:
        data = json.load(fp)[0]

    for entry in catalog:
        if prev[0][0:3] != entry[0][0:3]:
            size += 1
            row += 1
            if entry[0] != 'PE0001':
                ws_table.write(row, col, substitute_number_with_name(entry[0][0:3], data))  # абонент
            else:
                ws_table.write(row, col, substitute_number_with_name(entry[0], data))  # абонент

        ws_table.write(row, col + int(entry[0][4:]), round(entry[1]/60), Formats(workbook).center)
        prev = entry

    ws_table.write(size, 0, 'Общий итог', Formats(workbook).center)
    summary_calls_vertically(workbook, ws_table, catalog)
    summary_calls_horizontally(workbook, ws_table, catalog, size)

    return size


# создание диаграммы
def create_stacked_chart(workbook, size, _from='2018-10-10', _to='2018-10-11'):

    chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    chart.set_title({'name': 'Звонки по АТС за период с {} по {}'.format(_from[0:10], _to[0:10])})
    chart.set_y_axis({'name': 'Продолжительность в минутах'})
    chart.add_series({
        'name': 'Внутренний',
        'categories': '=Таблица!$A$4:$A${}'.format(size),
        'values': '=Таблица!$B$4:$B${}'.format(size)
    })
    chart.add_series({
        'name': 'Городской',
        'categories': '=Таблица!$A$4:$A${}'.format(size),
        'values': '=Таблица!$C$4:$C${}'.format(size)
    })
    chart.add_series({
        'name': 'Межгородний',
        'categories': '=Таблица!$A$4:$A${}'.format(size),
        'values': '=Таблица!$D$4:$D${}'.format(size)
    })
    chart.add_series({
        'name': 'Сотовый',
        'categories': '=Таблица!$A$4:$A${}'.format(size),
        'values': '=Таблица!$E$4:$E${}'.format(size)
    })
    chart.set_style(10)

    chartsheet = workbook.add_chartsheet('Диаграмма')
    chartsheet.set_chart(chart)
    chartsheet.activate()


# сортируем абонентов по номеру их телефона
def custom_sort(input_entry):
    return input_entry[0:3]


def run(logs, _from, _to):
    xlsxname = "Отчет_АТС_%s" % (datetime.now().strftime('%Y%m%d_%H%M%S'))
    workbook = create_workbook(xlsxname)

    # создание первой страницы с логами
    export_to_xlsx_raw_data(logs, workbook)

    # создание второй страницы со сводной таблицей
    catalog = {}
    pe0001 = count_each_number_and_get_pe0001(logs, catalog)
    sorted_catalog = sorted(catalog.items(), key=custom_sort) # сортируем абонентов по номеру их телефона
    sorted_catalog.append(['PE0001', pe0001])
    size = create_summary_table(workbook, sorted_catalog, _from, _to)

    # создание диаграммы
    create_stacked_chart(workbook, size, _from, _to)

    workbook.close()

    return xlsxname
