from calendar import monthcalendar
from datetime import datetime

from xlrd import open_workbook


def get_data(column):
    """Функция получает данные из Excel-файла по заданному столбцу"""
    # Открываем книгу в текущей директории
    workbook = open_workbook("task_support.xls")
    # Данные на листе 2
    worksheet = workbook.sheet_by_index(1)
    data = []
    # Данные лежат со 2-й по 1002-ую строки
    for row in range(2, 1002):
        data.append(worksheet.cell_value(row, column))
    return data


def num_of_even_numbers():
    """Функция считает кол-во четных чисел в массиве"""
    data = get_data(1)
    # Получим список только четных чисел
    data = [i for i in data if i % 2 == 0]
    print(len(data))


def num_of_simple_numbers():
    """Функция считает кол-во простых чисел в массиве"""
    data = get_data(2)
    count = 0
    for num in data:
        # 0 и 1 не являются простыми
        if num > 1:
            k = 0
            # Достаточно проверить делители от половины числа
            for i in range(2, int(num) // 2 + 1):
                if num % i == 0:
                    k += 1
            if k == 0:
                count += 1
    print(count)


def num_less_value(value):
    """Функция считает кол-во строк в массиве, которые меньше заданного
    значения
    """
    data = get_data(3)
    data = [float(num.replace(',', '.').replace(' ', '')) for num in data]
    count = 0
    for num in data:
        if num < value:
            count += 1
    print(count)


def num_days_of_week_by_abbriviated_name(day):
    """Функция считает количество строк в массиве по указанному
    cокращенному названию дня недели + "пасхалка" ;)
    """
    data = get_data(4)
    count = 0
    for i in data:
        try:
            datetime_object = datetime.strptime(i, '%a %b %d %H:%M:%S %Y')
        except Exception as error:
            print(error)
        else:
            if datetime.timetuple(datetime_object)[6] == 1:
                count += 1
    print(count)


def num_days_of_week_by_serial_num(serial_num):
    """Функция считает кол-во строк в массиве """
    data = get_data(5)
    data = [datetime.strptime(i, '%Y-%m-%d %H:%M:%S.%f') for i in data]
    count = 0
    for i in data:
        if i.weekday() == serial_num:
            count += 1
    print(count)


def num_last_days_of_week_in_month(num_day):
    """Функция считает кол-во последних заданных дней недели в месяце"""
    data = get_data(6)
    data = [datetime.strptime(i, '%m-%d-%Y') for i in data]
    count = 0
    for i in data:
        # Получим максимальную дату в месяце по нужному дню недели
        last_date = max(
            week[num_day] for week in monthcalendar(i.year, i.month)
        )
        if last_date == i.day:
            count += 1
    print(count)


num_of_even_numbers()
num_of_simple_numbers()
num_less_value(0.5)
num_days_of_week_by_abbriviated_name('Tue')
num_days_of_week_by_serial_num(1)
num_last_days_of_week_in_month(1)
