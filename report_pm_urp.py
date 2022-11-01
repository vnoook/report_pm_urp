# Программа для создания отчёта Стат-Урп.
#
# В папке с программой должен быть файл отчёта из АРМ Статистика,
# а также файл для сбора данных из таблиц ПМ_УРП.
#
#     Максим Красовский \ октябрь 2022 \ noook@yandex.ru
# ...
# INSTALL
# pip install openpyxl
# COMPILE
# pyinstaller -F report_pm_urp.py
# ...

import os.path
import openpyxl
import datetime


# функция составления названия файла для сохранения
def name_of_file():
    # текущие месяц и год
    number_of_month = datetime.datetime.today().month
    number_of_year = datetime.datetime.today().year

    # если запустили в январе, то (месяц и год) надо сменить на (декабрь и (год-1))
    # иначе (месяц-1)
    if number_of_month == 1:
        number_of_month = 12
        number_of_year -= 1
    else:
        number_of_month -= 1

    # если номер месяца цифра, то добавить 0 в начало
    # иначе просто перевести в строку
    if number_of_month < 10:
        name_month = '0'+str(number_of_month)
    else:
        name_month = str(number_of_month)

    file_name = 'ПМ-' + name_month + '-' + str(number_of_year) + '-УРП.xlsx'
    return file_name


def file_exist(fle_name):
    if os.path.exists(fle_name):
        return True
    else:
        print(f'Ожидается файл "{fle_name}"')
        return False


# --- файлы для работы
# файл из статистики с данными
# file_stat = 'ПМ-09-2022-УРП.xlsx'
file_stat = name_of_file()
# файл для заполнения
file_pm_urp = 'Таблицы_ПМ_УРП (Разделы 2 и 3).xlsx'

if not file_exist(file_stat) or not file_exist(file_pm_urp):
    input('Нажмите ENTER')
    exit()

# словарь для хранения разделов куда вставляются данные,
# листы откуда берутся данные и с каких ячеек, а потом куда кладутся
dict_data = {
             'Раздел2': {     # откуда ..... куда
                         '7': ('B60:S60',  'B9:S9'),
                         '8': ('B60:S60',  'B10:S10'),
                         '9': ('B60:S60',  'B11:S11'),
                         '10': ('B60:S60', 'B12:S12'),
                         '11': ('B60:S60', 'B13:S13'),
                         '12': ('B60:S60', 'B14:S14'),
                         '13': ('B60:S60', 'B15:S15'),
                         '14': ('B60:S60', 'B16:S16'),
                         '15': ('B60:S60', 'B17:S17'),
                         '16': ('B60:S60', 'B18:S18'),
                         '17': ('B60:S60', 'B19:S19'),
                         '18': ('B61:S61', 'B20:S20'),
                         '19': ('B61:S61', 'B21:S21'),
                         '20': ('B61:S61', 'B22:S22'),
                         '21': ('B61:S61', 'B23:S23')
                        },
             'Раздел3': {
                         '22': ('B60:AE60', 'B8:AE8'),
                         '23': ('B60:AE60', 'B9:AE9'),
                         '24': ('B60:AE60', 'B10:AE10'),
                         '25': ('B60:AE60', 'B11:AE11'),
                         '26': ('B60:AE60', 'B12:AE12'),
                         '27': ('B60:AE60', 'B13:AE13')
                        }
            }

# --- открываю файлы
# файл с данными
wb_stat = openpyxl.load_workbook(file_stat)
# заполняемый файл
wb_pm_urp = openpyxl.load_workbook(file_pm_urp)

# --- алгоритм считывания и записи поячеечно
# цикл прохода по разделам (листам) в заполняемом файле
for wb_pm_urp_sheet, wb_pm_urp_data in dict_data.items():
    # цикл прохода по листам файла источника данных
    for wb_stat_sheet, wb_stat_data in wb_pm_urp_data.items():
        # беру лист в файле с данными
        wb_stat_s = wb_stat[wb_stat_sheet]
        # беру лист в файле для заполнения
        wb_pm_urp_s = wb_pm_urp[wb_pm_urp_sheet]

        # диапазоны ячеек в файлах
        cells_range_from = wb_stat_data[0]
        cells_range_to = wb_stat_data[1]

        # получаю диапазон ячеек в файле источнике
        wb_cells_range_from = wb_stat_s[cells_range_from][0]
        # получаю диапазон ячеек в файле для записи
        wb_cells_range_to = wb_pm_urp_s[cells_range_to][0]

        # --- запись данных в файл для записи
        # если диапазоны сходятся по длине, то можно записывать, иначе сообщение
        if len(wb_cells_range_from) == len(wb_cells_range_to):
            # прохожусь по длине вложенных списков
            for ind in range(len(wb_cells_range_from)):
                wb_cells_range_to[ind].value = wb_cells_range_from[ind].value
        else:
            print(f'В строке {wb_stat_sheet}: {wb_stat_data} не совпадают длины диапазонов!\n'
                  f'Антон, исправь диапазон в этой строке.\n')

# закрываю файл из которого беру данные
wb_stat.close()

# сохраняю файл для записи и закрываю его
wb_pm_urp.save(file_pm_urp)
wb_pm_urp.close()

# закрываю программу
print()
print('Перенос данных между файлами закончен удачно.')
print()
input('Нажмите ENTER')
