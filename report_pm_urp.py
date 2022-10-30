import openpyxl

# файлы для работы
# файл из статистики с данными
file_stat = 'ПМ-09-2022-УРП.xlsx'
# файл для заполнения
file_pm_urp = 'Таблицы_ПМ_УРП (Разделы 2 и 3).xlsx'

# словарь для хранения раздел куда вставляются данные
# лист откуда берутся данные и с каких ячеек, а потом куда кладутся
dict_data = {
             'Раздел2': {  #    откуда      куда
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

# открываю файлы
# файл с данными
wb_stat = openpyxl.load_workbook(file_stat)
# заполняемый файл
wb_pm_urp = openpyxl.load_workbook(file_pm_urp)

# --- получение данных из файла источника
# список для хранения данных из источника
data_from = []
# цикл прохода по разделам (листам) в заполняемом файле
for wb_pm_urp_sheet, wb_pm_urp_data in dict_data.items():
    # цикл прохода по листам файла источника данных
    for wb_stat_sheet, wb_stat_data in wb_pm_urp_data.items():
        # беру лист в файле с данными
        wb_stat_s = wb_stat[wb_stat_sheet]

        # диапазон ячеек из файла статистики
        cells_range_from = wb_stat_data[0]

        # назначаю диапазон ячеек в файле источнике
        wb_cells_range_from = wb_stat_s[cells_range_from]

        # собираю информацию во временный список
        temp_list = []

        # прохожу по диапазону ячеек в файле источнике
        for cell in wb_cells_range_from[0]:
            temp_list.append(cell.value)

        data_from.append(temp_list)
        temp_list = []

print(data_from)
# print(*data_from, sep='\n')

# --- получение списка ячеек для записи
# цикл прохода по разделам (листам) в заполняемом файле
for wb_pm_urp_sheet, wb_pm_urp_data in dict_data.items():
    # беру лист в файле для записи
    wb_pm_urp_s = wb_pm_urp[wb_pm_urp_sheet]

    for wb_cells_for_data in wb_pm_urp_data.values():
        # диапазон ячеек в файл заполнения
        cells_range_to = wb_cells_for_data[1]
        print(cells_range_to)

        for data in data_from:
            for cell in wb_pm_urp_s[cells_range_to][0]:
                # print(cell, ' = ', data)
                pass




        print()
    print()




# закрываю файл из которого беру данные
wb_stat.close()

# сохраняю файл шаблона и закрываю его
wb_pm_urp.save(file_pm_urp)
wb_pm_urp.close()

# # закрываю программу
# input('Нажмите ENTER')
