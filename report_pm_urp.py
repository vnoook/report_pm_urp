import openpyxl

# файлы для работы
file_stat = 'ПМ-09-2022-УРП.xlsx'
file_pm_urp = 'Таблицы_ПМ_УРП (Разделы 2 и 3).xlsx'

# словарь для хранения источника и
dict_data = {
             'Раздел2': {
                         '7': ['B60:S60', 'B9:S9'],
                         '8': ['B60:S60', 'B10:S10'],
                         '9': ['B60:S60', 'B11:S11'],
                         '10': ['B60:S60', 'B12:S12'],
                         '11': ['B60:S60', 'B13:S13'],
                         '12': ['B60:S60', 'B14:S14'],
                         '13': ['B60:S60', 'B15:S15'],
                         '14': ['B60:S60', 'B16:S16'],
                         '15': ['B60:S60', 'B17:S17'],
                         '16': ['B60:S60', 'B18:S18'],
                         '17': ['B60:S60', 'B19:S19'],
                         '18': ['B61:S61', 'B20:S20'],
                         '19': ['B61:S61', 'B21:S21'],
                         '20': ['B61:S61', 'B22:S22'],
                         '21': ['B61:S61', 'B23:S23']
                        },
             'Раздел3': {
                         '22': ['B60:AE60', 'B8:AE8'],
                         '23': ['B60:AE60', 'B9:AE9'],
                         '24': ['B60:AE60', 'B10:AE10'],
                         '25': ['B60:AE60', 'B11:AE11'],
                         '26': ['B60:AE60', 'B12:AE12'],
                         '27': ['B60:AE60', 'B13:AE13']
                        }
            }

wb_stat = openpyxl.load_workbook(file_stat)
wb_pm_urp = openpyxl.load_workbook(file_pm_urp)

for wb_pm_urp_sheet, wb_pm_urp_data in dict_data.items():
    # print(wb_pm_urp_sheet, wb_pm_urp_data, sep='\n')
    wb_pm_urp_s = wb_pm_urp[wb_pm_urp_sheet]

    for wb_stat_sheet, wb_stat_data in wb_pm_urp_data.items():
        print(wb_stat_sheet, wb_stat_data)
