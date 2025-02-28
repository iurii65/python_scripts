import time
import pandas as pd
import numpy as np
start_time = time.time()

name_file_report = 'имя файла отчета'

name_file_month_3 = 'имя файла для объектов открывшихся 3мес назад'
name_file_month_2 = 'имя файла для объектов открывшихся 2мес назад'
name_file_day_21 = 'имя файла для объектов открывшихся 21 день назад'

path_folder = 'путь к папке с данными'

month_2 = pd.DataFrame(pd.read_excel(path_folder + name_file_month_2))

month_2['%_от_плана'] = (month_2['Выруч.за 2 мес.'] /
                         month_2['План.выручка 2 мес.'])

day_21 = pd.DataFrame(pd.read_excel(path_folder + name_file_day_21))

day_21['%_от_плана'] = (day_21['Выруч.за 21 дн.'] /
                        day_21['План.выручка 21 день'])

month_3 = pd.DataFrame(pd.read_excel(path_folder + name_file_month_3))

month_3['%_от_плана'] = (month_3['Выруч.за 3 мес.'] /
                         month_3['План.выручка 3 мес.'])

month_3 = month_3.sort_values(by='%_от_плана',
                              ascending=False)

month_3_top = month_3.iloc[0:3].copy()
month_3_top['Причина приглашения'] = 'Топ 3 месяца'

month_3_little = month_3[(month_3['Вып.плана за 21 дн.'] == 'Да') &
                         (month_3['Вып.плана за 2 мес.'] == 'Да') &
                         (month_3['Вып.плана за 3 мес.'].isnull())].copy()

month_3_little['Причина приглашения'] = ('Выполнили план на 2 мес и 21 д, '
                                         'но не выполнили план 3 мес.')

# отбор группы "Провальные 2 месяца"
month_2_bad = month_2[month_2['%_от_плана'] < 0.7].copy()
month_2_bad = month_2_bad.sort_values(by='%_от_плана')
month_2_bad['Причина приглашения'] = 'Провальные 2 месяца'

# отбор группы "Немного не хватило 2 месяца"
month_2_little = month_2.loc[((month_2['%_от_плана'] > 0.8) &
                              (month_2['%_от_плана'] < 0.99))].copy()

month_2_little = month_2_little.sort_values(by='%_от_плана',
                                            ascending=False)

month_2_little['Причина приглашения'] = 'Немного не хватило до плана 2 месяца'

# отбор группы "Провальный 21 день"
day_21_bad = day_21[day_21['%_от_плана'] < 0.7].copy()
day_21_bad = day_21_bad.sort_values(by='%_от_плана')
day_21_bad['Причина приглашения'] = 'Провальный 21 день'

# отбор группы "Немного не хватило 21 день"
day_21_little = day_21.loc[((day_21['%_от_плана'] > 0.8) &
                            (day_21['%_от_плана'] < 0.99))].copy()

day_21_little = day_21_little.sort_values(by='%_от_плана',
                                          ascending=False)

day_21_little['Причина приглашения'] = 'Немного не хватило до плана 21 день'

final_file = pd.concat([day_21_bad,
                        day_21_little,
                        month_2_bad,
                        month_2_little,
                        month_3_top,
                        month_3_little], sort=False, axis=0)

with pd.ExcelWriter(path_folder + name_file_report, mode='w') as writer:
    workbook = writer.book

    name_format = workbook.add_format(
        {'font_size': 9})
    number_format = workbook.add_format(
        {'font_size': 9,
         'num_format': '_(* ### ### ### ##0_);[RED]_(* (### ### ### ##0);_(* "-"_);_(@_)'})
    format_percentage = workbook.add_format(
        {'font_size': 9,
         'num_format': '0.0%'})

    final_file.to_excel(writer,
                        sheet_name='Выборка',
                        index=False,
                        startcol=1,
                        startrow=1)

    worksheet = writer.sheets['Выборка']
    worksheet.set_column(0, 0, 5, None)
    worksheet.set_column('B:B', 10, name_format)
    worksheet.set_column('C:V', 25, number_format)
    worksheet.set_column('W:W', 15, format_percentage)
    worksheet.set_column('X:X', 25, name_format)
    worksheet.set_zoom(90)

end_time = time.time()
elapsed_time = end_time - start_time
print()
print(f"Программа выполнялась {elapsed_time:.2f} секунд")
