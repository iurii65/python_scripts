# срипт для преобразование выгрузки 1С в плоскую таблицу
import pandas as pd
import numpy as np
pd.set_option('display.max_columns', None)

name_file_report = 'имя готового отчета'

name_file_1c = 'имя файла из 1С'

path_folder_file = 'путь до файла 1С (по этому же пути готовый результат)'

# загружаем Excel и преобразовали в df Pandas, дали название столбцов
market_bonus = pd.DataFrame(pd.read_excel(path_folder_file + name_file_1c,
                                          names=['name', 'amount']))

num_frame = (market_bonus[market_bonus['amount']
             .notna()]
             .index
             .tolist())[:2]

# создали список, где 0 значение номер строки с "оборот по дт"
# 1 значение номер строки с тотальным значением
num_one, num_two = num_frame

# количество уровней
number_of_levels = num_two - num_one

# удалили лишние строки
new_market_bonus = market_bonus.iloc[num_two:]

# названия уровней
name_column = market_bonus['name'].tolist()[num_one:num_two]

# плоская таблица
new_df = pd.DataFrame(columns=name_column)
new_df['amount'] = None

# список значений
target_row = [0] * (number_of_levels + 1)

# список значений по уровням
level_value = [0] * number_of_levels

# список кумулятивных значений по уровням
level_value_cum = [0] * number_of_levels

# текущий уровень в итерации
level = 0

for i in range(num_two, len(new_market_bonus) + num_two):

    # записал текущее значение в массив текущих значений по уровням
    level_value[level] = new_market_bonus.loc[i, 'amount']

    # добавил текущее значение в массив кумулятивных значений по уровням
    level_value_cum[level] += level_value[level]
    level_value_cum[level] = round(level_value_cum[level], 2)

    # заполняем строку с информацией
    target_row[level] = new_market_bonus.loc[i, 'name']


    # проверяем спустились ли на нижний уровень
    if level == number_of_levels - 1:
        # добавляем значение
        target_row[number_of_levels] = level_value[level]

        # вносим строку в df
        new_df = new_df._append(pd.Series(target_row, index=new_df.columns),
                                ignore_index=True)

        # проверка на заполненность группы
        for j in range(number_of_levels, 0, -1):
            if level_value_cum[level] == level_value[level - 1]:
                level_value_cum[level] = 0
                level -= 1

    else: #спускаемся на уровень ниже
        level += 1

total_plot = new_df['amount'].sum()
total_1C = market_bonus.at[num_two, 'amount']

if round(total_plot, 2) == round(total_1C, 2):
    print('Успешно')
    new_df.to_excel(path_folder_file + name_file_report,
                    sheet_name='float_table',
                    index=False)

else:
    print('ВНИМАНИЕ, ошибка тотал не сходится ')
