import pandas as pd
import os as os
import numpy as np

# печатать в xlsx отчет
create_a_report_to_firm = False
total_report = True

# год анализа
year = 'анализируемый год'

# корневой путь
root_path = ('корневой путь, где храняться папки')

# наименования папок
name_folder_report = 'имя папки для сохранения отчетов'
name_folder_old = 'имя папки файлами из старой базы'
name_folder_new = 'имя папки файлами из новой базы'

# файл чек
name_check = 'файл для проверки отчетов'

check_file = pd.read_excel(root_path + name_check,
                           skiprows=1)

# получаем список имен CSV файлов
list_name_old = os.listdir(root_path + name_folder_old)
# получаем список имен CSV файлов
list_name_new = os.listdir(root_path + name_folder_new)

# сводный df
check_df = pd.DataFrame(columns=[])

for old, new in zip(list_name_old, list_name_new):
    print(f'Из старой базы выбран файл – {old}')
    print(f'Из новой базы выбран файл – {new}')

    # загрузка файла из старой базы 1С
    df_old = pd.read_csv(root_path + name_folder_old + '/' + old,
                         encoding='ansi',
                         sep=';',
                         decimal=',',
                         usecols=[1, 2, 3, 4, 5, 6, 7, 8],
                         names=[],
                         header=0,
                         dtype={4: str})

    name_firm_old = df_old[''].iloc[0]

    # загрузка файла из новой базы 1С
    df_new = pd.read_csv(root_path + name_folder_new + '/' + new,
                         encoding='ansi',
                         sep=';',
                         decimal=',',
                         usecols=[1, 2, 6, 14, 15, 16, 17],
                         names=[],
                         header=0,
                         dtype={6: str})

    name_firm_new = df_new[''].iloc[0]

    # проверка соответствия сверяемых компаний
    if name_firm_new == name_firm_old:
        print(f'Фирма в старом и новом файле сходится {name_firm_old}')

        # обработка файла из старой базы
        part1_old_df = (
            df_old)[
            df_old['']
            .str
            .contains('',
                      na=False)]

        part2_old_df = (
            df_old)[
            df_old['']
            .str
            .contains('',
                      na=False)]

        total_old_df = pd.concat([part1_old_df, part2_old_df],
                                 axis=0)

        sum_count_old = total_old_df[''].sum() * -1

        total_old_df = total_old_df.pivot_table(
            index='',
            values=['', ''],
            aggfunc='sum').reset_index()

        total_old_df[['', '']] = (
                total_old_df[['', '']] * -1)

        # обработка файла из новой базы
        sum_count_new = df_new[''].sum()

        total_new_df = df_new.pivot_table(
            index='',
            values=['', ''],
            aggfunc='sum').reset_index()

        # создание DF с уникальным кодом товора
        df_name_product_old = df_old[['',
                                      '']]

        df_name_product_new = df_new[['',
                                      '']]

        report_tax = pd.concat([df_name_product_old, df_name_product_new],
                               sort=False,
                               axis=0)

        report_tax = report_tax.drop_duplicates(subset='')

        # добавление информации о сумме/кол-ве товаров из 2-ух баз 1С
        report_tax = pd.merge(report_tax, total_new_df,
                              how='left',
                              on='')

        report_tax = pd.merge(report_tax, total_old_df,
                              how='left',
                              on='',
                              suffixes=('_новая', '_старая'))

        report_tax[''] = df_new[''].iloc[0]
        report_tax[''] = name_firm_new

        report_tax = report_tax.reindex(columns=['',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 '',
                                                 ''])

        # создание DF с информацией об изменениях в новой базе
        new_return = (
            df_new)[
            df_new['']
            .str
            .contains('',
                      na=False)]

        name_columns_new_return = {'': '',
                                   '': ''}

        new_return = new_return.rename(columns=name_columns_new_return)

        new_return = new_return.pivot_table(
            index='',
            values=['', ''],
            aggfunc='sum').reset_index()

        # добавление информации об изменениях в основной отчет
        report_tax = pd.merge(report_tax, new_return,
                              how='left',
                              on='')

        report_tax[['',
                    '',
                    '',
                    '',
                    '',
                    '']] \
            = (report_tax[['',
                           '',
                           '',
                           '',
                           '',
                           '']].fillna(0))

        report_tax[''] = (
                report_tax[''] -
                report_tax[''] -
                report_tax[''])

        report_tax[''] = (
                report_tax[''] -
                report_tax[''] -
                report_tax[''])

        # создание переменных обозначающих тоталы значений
        total_count_new_yt = report_tax[''].sum()
        total_summa_new_yt = report_tax[''].sum()
        total_count_old_yt = report_tax[''].sum()
        total_summa_old_yt = report_tax[''].sum()
        total_count_new_return = report_tax[''].sum()
        total_summa_new_return = report_tax[''].sum()
        total_count_diff_yt = report_tax[''].sum()
        total_summa_diff_yt = report_tax[''].sum()

        new_name_columns = {
            '': '',
            '': '',
            '': '',
            '': '',
            '': '',
            '': '',
            '': '',
            '': ''}

        report_tax = report_tax.rename(columns=new_name_columns)

        # добавление информации из проверочного файла
        df_check_inn = check_file[
            check_file[''] == report_tax[''].iloc[0]]

        # значение показателей из проверочного файла
        check_new_cost = (
            df_check_inn[''].iloc[
                0].sum())

        check_old_cost = (
            df_check_inn[''].iloc[
                0].sum())


        # проверка, что в исходных и обработанных файлах тотал сходится
        if ((np.rint(total_count_old_yt) == np.rint(sum_count_old))
                and (np.rint(total_count_new_yt) == np.rint(sum_count_new))):
            status_report = 'Отчет создан'
            print(f'Отчет по компании {name_firm_old} создан')

        else:
            status_report = 'Отчет не создан'
            print(f'Отчет по компании {name_firm_new} не создан по '
                  f'причине расхождения первоначальных значений')

            continue

        # заполение сводного DF
        check_list = [name_firm_new,
                      int(report_tax[''].iloc[0]),
                      status_report,
                      float(check_new_cost),
                      float(check_old_cost),
                      float(total_count_new_yt),
                      float(total_summa_new_yt),
                      float(total_count_old_yt),
                      float(total_summa_old_yt),
                      float(total_count_new_return),
                      float(total_summa_new_return),
                      float(total_count_diff_yt),
                      float(total_summa_diff_yt)]

        check_df.loc[len(check_df)] = check_list

        # разницы между контрольным файлом и созданным отчетом
        check_df[''] \
            = (check_df['']
               - check_df[''])

        check_df[''] \
            = (check_df['']
               - check_df[''])

        #создание xlsx фала с отчетом по компании
        if create_a_report_to_firm:
            with pd.ExcelWriter(
                    root_path +
                    name_folder_report +
                    year +
                    name_firm_new +
                    '.xlsx',
                    mode='w') as writer:
                workbook = writer.book

                header_format = workbook.add_format(
                    {'bold': True,
                    'font_size': 9})
                number_format = workbook.add_format(
                    {'font_size': 9,
                    'num_format': '### ### ##0'})
                firm_format = workbook.add_format(
                    {'font_size': 9})

                report_tax.to_excel(writer,
                                    sheet_name='свод',
                                    index=False,
                                    startcol=1,
                                    startrow=5)

                worksheet = writer.sheets['свод']

                worksheet.write('F5', total_count_new_yt)
                worksheet.write('G5', total_summa_new_yt)
                worksheet.write('H5', total_count_old_yt)
                worksheet.write('I5', total_summa_old_yt)
                worksheet.write('J5', total_count_new_return)
                worksheet.write('K5', total_summa_new_return)
                worksheet.write('L5', total_count_diff_yt)
                worksheet.write('M5', total_summa_diff_yt)

                worksheet.set_column('B:D', 12, firm_format)
                worksheet.set_column('A:A', 5, None)
                worksheet.set_column('F:AG', 30, number_format)
                worksheet.set_column('E:E', 55, firm_format)
                worksheet.set_zoom(80)

                print(f'Создание отчета в xlsx по {name_firm_new} окончено')

    else:
        print('Компании в файлах не сходятся!!!')
        continue

if total_report:
    with pd.ExcelWriter(
            root_path + name_folder_report + year + '.xlsx',
            mode='w') as writer:
        workbook = writer.book
        header_format = workbook.add_format(
            {'bold': True,
             'font_size': 9})
        number_format = workbook.add_format(
            {'font_size': 9,
             'num_format': '### ### ### ##0'})
        firm_format = workbook.add_format(
            {'font_size': 9}
        )
        check_df.to_excel(writer,
                          sheet_name='свод',
                          index=False,
                          startcol=1,
                          startrow=2)

        worksheet = writer.sheets['свод']

        worksheet.set_column('B:D', 15, firm_format)
        worksheet.set_column('A:A', 5, None)
        worksheet.set_column('E:N', 25, number_format)
        worksheet.set_column('O:P', 40, number_format)
        worksheet.set_zoom(80)
