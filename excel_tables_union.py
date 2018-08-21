import pandas as pd
from collections import Counter
from os import listdir
from os.path import splitext, join
import openpyxl
import warnings
import sys
import os
import json



if not sys.warnoptions:
    warnings.simplefilter("ignore")


mapping = {
    'Наименование' : ['Номенклатура', 
                      'Наим', 
                      'Назв', 
                      'Товар',
                      'Наименование',
                      'Модель',
                      'НАИМЕНОВАНИЕ',
                      'Марка'],
    'Описание' : ['Объем', 
                  'Масса', 
                  'Кол-во',
                  'Количество',
                  'Мин. Отгрузка', 
                  'Ø', 
                  'Тип', 
                  'ИНФ', 
                  'Инф',
                  'инф'
                  'парам', 
                  'Ед', 
                  'назнач', 
                  'упак',
                  'вес',
                  'вип',
                  'примеч',
                  'Длин',
                  'Матер',
                  'Описание'],
    'Цена' : ['Цена', 
              'ВИП', 
              'МРЦ',
              'VIP',
              'Валюта',
              'Дилерская',
              'цена',
              'Сумма',
              'сумма',
              'Стоимость',
              'ЦЕНА',
              'Цены',
              'Цены оптовые',
              'Цена без НДС']
          }

def read_xls(path):
    return openpyxl.reader.excel.load_workbook(path, data_only=True)


def drop_none_rows(dframe):
    drop_rows = []
    for i in range(dframe.index.size):
        if (dframe.iloc[i,:].isnull().sum()/len(dframe.columns.values)) >= 0.85:
            drop_rows.append(i)
    return dframe.drop(dframe.index[drop_rows], axis=0)


def drop_none_columns(dframe):
    drop_columns = []
    for i in dframe.columns:
        if ((dframe.loc[:,i].isnull().sum()/dframe.index.size) >= 0.95):
            drop_columns.append(i)
    return dframe.drop(drop_columns, axis=1)
    

def replace_zeros_with_nones(dframe):
    return dframe.replace([0], [None])


def replace_nones_with_str(dframe):
    return dframe.replace([None], [''])


def create_drop_row_list(dframe, mapping):
    drop_rows = []
    br = False
    for row in range(dframe.index.size):
        for col_num, col_val in enumerate(dframe.columns.values):
            for value in mapping['Наименование']:
                try:
                    if value in dframe.iloc[row, col_num]:
                        drop_rows.append(row)
                        br = True
                        break
                except TypeError:
                    continue
                if br:
                    break
            if br:
                break
        if br:
            br = False
            continue
    return drop_rows


def split_dframe(drop_rows, dframe):
    if len(drop_rows) > 1:
        dframes = []
        drop_rows.append(dframe.index.size)
        for i, v in enumerate(drop_rows[:-1]):
            dframes.append(dframe.iloc[drop_rows[i]:drop_rows[i+1],:])
        return dframes
    elif len(drop_rows) == 1:
        return [replace_nones_in_names_row(dframe.iloc[drop_rows[0]:])]
    else:
        return [create_first_row_as_name(dframe)]
        

def fill_missing_values(dframe):
    for row in range(dframe.index.size):
        for col_num, col_val in enumerate(dframe.columns.values):
            if dframe.iloc[row, col_num] == '':
                if (row == 0) & (col_num != 0):
                    dframe.iloc[row, col_num] = dframe.iloc[row, col_num - 1]
                elif row > 0:
                    dframe.iloc[row, col_num] = dframe.iloc[row - 1, col_num]
    return dframe


def add_names_to_values(dframe, mapping):
    re_frame = dframe
    for col_num, col_val in enumerate(re_frame.columns.values):
        for value in mapping['Описание']:
            if value in str(re_frame.iloc[0, col_num]):
                for row in range(1, re_frame.index.size):
                    re_frame.iloc[row, col_num] = '{}: {}'.format(re_frame.iloc[0, col_num], re_frame.iloc[row, col_num])
                break
    return re_frame


def rename_columns(dframe, mapping):
    new_column_names = {}
    br = False
    for col_num, col_val in enumerate(dframe.columns.values):
        new_column_names[col_num] = 'Другое'
        for key, value in mapping.items():
            for ii in value:
                if ii in str(dframe.iloc[0, col_num]):
                    new_column_names[col_num] = key
                    br = True
                    break
            if br:
                br = False
                break     
    dframe.columns = new_column_names.values()
    return dframe


def drop_first_row(dframe):
    new_frame = dframe.drop(dframe.index[0])
    return new_frame


def merge_columns(dframe):
    count_names = Counter(dframe.columns.values)
    res_dframe = pd.DataFrame()
    for i, v in count_names.items():
        if i == 'Другое':
            continue
        if v > 1:
            res_dframe[i] = dframe[i].apply(lambda x: ' | '.join(x.astype(str)),axis=1)
        else:
            res_dframe[i] = dframe[i]
    return res_dframe


def create_first_row_as_name(dframe):
    new_column_names = {}
    for col_num, col_val in enumerate(dframe.columns.values):
        try:
            int(dframe.iloc[0, col_num])
            new_column_names[col_num] = 'Описание'
        except (ValueError, TypeError):
            new_column_names[col_num] = 'Наименование'
    try:
        dframe.loc[-1] = [i for i in new_column_names.values()]
        dframe.sort_index(inplace=True)
        return dframe
    except ValueError:
        return pd.DataFrame()


def replace_nones_in_names_row(dframe):
    frm = dframe.reset_index()
    for col_num, col_val in enumerate(frm.columns):
        if not frm.iloc[0, col_num]:
            for key, value in mapping.items():
                for pattern in value:
                    x = frm[frm[col_val].astype('str').str.contains(pattern)].index
                    if x.size > 0:
                        frm.iloc[0, col_num] = frm.loc[x[0], col_val] 
    return frm


def get_files_list(path):
    excel_files_list = []
    for i in listdir(path):
        if splitext(i)[1].lower() in ['.xlsx', '.xls', '.xlsm']:
            if not '~$' in splitext(i)[0]:
                excel_files_list.append(i)
    return excel_files_list


def union_df_list(dframes_list):
    return pd.concat(dframes_list).reset_index().drop(['index'], axis=1)


def add_file_name_column(dframe, file_name):
    dframe['Файл'] = [file_name for i in range(dframe.index.size)]
    return dframe


def parse_dframe_from_excel(files_path, file_name):
    return pd.read_excel(join(files_path, file_name))


if __name__ == "__main__":
    result_frames = []

    file_path = os.curdir
    files_list = get_files_list(file_path)
    print('\n')
    for num, file_name in list(enumerate(files_list, start=1)):

        dframes = []
        dframes.append(parse_dframe_from_excel(file_path, file_name))
        print('#{} Название файла: {}'.format(num, file_name))
        print('Done! - 1 stage')

        test_frames = []
        for i in dframes:
            test_frame = split_dframe(create_drop_row_list(i, mapping), i)
            test_frames += test_frame

        test_frames_2 = []
        for i in test_frames:
            x = drop_none_columns(replace_zeros_with_nones(i))
            if x.size:
                x = replace_nones_with_str(drop_none_rows(x))
            else:
                print(i, )
                continue
            if x.size:
                test_frames_2.append(fill_missing_values(x))
            else:
                print(i)
                continue
        print('Done! - 2 stage')

        test_frames_3 = []
        for i in test_frames_2:
            test_frames_3.append(merge_columns(drop_first_row(rename_columns(add_names_to_values(i, mapping), mapping))))
        print('Done! - 3 stage')
        print('\n')

        for i in test_frames_3:
            i = add_file_name_column(i, file_name)

        result_frames += test_frames_3


    result_frame = union_df_list(result_frames)
    result_frame.to_csv(path_or_buf='result_prices.csv', sep=';', encoding='utf-8-sig')

    print('Done! - 4| Saved\n')
    input('ВЫПОЛНЕНО УСПЕШНО!!! Нажмите Enter для выхода из программы...')