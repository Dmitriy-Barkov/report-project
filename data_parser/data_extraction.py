from parser import ParserError

import pandas as pd
import numpy as np
import os.path
import re

# ROOT = '/Users/Dmitriy/PycharmProjects/PdfReport'
ROOT_FULL = os.path.dirname(__file__).rsplit(os.path.basename(os.path.dirname(__file__)))[0]
ROOT = ROOT_FULL[ROOT_FULL.find("\\"):]


""" Связи для переименновывания колонок """
var_names = {
    'Протокол №': 'protocol_num',
    'Наименование предприятия': 'client_name',
    'Место установки': 'substation_name',
    'Диспетчерское наименование': 'equip_name',
    'Место отбора': 'testing_part',
    'Тип оборудования': 'equip_type',
    'Заводской номер': 'equip_serial_num',
    'Тип защиты': 'def_type',
    'Класс напряжения': 'voltage_class',
    'Марка масла': 'oil_brand',
    'Год выпуска': 'production_year',
    'Год ввода': 'setting_year',
    'Причина отбора': 'reason',
    'Дата отбора': 'date',
    'Температура окр': 'atm_temperature',
    'Температура масла': 'oil_temperature',
    'Комментарий': 'comment',
    'Нагрузка': 'power_demand',
    'Номер шприца': 'syringe_num',
    'Номер протокола': 'protocol_num_cut',
    'Дата протокола': 'protocol_date',
    'Шифр протокола': 'protocol_key',
    'Шифр пробы': 'probe_code',
    'Дата доставки пробы': 'delivery_date',
    'Условия проведения': 'test_conditions',
    'Дата выполнения испытания': 'test_date',
    'H2': 'H2',
    'Н2': 'H2',
    'CH4': 'CH4',
    'C2H4': 'C2H4',
    'C2H6': 'C2H6',
    'C2H2': 'C2H2',
    'CO2': 'CO2',
    'CO': 'CO',
    'CО': 'CO',
    'N2_O2': 'N2_O2',
    'ОГС': 'N2_O2',
    'СРГ': 'SRG'
}

mismatch_search_list = {
        'H2': 'водорода',
        'CH4': 'метана',
        'C2H4': 'этилена',
        'C2H6': 'этана',
        'C2H2': 'ацетилена',
        'CO2': 'диоксида углерода',
        'CO': 'оксида углерода',
        'N2_O2': 'ОГС',
        'SRG': 'СРГ',
        'C2O3': 'Паль',
}


""" Форматирование таблицы """

def change_columns_names(connections_dict, df, change_dict = {}):
    # Функция для создания словаря смены колонок

    for i in connections_dict.keys():
        pattern_chemical = df.filter(regex=f'^({i}) *\W').columns.values
        pattern_words = df.filter(regex=f'^{i}').columns.values
        if len(pattern_chemical) == 1:
            change_dict[pattern_chemical[0]] = f'{connections_dict[i]}'
        elif len(pattern_chemical) > 1:
            change_dict[pattern_chemical[0]] = f'{connections_dict[i]}_normative'
            change_dict[pattern_chemical[1]] = f'{connections_dict[i]}_value'
        elif len(pattern_words) > 0:
            change_dict[pattern_words[0]] = f'{connections_dict[i]}'
    return change_dict


def get_comparision_list(df, ready_list=[]):
    # Функция отбора колонок для сравнения
    for pattern in ['value', 'normative']:
        ready_list.append([i for i in df.columns.unique() if pattern in i])
    return ready_list


def set_notation(dataframe, comparision_columns):
    joined_list = [*comparision_columns[1], *comparision_columns[0]]
    dataframe[joined_list] = dataframe[joined_list].apply(lambda x: ['%.5f' % i if isinstance(i, float) else i for i in x ])
    return dataframe

def change_data_type(df, column_name):
    try:
        part = pd.to_datetime(df[column_name]).dt.strftime('%Y')
    except:
        print(f"В колонке {column_name} указаны некорректные данные")
    else:
        return part

def reformat_table(df):

    # Устанавливаем список колонок, подлежащих удалению
    prohibited_columns = ['Соответствие', 'По чему?']

    # Удаляем лишние колонки
    df = df.drop(columns=prohibited_columns).dropna(subset=['Наименование предприятия'])

    # Переименновываем столбцы
    df.rename(columns=change_columns_names(connections_dict=var_names, df=df), inplace=True)

    # Приводим столбцы к формату даты
    df['date'] = pd.to_datetime(df['date']).dt.strftime('%d.%m.%Y')
    df['delivery_date'] = pd.to_datetime(df['delivery_date']).dt.strftime('%d.%m.%Y')
    # df['production_year'] = pd.to_datetime(df['production_year']).dt.strftime('%Y')
    df['production_year'] = change_data_type(df, column_name='production_year')
    df['setting_year'] = change_data_type(df, column_name='setting_year')

    # df['setting_year'] = pd.to_datetime(df['setting_year']).dt.strftime('%Y')
    df['test_date'] = pd.to_datetime(df['test_date']).dt.strftime('%d.%m.%Y')

    # Приводим колнку к строковому типу
    df['atm_temperature'] = df['atm_temperature'].astype("string")

    # Приводим колнку к целочисленному типу
    df.voltage_class = df.voltage_class.astype(int)
    df.equip_serial_num = df.equip_serial_num.astype(int)

    # Заполняем колонку тип защиты в случае, если в ней None
    df.fillna({'def_type': 'без специальных защит'}, inplace=True)

    # Заполняем оставшиеся None прочерками
    df.fillna('-', inplace=True)

    # Создаем столбец с результатами сравнения параметров м/у собой.
    # В случае несовпадения - записываем названия выбивающихся параметров в колонку mismatches_formula
    comparision_columns = get_comparision_list(df)

    # Получаем 2 таблицы: со значениями эксперимента и с нормами
    values_df = df.filter(items=comparision_columns[0])
    normative_df = df.filter(items=comparision_columns[1])

    # Создаем таблицу для проведения сравнения
    comparison_table = pd.DataFrame(np.where(values_df.values <= normative_df.values, 'True', 'False'),
                                    columns=[i.split('_')[0] for i in values_df.columns])

    # Заносим полученные различия в столбец mismatches_formula в основной df
    df['mismatches_formula'] = comparison_table.apply(
        lambda row: row[row == 'False'].index.to_list() if row[row == 'False'].shape[0] > 0 else '',
        axis=1
    )

    # Получаем колонку - эквивалент со значениями выбившихся показателей
    df['mismatches'] = df.mismatches_formula.agg(lambda x: ', '.join([mismatch_search_list[i] for i in x]))

    # Изменим научную нотацию
    df = set_notation(dataframe=df, comparision_columns=comparision_columns)

    return df


def get_value(df):
    vars_list = [i for i in df.apply(lambda row: row.to_dict(), axis=1)]
    return vars_list


def find_file(ROOT):
    FILES_SRC = os.path.join(ROOT, 'parser_engine/AnalyzedFiles')
    for file in os.listdir(FILES_SRC):
        if file.endswith('.xlsx'):
            # EXCEL_FILE = os.path.abspath(file)
            path = os.path.normpath(os.path.join(FILES_SRC, file))
    return path


def operate():
    EXCEL_FILE = find_file(ROOT)
    input_df = pd.read_excel(EXCEL_FILE)
    table = reformat_table(df=input_df)
    return get_value(table)
