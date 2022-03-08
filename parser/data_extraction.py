import pandas as pd
import numpy as np
import os.path

# ROOT = '/Users/Dmitriy/PycharmProjects/PdfReport'
ROOT_FULL = os.path.dirname(__file__).rsplit(os.path.basename(os.path.dirname(__file__)))[0]
ROOT = ROOT_FULL[ROOT_FULL.find("\\"):]


var_names = [
    'protocol_num',
    'client_name',
    'substation_name',
    'equip_name',
    'testing_part',
    'equip_type',
    'equip_serial_num',
    'def_type',
    'voltage_class',
    'oil_brand',
    'production_year',
    'setting_year',
    'reason',
    'date',
    'atm_temperature',
    'oil_temperature',
    'comment',
    'syringe_num',
    'power_demand',
    'protocol_num_cut',
    'protocol_date',
    'protocol_key',
    'probe_code',
    'delivery_date',
    'test_conditions',
    'test_date',
    'H2_normative',
    'CH4_normative',
    'C2H4_normative',
    'C2H6_normative',
    'C2H2_normative',
    'CO2_normative',
    'CO_normative',
    'N2_O2_normative',
    'H2_value',
    'CH4_value',
    'C2H4_value',
    'C2H6_value',
    'C2H2_value',
    'CO2_value',
    'CO_value',
    'N2_O2_value',
]


def get_count(number):
    s = str(number)
    if number == 0.0:
        return 0
    elif '.' in s:
        return abs(s.find('.') - len(s)) - 1
    elif 'e' in s:
        reps = s.find('-') + 2
        nums_count = s[reps]
        return int(nums_count)
    elif s == 'nan':
        return -1
    else:
        return 0

def find_file(ROOT):
    FILES_SRC = os.path.join(ROOT, 'parser/AnalyzedFiles')
    EXCEL_FILES = []
    files = os.listdir(FILES_SRC)
    for i in files:
        if i.endswith('.xlsx'):
            EXCEL_FILES.append(os.path.join(FILES_SRC, i))
    return EXCEL_FILES


def extraction_titles(EXCEL_FILES):
    for extraction_file in EXCEL_FILES:
        print(extraction_file)
        input_df = pd.read_excel(extraction_file)
        table = input_df.iloc[:, 0:42]
        # titles_list = table.columns.values.tolist()
        if len(table.columns) == len(var_names):
            table.columns = var_names
            table['date'] = pd.to_datetime(table['date']).dt.strftime('%d.%m.%Y')
            table['delivery_date'] = pd.to_datetime(table['delivery_date']).dt.strftime('%d.%m.%Y')
            table['test_date'] = pd.to_datetime(table['test_date']).dt.strftime('%d.%m.%Y')
            table['atm_temperature'].astype(str)
            get_mismatches(table)
            return table
        else:
            print('the length between files does not match')


def get_value(processed_df, vars_list=[]):
    for idx, row in processed_df.iterrows():
        counter = 0
        vars_dict = {}
        for value in row:
            if row.index[counter] is 'mismatches_formula':
                vars_dict[row.index[counter]] = value.split(',')
                counter += 1
                continue
            elif type(value) is float and get_count(value) > 4:
                vars_dict[row.index[counter]] = f'{value:.5f}'
                counter += 1
                continue
            elif type(value) is float and get_count(value) == 0:
                vars_dict[row.index[counter]] = round(value)
                counter += 1
                continue
            vars_dict[row.index[counter]] = value
            counter += 1
        vars_list.append(vars_dict)
    return vars_list


def find_tickets(table, mismatch_search_list):

    column_names_list = table.columns.to_list()
    counter = 0
    for col_name in column_names_list:
        if mismatch_search_list[3][0] in col_name:
            counter += 1
            if counter // 2:
                value_ticket = col_name[col_name.find('_'):]
                break
            normative_ticket = col_name[col_name.find('_'):]

    return value_ticket, normative_ticket


def get_mismatches(table):

    mismatch_search_list = [
        ('H2', 'водорода'),
        ('CH4', 'метана'),
        ('C2H4', 'этилена'),
        ('C2H6', 'этана'),
        ('C2H2', 'ацетилена'),
        ('CO2', 'диоксида углерода'),
        ('CO', 'оксида углерода'),
        ('N2_O2', 'ОГС'),
    ]

    value_ticket, normative_ticket = find_tickets(table, mismatch_search_list)

    for element in mismatch_search_list:
        normative_element = element[0] + normative_ticket
        processed_table = table[(table[normative_element] != "-")]
        if processed_table.empty:
            table['mismatches_formula'] = '-'
            break
        comparement = table[element[0] + normative_ticket] < table[element[0] + value_ticket]
        if 'mismatches' not in table.columns:
            table['mismatches'] = np.where(comparement, element[1], '')
            table['mismatches_formula'] = np.where(comparement, element[0], '')
            continue
        comparing_index = np.where(table[element[0]+normative_ticket] < table[element[0]+value_ticket])
        prev_value_full = table['mismatches'].values[comparing_index]
        prev_value_formula = table['mismatches_formula'].values[comparing_index]
        # if 0 not in comparing_index[0]:
        #     continue
        if not comparing_index:
            continue
        if not any(prev_value_full):
            table.loc[comparing_index[0], 'mismatches'] = element[1]
            table.loc[comparing_index[0], 'mismatches_formula'] = element[0]
        else:
            for idx, value in enumerate(prev_value_full):
                table.at[comparing_index[0][idx], 'mismatches'] = '{}, {}'.format(value, element[1])
                table.at[comparing_index[0][idx], 'mismatches_formula'] = '{},{}'.format(prev_value_formula[idx],
                                                                                         element[0])


def operate():
    EXCEL_FILES = find_file(ROOT)
    table = extraction_titles(EXCEL_FILES)
    return get_value(table)

