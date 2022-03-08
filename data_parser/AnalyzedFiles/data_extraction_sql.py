import pandas as pd
import mysql.connector as sql_connector
import os.path

NAME = sql_connector.connect(host='192.168.1.128', user='user', password='CiscoU238',  port='3306', auth_plugin='mysql_native_password')
print(NAME.is_connected())


# ROOT = '/Users/Dmitriy/PycharmProjects/PdfReport'
#
# var_names = [
#     'protocol_num',
#     'client_name',
#     'substation_name',
#     'equip_name',
#     'testing_part',
#     'equip_type',
#     'equip_serial_num',
#     'def_type',
#     'voltage_class',
#     'oil_brand',
#     'production_year',
#     'setting_year',
#     'reason',
#     'date',
#     'atm_temperature',
#     'oil_temperature',
#     'comment',
#     'syringe_num',
#     'power_demand',
#     'protocol_num_cut',
#     'protocol_date',
#     'protocol_key',
#     'probe_code',
#     'delivery_date',
#     'test_conditions',
#     'test_date',
#     'H2_normative',
#     'CH4_normative',
#     'C2H4_normative',
#     'C2H6_normative',
#     'C2H2_normative',
#     'CO2_normative',
#     'CO_normative',
#     'N2_O2_normative',
#     'H2_value',
#     'CH4_value',
#     'C2H4_value',
#     'C2H6_value',
#     'C2H2_value',
#     'CO2_value',
#     'CO_value',
#     'N2_O2_value',
# ]
#
#
# def find_file(ROOT):
#     FILES_SRC = os.path.join(ROOT, 'parser/AnalyzedFiles')
#     EXCEL_FILES = []
#     files = os.listdir(FILES_SRC)
#     for i in files:
#         if i.endswith('.xlsx'):
#             EXCEL_FILES.append(os.path.join(FILES_SRC, i))
#     return EXCEL_FILES
#
#
# def extraction_titles(EXCEL_FILES):
#     for extraction_file in EXCEL_FILES:
#         print(extraction_file)
#         input_df = pd.read_excel(extraction_file)
#         table = input_df.iloc[:, 0:42]
#         # titles_list = table.columns.values.tolist()
#         if len(table.columns) == len(var_names):
#             table.columns = var_names
#             table['date'] = pd.to_datetime(table['date']).dt.strftime('%d.%m.%Y')
#             table['delivery_date'] = pd.to_datetime(table['delivery_date']).dt.strftime('%d.%m.%Y')
#             table['test_date'] = pd.to_datetime(table['test_date']).dt.strftime('%d.%m.%Y')
#             table['atm_temperature'].astype(str)
#             return table
#         else:
#             print('the length between files does not match')
#         # return
#
#
# def get_value(processed_df, vars_list=[]):
#     for idx, row in processed_df.iterrows():
#         counter = 0
#         vars_dict = {}
#         for value in row:
#             vars_dict[row.index[counter]] = value
#             counter += 1
#         vars_list.append(vars_dict)
#     return vars_list
#
#
# def operate():
#     EXCEL_FILES = find_file(ROOT)
#     table = extraction_titles(EXCEL_FILES)
#     return get_value(table)

