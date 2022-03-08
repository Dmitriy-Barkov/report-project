import os.path
from data_parser import data_extraction as de
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import os

ROOT = de.ROOT
ASSETS_DIR = os.path.join('/', ROOT[1:], 'assets')

TEMPLATE_SRC = os.path.join(ROOT, 'templates')
CSS_SRC = os.path.join(ROOT, 'static/css')
DEST_DIR = os.path.join(ROOT, 'output')

TEMPlATE = 'charg.html'
CSS = 'style.css'
# OUTPUT_FILENAME = 'report.pdf'


def start(template_vars, OUTPUT_FILENAME):
    print('Starting generate report...')
    env = Environment(loader=FileSystemLoader(TEMPLATE_SRC))

    template = env.get_template(TEMPlATE)
    css = os.path.join(CSS_SRC, CSS)

    # variables
    # template_vars = {"protocol_num": "ХА-11-09-03-21",
    #                  "client_name": 'ООО "Заказчик"',
    #                  "substation_name": 'ПС "Рендерная"',
    #                  'assets_dir': 'file://' + ASSETS_DIR,
    #                  }
    adding_vars = [
        ('assets_dir', 'file://' + ASSETS_DIR),
        ('Organization_name', 'Сибэнергодиагностика'),
        ('Address', 'Россия, г. Новосибирск, Микрорайон Зеленый Бор 3'),
        ('identification_number', '5405367743'),
        ('correcting_number', '540501001'),
        ('URL', 'https://sibenedia.ru'),
        ('email', 'info@sibenedia.ru'),
        ('telephone', '+7(383)269-21-10'),
        ('equipment_name', 'Хроматографический комплекс'),
        ('type_name', 'Хроматэк - Кристалл 5000.1'),
        ('serial_num', '253152'),
        ('accuracy', 'Предел допускаемого значения относительного СКО выходного сигнала не более 2%. Предел допускаемого значения изменения выходного сигнала ±5%.'),
        ('cert_number', 'С-НН/20-12-2021/118741170'),
        ('next_cert_date', '20.12.2022'),
        ('normative_doc', 'СТО 34.01-23.1-001-2017'), # Либо РД 153-34.0-46.302-00, либо СТО 34.01-23.1-001-2017
        ('accreditation_number', 'РОСС RU. 0001.22HP40'),
        ('emission_date', '10 июня 2016 года'),
    ]

    for element in adding_vars:
        template_vars[element[0]] = element[1]

    # rendering report
    rendered_string = template.render(template_vars)
    html = HTML(string=rendered_string)
    report = os.path.normpath(os.path.join(DEST_DIR, OUTPUT_FILENAME))
    # report = os.path.join(DEST_DIR, OUTPUT_FILENAME)
    html.write_pdf(report, stylesheets=[css])
    print('file generated successfully and under {}'.format(DEST_DIR))


if __name__ == '__main__':
    operating_vars_list = de.operate()
    counter = 0
    for list in operating_vars_list:
        counter += 1
        OUTPUT_FILENAME = 'Протокол {}, {}, {}.pdf'.format(list['protocol_num'], list['substation_name'], list['equip_name'])
        start(template_vars=list, OUTPUT_FILENAME=OUTPUT_FILENAME)


