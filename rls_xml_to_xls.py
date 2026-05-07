#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
XML to XLSX Parser for complete_inventory_v2 format
"""

import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd
import sys
import os


def parse_xml_value(element, default=''):
    """Извлекает текстовое значение элемента, обрабатывая xsi:nil='true'"""
    if element is None:
        return default
    # Проверяем атрибут xsi:nil
    if element.get('{http://www.w3.org/2001/XMLSchema-instance}nil') == 'true':
        return default
    text = element.text
    return text.strip() if text else default


def parse_xml_file(xml_path):
    """Парсит XML файл и возвращает список словарей с данными"""

    # Определяем пространства имён
    namespaces = {
        'xsd': 'http://www.w3.org/2001/XMLSchema',
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
    }

    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Находим все элементы complete_inventory_v2
    items = root.findall('.//complete_inventory_v2')

    data = []

    for item in items:
        row = {}

        # Основные поля
        row['packing_id'] = parse_xml_value(item.find('packing_id'))
        row['desc_id'] = parse_xml_value(item.find('desc_id'))
        row['prep_id'] = parse_xml_value(item.find('prep_id'))
        row['trade_name_id'] = parse_xml_value(item.find('trade_name_id'))
        row['trade_name_rus'] = parse_xml_value(item.find('trade_name_rus'))
        row['trade_name_rus_html'] = parse_xml_value(item.find('trade_name_rus_html'))
        row['lat_name_id'] = parse_xml_value(item.find('lat_name_id'))
        row['lat_name'] = parse_xml_value(item.find('lat_name'))

        # Лекарственная форма
        row['dosage_form_id'] = parse_xml_value(item.find('dosage_form_id'))
        row['dosage_form_full_name'] = parse_xml_value(item.find('dosage_form_full_name'))
        row['dosage_form_short_name'] = parse_xml_value(item.find('dosage_form_short_name'))
        row['dose'] = parse_xml_value(item.find('dose'))
        row['dose_amount'] = parse_xml_value(item.find('dose_amount'))

        # Упаковка
        row['pack1_id'] = parse_xml_value(item.find('pack1_id'))
        row['amount1'] = parse_xml_value(item.find('amount1'))
        row['pack2_id'] = parse_xml_value(item.find('pack2_id'))
        row['amount2'] = parse_xml_value(item.find('amount2'))
        row['pack3_id'] = parse_xml_value(item.find('pack3_id'))
        row['amount3'] = parse_xml_value(item.find('amount3'))

        # Производитель
        row['as_id'] = parse_xml_value(item.find('as_id'))
        row['producer_id'] = parse_xml_value(item.find('producer_id'))
        row['producer_tran'] = parse_xml_value(item.find('producer_tran'))
        row['producer_orig'] = parse_xml_value(item.find('producer_orig'))
        row['producer_country_id'] = parse_xml_value(item.find('producer_country_id'))
        row['producer_country'] = parse_xml_value(item.find('producer_country'))

        # Упаковщик
        row['packer_id'] = parse_xml_value(item.find('packer_id'))
        row['packer_country_id'] = parse_xml_value(item.find('packer_country_id'))
        row['amount'] = parse_xml_value(item.find('amount'))

        # Регистрация
        row['dfc_id'] = parse_xml_value(item.find('dfc_id'))
        row['completeness_id'] = parse_xml_value(item.find('completeness_id'))
        row['reg_id'] = parse_xml_value(item.find('reg_id'))
        row['reg_number'] = parse_xml_value(item.find('reg_number'))

        # Даты регистрации
        reg_date = parse_xml_value(item.find('reg_date'))
        row['reg_date'] = reg_date[:10] if reg_date and len(reg_date) >= 10 else reg_date
        row['rereg_date'] = parse_xml_value(item.find('rereg_date'))

        reg_cancel = parse_xml_value(item.find('reg_cancel_date'))
        row['reg_cancel_date'] = reg_cancel[:10] if reg_cancel and len(reg_cancel) >= 10 else reg_cancel

        row['reg_status_id'] = parse_xml_value(item.find('reg_status_id'))
        row['reg_status'] = parse_xml_value(item.find('reg_status'))

        # Регистратор
        row['registrator_id'] = parse_xml_value(item.find('registrator_id'))
        row['registrator_tran'] = parse_xml_value(item.find('registrator_tran'))
        row['registrator_orig'] = parse_xml_value(item.find('registrator_orig'))
        row['registrator_country_id'] = parse_xml_value(item.find('registrator_country_id'))
        row['registrator_country'] = parse_xml_value(item.find('registrator_country'))

        # НТФР и условия хранения
        row['ntfr_id'] = parse_xml_value(item.find('ntfr_id'))
        row['ntfr_name'] = parse_xml_value(item.find('ntfr_name'))
        row['lt_id'] = parse_xml_value(item.find('lt_id'))
        row['lt_name'] = parse_xml_value(item.find('lt_name'))
        row['lt_month'] = parse_xml_value(item.find('lt_month'))
        row['sc_id'] = parse_xml_value(item.find('sc_id'))
        row['sc_name'] = parse_xml_value(item.find('sc_name'))
        row['sc_short_name'] = parse_xml_value(item.find('sc_short_name'))

        # Прочее
        row['actdate'] = parse_xml_value(item.find('actdate'))
        row['weight'] = parse_xml_value(item.find('weight'))
        row['picname'] = parse_xml_value(item.find('picname'))

        # Готовые форматы
        row['prep_short'] = parse_xml_value(item.find('prep_short'))
        row['prep_full'] = parse_xml_value(item.find('prep_full'))
        row['registration'] = parse_xml_value(item.find('registration'))
        row['firms'] = parse_xml_value(item.find('firms'))

        data.append(row)

    return data


def save_to_excel(data, output_path):
    """Сохраняет список словарей в Excel файл на лист РЛС"""
    df = pd.DataFrame(data)

    # Переупорядочим колонки для удобства (основные поля в начало)
    preferred_order = [
        'packing_id', 'trade_name_rus', 'lat_name', 'dosage_form_full_name',
        'dose', 'producer_tran', 'producer_country', 'reg_number', 'reg_date',
        'reg_status', 'registrator_tran', 'ntfr_name', 'lt_name', 'sc_name',
        'prep_full', 'registration', 'firms'
    ]

    # Добавляем остальные колонки, которых нет в preferred_order
    all_cols = list(df.columns)
    remaining_cols = [col for col in all_cols if col not in preferred_order]
    final_order = preferred_order + remaining_cols

    # Применяем порядок, если все колонки существуют
    existing_cols = [col for col in final_order if col in df.columns]
    df = df[existing_cols]

    # Сохраняем в Excel на лист "РЛС"
    # Проверяем существует ли файл
    import os
    if os.path.exists(output_path):
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='РЛС', index=False)
    else:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='РЛС', index=False)
    
    print(f"✓ Данные успешно сохранены в {output_path} на лист 'РЛС'")
    print(f"✓ Записей экспортировано: {len(df)}")


def main():
    # Пути к файлам (можно изменить или передать через аргументы командной строки)
    if len(sys.argv) >= 2:
        xml_file = sys.argv[1]
    else:
        xml_file = 'inventory.xml'  # имя вашего XML файла

    # Фиксированный выходной файл
    output_file = 'output/esklp/esklp_full.xlsx'

    # Проверка существования файла
    if not os.path.exists(xml_file):
        print(f"❌ Ошибка: файл '{xml_file}' не найден!")
        sys.exit(1)

    try:
        print(f"📄 Парсинг файла: {xml_file}")
        data = parse_xml_file(xml_file)

        if not data:
            print("⚠️  Предупреждение: данные не найдены в XML файле")
            sys.exit(0)

        print(f"✅ Найдено записей: {len(data)}")
        save_to_excel(data, output_file)

    except ET.ParseError as e:
        print(f"❌ Ошибка парсинга XML: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Неожиданная ошибка: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()

