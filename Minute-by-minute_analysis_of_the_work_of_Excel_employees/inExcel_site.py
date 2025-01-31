"""
Все права защищены (c) 2024. 
Данный скрипт предназначен для обработки информации о сайтах.
Автор кода не предоставляет прав на использование или распространение данного ПО.
"""

from colorama import init, Fore
import pandas as pd
import os


def process_excel(input_file, url_mapping, output_file):
    """
    :param input_file: Путь к входному файлу Excel, который будет обработан.
    :param url_mapping: Префиксы URL-адресов сопоставления словаря с их заменами.
    :param output_file: Путь к выходному файлу Excel, в котором будут сохранены обработанные данные.
    :return: Путь к сохраненному файлу Excel с обновленными URL-адресами сайтов.
    """
    print(Fore.GREEN + 'Идет соответствие сайтов внутри excel файла')
    global output_file_excelSite
    init()
    output_dir = os.path.dirname(input_file)
    output_file_excelSite = os.path.join(output_dir, 'Отчеты_отделов_и_сотрудников_обновленный.xlsx')
    excel_file = pd.ExcelFile(input_file)
    processed_sheets = {}
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        if 'Сайт' in df.columns:
            def replace_site_value(cell_value):
                if pd.isnull(cell_value):
                    return cell_value
                for url_prefix, replacement in url_mapping.items():
                    if str(cell_value).startswith(url_prefix):
                        return replacement
                return cell_value
            df['Сайт'] = df['Сайт'].apply(replace_site_value)
        processed_sheets[sheet_name] = df
    with pd.ExcelWriter(output_file_excelSite, engine='openpyxl') as writer:
        for sheet_name, df in processed_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Файл с обновленными сайтами сохранен в: {output_file_excelSite}")
    print(Fore.BLUE + 'process_excel и inExcel_site.py выполнены\n')
    return output_file_excelSite
