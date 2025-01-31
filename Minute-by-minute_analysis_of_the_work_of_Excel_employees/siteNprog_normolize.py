"""
Все права защищены (c) 2024. 
Данный скрипт предназначен для нормализации информации.
Автор кода не предоставляет прав на использование или распространение данного ПО.
"""

from colorama import init, Fore
from colorama import Style
from lxml import etree
import os


# Функция для извлечения информации из отчета о программах
def extract_report_data_prog(found_paths):
    """
    :param found_paths: Путь к входному файлу отчета XML.
    :return: Путь к выходному XML-файлу после извлечения и очистки.
    """
    print(Fore.GREEN + 'Идет извлечения информации из отчета о программах')
    global output_file_path_prog
    directory = os.path.dirname(found_paths)
    output_file = os.path.join(directory, 'prog_index.xml')
    parser = etree.XMLParser(remove_blank_text=True, resolve_entities=False)
    tree = etree.parse(found_paths, parser)
    root = tree.getroot()
    for report in root.iter('report'):
        for name in report.iter('name'):
            if name.text == "Программы":
                unwanted_tags = ['user_domain', 'user_name', 'title', 'path', 'text']
                for tag in unwanted_tags:
                    for element in report.findall('.//' + tag):
                        element.getparent().remove(element)
                for user in report.findall('.//user'):
                    fio = user.find('fio')
                    if fio is not None and fio.text is not None and "Телефон" in fio.text:
                        user.getparent().remove(user)
                report_data = etree.tostring(report, pretty_print=True, encoding='unicode')
                with open(output_file, 'w', encoding='utf-8') as file:
                    file.write(report_data)
                output_file_path_prog = output_file.replace('\\', '/')
                print(f'Путь до xml файла с программами: {output_file_path_prog}')
                print(Fore.BLUE + 'extract_report_data_prog выполнено')
                return output_file_path_prog
    return "Тег <report> с указанным <name>Программы</name> не найден."


# Функция для извлечения информации из отчета о сайтах
def extract_report_data_site(found_paths):
    """
    :param found_paths: Путь к входному XML-файлу.
    :return: Путь к выходному XML-файлу, содержащему отфильтрованные данные отчета или сообщение об ошибке, если указанный отчет не найден.
    """
    print(Fore.GREEN + 'Идет извлечения информации из отчета о сайтах')
    global output_file_path_site
    directory = os.path.dirname(found_paths)
    output_file = os.path.join(directory, 'site_index.xml')
    parser = etree.XMLParser(remove_blank_text=True, resolve_entities=False)
    tree = etree.parse(found_paths, parser)
    root = tree.getroot()
    for report in root.iter('report'):
        for name in report.iter('name'):
            if name.text == "Сайты":
                unwanted_tags = ['user_domain', 'user_name', 'title', 'path', 'text']
                for tag in unwanted_tags:
                    for element in report.findall('.//' + tag):
                        element.getparent().remove(element)
                for user in report.findall('.//user'):
                    fio = user.find('fio')
                    if fio is not None and fio.text is not None and "Телефон" in fio.text:
                        user.getparent().remove(user)
                report_data = etree.tostring(report, pretty_print=True, encoding='unicode')
                with open(output_file, 'w', encoding='utf-8') as file:
                    file.write(report_data)
                output_file_path_site = output_file.replace('\\', '/')
                print(f'Путь до xml файла с сайтами: {output_file_path_site}')
                print(Fore.BLUE + 'extract_report_data_site и siteNprog_normolize.py выполнены\n')
                return output_file_path_site
    return "Тег <report> с указанным <name>Сайты</name> не найден."
