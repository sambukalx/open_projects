"""
Все права защищены (c) 2024. 
Данный скрипт предназначен для поиска файлов.
Автор кода не предоставляет прав на использование или распространение данного ПО.
"""

from colorama import init, Fore
from colorama import Style
import os


def find_xml_files(start_path):
    """
    :param start_path: Путь к каталогу для начала поиска XML-файлов.
    :return: Путь к найденному файлу index.xml с косой чертой или «Нет», если он не найден.
    """
    print(Fore.GREEN + 'Идет поиск xml фалов')
    global found_paths
    for root, dirs, files in os.walk(start_path):
        if 'index.xml' in files:
            full_path = os.path.join(root, 'index.xml')
            found_paths = full_path.replace('\\', '/')
    print(f'Путь до index.xml: {found_paths}')
    print(Fore.BLUE + 'find_xml_files и search_file.py выполнены\n')
    return found_paths
