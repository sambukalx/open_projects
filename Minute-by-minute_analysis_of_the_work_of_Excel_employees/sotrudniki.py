"""
Все права защищены (c) 2024. 
Данный скрипт предназначен для работы с TXT файлом.
Автор кода не предоставляет прав на использование или распространение данного ПО.
"""

from colorama import init, Fore
from colorama import Style


def parse_departments(people):
    """
    Эта функция читает файл, содержащий отделы и их сотрудников.
    И возвращает словарь, где ключами являются названия отделов и значения
    представляют собой списки сотрудников этих отделов.

    :param people: Путь к текстовому файлу, содержащему отделы и сотрудников.
    :return: словарь с названиями отделов в качестве ключей и списками имен сотрудников в качестве значений.
    """
    print(Fore.GREEN + 'Идет определение сотрудников и отделов')
    departments = {}
    current_department = None
    with open(people, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    for line in lines:
        line = line.strip()
        if line.endswith(':'):
            current_department = line[:-1]
            departments[current_department] = []
        elif line:
            departments[current_department].append(line)
    print(Fore.BLUE + "parse_departments и sotrudniki.py выполнены\n")
    return departments
