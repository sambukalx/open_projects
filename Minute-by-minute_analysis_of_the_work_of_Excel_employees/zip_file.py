"""
Все права защищены (c) 2024. 
Данный скрипт предназначен для работы с ZIP файлами.
Автор кода не предоставляет прав на использование или распространение данного ПО.
"""

from colorama import init, Fore
from colorama import Style
import os
import zipfile


def unzip_file(zip_path):
    """
    :param zip_path: Путь к zip-файлу, который необходимо извлечь.
    :return: Путь к папке, в которую было извлечено содержимое zip-файла.
    """
    print(Fore.GREEN + 'Идет распаковка архива')
    dir_name = os.path.dirname(zip_path)
    folder_name = os.path.splitext(os.path.basename(zip_path))[0]
    output_folder = os.path.join(dir_name, folder_name)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)
    start_path = output_folder.replace('\\', '/')
    print(f"Путь до распакованной папки: {start_path}")
    print(Fore.BLUE + 'unzip_file и zip_file.py выполнены\n')
    return start_path
