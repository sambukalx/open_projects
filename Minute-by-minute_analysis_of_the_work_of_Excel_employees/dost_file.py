"""
Все права защищены (c) 2024. 
Данный скрипт предназначен для работы с файлами.
Автор кода не предоставляет прав на использование или распространение данного ПО.
"""

import os
import shutil
import logging
logger = logging.getLogger(__name__)


def fileto_log(files, output_excel_path):
    """
    :param files: Список путей к файлам, которые необходимо переместить в каталог журналов.
    :type files: список
    :param output_excel_path: Путь к каталогу, в котором будет создан каталог журнала.
    :type output_excel_path: str
    :return: Путь к созданному каталогу журнала, в который были перемещены файлы.
    :rtype: str
    """
    log_path = os.path.join(output_excel_path, 'log')
    if not os.path.exists(log_path):
        os.makedirs(log_path)
        logger.info(f"Создана папка {log_path}")
    else:
        logger.info(f"Папка {log_path} уже существует")
    for file in files:
        try:
            if os.path.exists(file):
                base_name = os.path.basename(file)
                destination = os.path.join(log_path, base_name)
                if os.path.exists(destination):
                    name, ext = os.path.splitext(base_name)
                    counter = 1
                    while os.path.exists(destination):
                        new_name = f"{name}_{counter}{ext}"
                        destination = os.path.join(log_path, new_name)
                        counter += 1
                shutil.move(file, destination)
                logger.info(f"Файл {file} перемещен в {destination}")
        except Exception as e:
            logger.error(f"Ошибка при перемещении файла {file}: {str(e)}")
    return log_path


def frename_xlsx_files(base_path):
    """
    :param base_path: Корневой каталог для начала поиска файлов .xlsx
    :return: None
    """
    for root, dirs, files in os.walk(base_path):
        current_folder = os.path.basename(root)
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
                new_file_name = f"{os.path.splitext(file)[0]}_{current_folder}.xlsx"
                new_file_path = os.path.join(root, new_file_name)
                os.rename(file_path, new_file_path)
                print(f"Файл {file_path} переименован в {new_file_path}")
