"""
Все права защищены (c) 2024. 
Данный скрипт предназначен для создания Excel файла.
Автор кода не предоставляет прав на использование или распространение данного ПО.
"""

from colorama import init, Fore
from colorama import Style
import pandas as pd
import os


def create_department_employee_excel(mas_sotrudniki, output_excel_path):
    """
    Функция для создания Excel файла с листами отделов и сотрудников.

    :param mas_sotrudniki: словарь с данными о сотрудниках и отделах (формат: {отдел: [список сотрудников]})
    :param output_excel_path: путь к папке, где будет создан Excel файл.
    :return: полный путь к созданному Excel файлу или None в случае ошибки.
    """
    print(Fore.GREEN + 'Идет создание excel фала с именами сотрудников')
    if not isinstance(mas_sotrudniki, dict):
        print("Ошибка: входные данные должны быть словарем.")
        return
    if not os.path.exists(output_excel_path):
        print(f"Ошибка: указанная папка не существует: {output_excel_path}")
        return
    output_excel_path = os.path.join(output_excel_path, 'Отчеты_отделов_и_сотрудников.xlsx')
    try:
        if os.path.exists(output_excel_path):
            os.remove(output_excel_path)
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as excel_writer:
            for dept, employees in mas_sotrudniki.items():
                if employees:  # Проверяем, что в отделе есть сотрудники
                    df_dept = pd.DataFrame({'Сотрудники': employees})
                    df_dept.to_excel(excel_writer, sheet_name=dept[:31], index=False)
                    for employee in employees:
                        df_employee = pd.DataFrame()
                        df_employee.to_excel(excel_writer, sheet_name=employee[:31], index=False)
                else:
                    print(f"Внимание: отдел '{dept}' не содержит сотрудников и будет пропущен.")
    except Exception as e:
        print(f"Ошибка при создании Excel файла: {e}")
        return
    print(f'Excel файл успешно создан и сохранен по пути: {output_excel_path}')
    print(Fore.BLUE + 'create_department_employee_excel и excel_crete.py выполнены\n')
    return output_excel_path
