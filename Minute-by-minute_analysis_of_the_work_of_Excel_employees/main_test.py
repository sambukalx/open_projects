"""
Все права защищены (c) 2024. 
Данный скрипт предназначен для тестового запуска проекта.
Автор кода не предоставляет прав на использование или распространение данного ПО.
"""

import json
from colorama import init, Fore
from colorama import Style
from excel_create import *
from myallpath import *
from search_file import *
from siteNprog_normolize import *
from siteNprog_toexcel import *
from sotrudniki import *
from zip_file import *
from zvonki_normolize import *
from zvonki_toexcel import *
from clearPath import *
from inExcel_site import *
from stahName import *
from infoWork_stah import *
from myalldata import *
from infoStah_toexcel import *
from format import *
from dost_file import *
from bitrix_normolize import *

init(autoreset=True)


def main():
    """
   Основная функция, организующая процесс выбора файлов, анализа данных, создания и изменения файлов Excel,
    и выполнение различных файловых операций, таких как распаковка, удаление, преобразование и сканирование каталогов.

    :return: None
    """
    try:

        # 0. Выбор файлов
        print(Fore.RED + "#0\nЗапуск select_files в allmypath.py")
        people, zip_path, zvonki, output_excel_path = select_files()  # myallpath.py

        # 1. Парсинг данных сотрудников и отделов
        print(Fore.RED + "#1\nЗапуск parse_departments в sotrudniki.py")
        mas_sotrudniki = parse_departments(people)  # sotrudniki.py

        # 2. Создание Excel-файла с отделами и сотрудниками
        print(Fore.RED + "#2\nЗапуск create_department_employee_excel в excel_create.py")
        output_xlsx_path = create_department_employee_excel(mas_sotrudniki, output_excel_path)  # excel_create.py

        # 3. Распаковка zip-файла
        print(Fore.RED + "#3\nЗапуск unzip_file в zip_file.py")
        start_path = unzip_file(zip_path)  # zip_file.py

        # 4. Удаление файлов PNG и других
        print(Fore.RED + "#4\nЗапуск delete_png_files в clearPath.py")
        delete_png_files(start_path)  # clearPath.py

        # 5. Удаление папок с малым количеством файлов
        print(Fore.RED + "#5\nЗапуск delete_small_folders в clearPath.py")
        delete_small_folders(start_path)  # clearPath.py

        # 6. Конвертация файлов XLS в XLSX
        print(Fore.RED + "#6\nЗапуск convert_xls_to_xlsx в clearPath.py")
        convert_xls_to_xlsx(start_path)  # clearPath.py

        # 7. Удаление файлов по определенным условиям
        print(Fore.RED + "#7\nЗапуск delete_x_files в clearPath.py")
        delete_x_files(start_path)  # clearPath.py

        # 8. Удаление папок на основе C2
        print(Fore.RED + "#8\nЗапуск delete_folders_based_on_C2_recursive в clearPath.py")
        delete_folders_based_on_C2_recursive(start_path)  # clearPath.py

        # 9. Удаление PACS файлов
        print(Fore.RED + "#9\nЗапуск delete_pacs в clearPath.py")
        delete_pacs(start_path)  # clearPath.py

        # 10. Поиск XML-файлов
        print(Fore.RED + "#10\nЗапуск find_xml_files в search_file.py")
        found_paths = find_xml_files(start_path)  # search_file.py

        # 11. Извлечение данных о программах из отчета
        print(Fore.RED + "#11\nЗапуск extract_report_data_prog в siteNprog_normolize.py")
        output_file_path_prog = extract_report_data_prog(found_paths)  # siteNprog_normolize.py

        # 12. Извлечение данных о сайтах из отчета
        print(Fore.RED + "#12\nЗапуск extract_report_data_site в siteNprog_normolize.py")
        output_file_path_site = extract_report_data_site(found_paths)  # siteNprog_normolize.py

        # 13. Загрузить данные о программах из XML файла
        print(Fore.RED + "#13\nЗапуск load_program_data в siteNprog_toexcel.py")
        program_data = load_program_data(output_file_path_prog)  # siteNprog_toexcel.py

        # 14. Загрузить данные о сайтах из XML файла
        print(Fore.RED + "#14\nЗапуск load_site_data в siteNprog_toexcel.py")
        site_data = load_site_data(output_file_path_site)  # siteNprog_toexcel.py

        # 15. Обновление Excel-файла с данными сотрудников
        print(Fore.RED + "#15\nЗапуск update_employee_sheets в siteNprog_toexcel.py")
        update_employee_sheets(output_xlsx_path, program_data, site_data)  # siteNprog_toexcel.py

        # 16. Обработка и сохранение данных Excel
        print(Fore.RED + "#16\nЗапуск process_excel в # inExcel_site.py")
        output_file_excelSite = process_excel(output_xlsx_path, url_mapping)  # inExcel_site.py

        # 17. Создание копии фала со звонками
        print(Fore.RED + "#17\nЗапуск create_file_copy в zvonki_normolize.py")
        copy_file_path_zv = create_file_copy(zvonki)  # zvonki_normolize.py

        # 18. Обработка и сохранение данных о звонках
        print(Fore.RED + "#18\nЗапуск process_and_save_calls_data в zvonki_normolize.py")
        process_and_save_calls_data(copy_file_path_zv, replacements)  # zvonki_normolize.py

        # 19. Обновление отчета по звонкам
        print(Fore.RED + "#19\nЗапуск zvonkiExcel в zvonki_toexcel.py")
        zvonkiExcel(output_file_excelSite, copy_file_path_zv)  # zvonki_toexcel.py

        # 20. Чистка фалов Стахановца и переименование
        print(Fore.RED + "#20\nЗапуск process_folders в stahName.py")
        process_folders(start_path, replacements)  # stahName.py

        # 21. Удаление ненужных файлов/папок
        print(Fore.RED + "#21\nЗапуск rename_folders_from_excel_cell в stahName.py")
        rename_folders_from_excel_cell(start_path)  # stahName.py

        # 22. Объединения всех сотрудников в один список
        print(Fore.RED + "#22\nЗапуск get_all_employees в infoWork_stah.py")
        all_employees = get_all_employees(mas_sotrudniki)  # infoWork_stah.py

        # 23. Сканирования всех подкаталогов
        print(Fore.RED + "#23\nЗапуск scan_folders в infoWork_stah.py")
        scan_folders(start_path, all_employees)  # infoWork_stah.py

        # 24. Вставка значений из стахановца в excel
        print(Fore.RED + "#24\nЗапуск update_excel_with_employee_data в infoStah_toexcel.py")
        update_excel_with_employee_data(info_work_stah, output_file_excelSite)  # infoStah_toexcel.py

        # 25. Преобразование файла битрикс xls в xlsx
        print(Fore.RED + "#25\nЗапуск convert_html_to_xlsx в bitrix_normolize.py")
        file_path_bit = convert_html_to_xlsx(bitrix_path)  # bitrix_normolize.py

        # 26. Форматирование файла Битрикса
        print(Fore.RED + "#26\nЗапуск replace_values_in_xlsx в bitrix_normolize.py")
        replace_values_in_xlsx(file_path_bit, bit_replacements)  # bitrix_normolize.py

        # 27. Форматирование файла отчета xlsx
        print(Fore.RED + "#27\nЗапуск format_excel_file в format.py")
        format_excel_file(output_file_excelSite)  # format.py

        # 28. Переименовывание основных файлов Стахановца
        print(Fore.RED + "#28\nЗапуск frename_xlsx_files в dost_file.py")
        frename_xlsx_files(start_path)  # dost_file.py

        # 29. Перенос временных файлов
        print(Fore.RED + "#29\nЗапуск fileto_log в dost_file.py")
        timefiles = [copy_file_path_zv, start_path, output_xlsx_path]
        fileto_log(timefiles, output_excel_path)  # dost_file.py

        print(
            Fore.RED + "Все " + Fore.YELLOW + "функции " + Fore.GREEN + "успешно " + Fore.BLUE + "выполнены " + Fore.BLACK + "!")

    except Exception as e:
        print(f"Произошла ошибка: {e}")


if __name__ == '__main__':
    main()
