"""
Все права защищены (c) 2024.
Данный скрипт обрабатывает Excel-файлы по заданным параметрам:
- Форматирует номер, 'ГО' и 'ИНН'.
- Сверяет данные с другим файлом для исключения дублей.
- Удаляет ненужные столбцы и строки по условию.
Код не подлежит использованию или копированию без разрешения автора.
"""

import os
import pandas as pd
from openpyxl import load_workbook
from tkinter import Tk, filedialog, Button, Label, Entry, StringVar

def format_number(x):
    """
    Преобразует значение x в целое число (с ведущими нулями до 6 знаков).
    """
    try:
        return f'{int(x):06d}'
    except (ValueError, TypeError):
        return x

def remove_duplicates(df, column_name):
    """
    Удаляет дубли по указанному столбцу (column_name), сохраняя первую строку.
    """
    df.drop_duplicates(subset=column_name, keep='first', inplace=True)

def filter_by_go(df, threshold):
    """
    Фильтрует строки, где значение 'ГО' больше либо равно threshold.
    Возвращает отфильтрованный DataFrame.
    """
    df = df[df['ГО'].astype(float) >= threshold]
    return df

def remove_inn_duplicates(df, comparison_file_path):
    """
    Сверяет 'ИНН' в df с данными в comparison_file_path (лист 'ИНН')
    и удаляет все совпадения из df.
    """
    comparison_df = pd.read_excel(comparison_file_path, sheet_name='ИНН')
    if 'ИНН' in df.columns and 'ИНН' in comparison_df.columns:
        inn_values = comparison_df['ИНН'].astype(str).tolist()
        df = df[~df['ИНН'].astype(str).isin(inn_values)]
    return df

def fill_go_month_year(df, value):
    """
    Заполняет столбец 'ГО_Месяц_Год' значением value с 1-й строки до последней ненулевой 'ГО'.
    """
    last_index = df['ГО'].last_valid_index()
    df.loc[1:last_index, 'ГО_Месяц_Год'] = value

def rename_columns(df):
    """
    Переименовывает столбцы в удобный формат.
    """
    rename_mapping = {
        'ГО': 'Сумма ГО',
        'Наименование компании': 'Название компании',
        'Область поставщика': 'Регион',
        'TZ': 'Часовой пояс поставщика',
        'Телефон 1': 'Рабочий телефон',
        'Телефон 2': 'Мобильный телефон',
        'Телефон 3': 'Домашний телефон',
        'Email 1': 'Рабочий e-mail',
        'Email 2': 'Частный e-mail',
        'Директор': 'Имя, Фамилия',
        'Название': 'Предмет тендера',
        'Номер': 'Номер тендера'
    }
    df.rename(columns=rename_mapping, inplace=True)

def process_files():
    """
    Основная функция:
    1) Читает основной файл.
    2) Форматирует 'Номер', 'ГО', 'ИНН'.
    3) Удаляет ненужные столбцы, дубли, строки ниже порога 'ГО'.
    4) Исключает ИНН, которые есть в файле сверки.
    5) Добавляет столбец 'ГО_Месяц_Год', заполняет его.
    6) Переименовывает некоторые столбцы.
    7) Сохраняет итоговые файлы (XLSX и CSV).
    """
    if not main_file_path or not comparison_file_path or not threshold_value.get() or not go_month_year_value.get():
        status_label.config(text="Не все параметры установлены.", fg="red")
        return

    df = pd.read_excel(main_file_path)

    if 'Номер' in df.columns:
        df['Номер'] = df['Номер'].apply(format_number)
    if 'ГО' in df.columns:
        df['ГО'] = df['ГО'].apply(format_number)
    if 'ИНН' in df.columns:
        df['ИНН'] = df['ИНН'].apply(format_number)

    df.at[0, 'A'] = 'Сумма ГО'
    df.at[0, 'G'] = 'Регион поставщика'
    df.at[0, 'I'] = 'Часовой пояс поставщика'
    df.at[0, 'S'] = 'Предмет тендера'
    df.at[0, 'U'] = 'Номер тендера'

    columns_to_delete = [
        "Обеспечение", "Срок поставки", "Область заказчика", "ОКВЭД", "Выручка",
        "Прибыль", "НМЦ", "Конечная цена", "Дата регистрации", "A", "G", "I", "S", "U"
    ]
    df.drop(columns=[col for col in columns_to_delete if col in df.columns], inplace=True)

    remove_duplicates(df, 'ИНН')

    if 'ГО' in df.columns:
        df = filter_by_go(df, float(threshold_value.get()))

    df = remove_inn_duplicates(df, comparison_file_path)

    df.insert(0, 'ГО_Месяц_Год', '')

    fill_go_month_year(df, go_month_year_value.get())

    rename_columns(df)

    output_filename = os.path.basename(main_file_path)
    output_file_path_excel = f' {output_filename}'
    df.to_excel(output_file_path_excel, index=False)

    output_file_path_csv = f' {os.path.splitext(output_filename)[0]}.csv'
    df.to_csv(output_file_path_csv, index=False, encoding='utf-8', sep=',')

    status_label.config(text="Готово", fg="green")
    print(f'Файл сохранен: {output_file_path_excel}')
    print(f'Файл сохранен: {output_file_path_csv}')

def select_main_file():
    """
    Диалог выбора основного Excel-файла.
    """
    global main_file_path
    main_file_path = filedialog.askopenfilename(
        initialdir="",
        title="Выберите основной файл",
        filetypes=[("Excel files", "*.xlsx")]
    )
    main_file_label.config(text=f"Выбранный файл: {main_file_path}")

def select_comparison_file():
    """
    Диалог выбора файла для сверки ИНН.
    """
    global comparison_file_path
    comparison_file_path = filedialog.askopenfilename(
        initialdir="",
        title="Выберите файл для сверки ИНН",
        filetypes=[("Excel files", "*.xlsx")]
    )
    comparison_file_label.config(text=f"Выбранный файл: {comparison_file_path}")

def clear_files():
    """
    Сбрасывает выбранные пути к файлам и поле 'ГО_'.
    """
    global main_file_path, comparison_file_path
    main_file_path = ""
    comparison_file_path = ""
    main_file_label.config(text="Файл не выбран")
    comparison_file_label.config(text="Файл не выбран")
    go_month_year_value.set("ГО_")
    status_label.config(text="Файлы и поле 'ГО_' очищены", fg="blue")

# --- Создаем интерфейс Tkinter ---
root = Tk()
root.title("Обработка файлов")

main_file_path = ""
comparison_file_path = ""

# Кнопки и поля для выбора файлов
main_file_button = Button(root, text="Выбрать основной файл", command=select_main_file)
main_file_button.pack()
main_file_label = Label(root, text="Файл не выбран")
main_file_label.pack()

comparison_file_button = Button(root, text="Выбрать файл для сверки ИНН", command=select_comparison_file)
comparison_file_button.pack()
comparison_file_label = Label(root, text="Файл не выбран")
comparison_file_label.pack()

# Ввод порога для 'ГО'
threshold_label = Label(root, text="Введите пороговое значение для 'ГО':")
threshold_label.pack()
threshold_value = StringVar(value="250000")
threshold_entry = Entry(root, textvariable=threshold_value)
threshold_entry.pack()

# Ввод значения для 'ГО_Месяц_Год'
go_month_year_label = Label(root, text="Введите значение для 'ГО_Месяц_Год':")
go_month_year_label.pack()
go_month_year_value = StringVar(value="ГО_")
go_month_year_entry = Entry(root, textvariable=go_month_year_value)
go_month_year_entry.pack()

# Кнопка «Обработать файлы»
process_button = Button(root, text="Обработать файлы", command=process_files)
process_button.pack()

# Кнопка «Очистить файлы»
clear_button = Button(root, text="Очистить файлы", command=clear_files)
clear_button.pack()

# Статус
status_label = Label(root, text="")
status_label.pack()

root.mainloop()
