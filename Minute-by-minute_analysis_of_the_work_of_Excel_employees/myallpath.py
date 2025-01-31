"""
Все права защищены (c) 2024. 
Данный скрипт предназначен для работы с папками.
Автор кода не предоставляет прав на использование или распространение данного ПО.
"""

import tkinter as tk
from tkinter import filedialog


def select_files():
    """
    :return: Кортеж, содержащий пути к выбранному текстовому файлу со сведениями о сотруднике,
    выбранному ZIP-архиву, выбранному файлу Excel с историей звонков и выбранному каталогу для вывода результатов.
    """
    root = tk.Tk()
    root.withdraw()
    # Выбор текстового файла
    people = filedialog.askopenfilename(
        title="Выберите файл с сотрудниками",
        filetypes=[("Text Files", "*.txt")]
    )
    # Выбор zip-архива
    zip_path = filedialog.askopenfilename(
        title="Выберите ZIP файл",
        filetypes=[("ZIP Files", "*.zip")]
    )
    # Выбор файла с историей звонков
    zvonki = filedialog.askopenfilename(
        title="Выберите файл с историей звонков",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    # Выбор директории для вывода итогов
    output_excel_path = filedialog.askdirectory(
        title="Выберите директорию для итогов"
    )
    print(f"Файл с сотрудниками: {people}")
    print(f"ZIP архив: {zip_path}")
    print(f"Файл с историей звонков: {zvonki}")
    print(f"Директория для итогов: {output_excel_path}")
    return people, zip_path, zvonki, output_excel_path
