"""
Все права защищены (c) 2024.
Данный скрипт обрабатывает Excel-файлы, удаляет информацию по заданным параметрам
и форматирует данные в нужном формате.
Запрещается копирование, распространение или модификация без письменного разрешения автора.
"""

import os
import sys
import json
import pandas as pd
from PyQt5 import QtWidgets, QtCore

CONFIG_PATH = os.path.join(
    os.environ.get('APPDATA', ''), 'LidZalivBitrix', 'config.json'
)


class SettingsDialog(QtWidgets.QDialog):
    """
    Окно «Настройки» для указания путей сохранения/загрузки файлов.
    """
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Настройки')
        self.config = config
        self.init_ui()

    def init_ui(self):
        layout = QtWidgets.QFormLayout()

        # Поле ввода «Директория исходных файлов»
        self.input_dir_edit = QtWidgets.QLineEdit(self.config.get('input_dir', ''))
        self.input_dir_btn = QtWidgets.QPushButton('...')
        self.input_dir_btn.clicked.connect(self.select_input_dir)
        input_dir_layout = QtWidgets.QHBoxLayout()
        input_dir_layout.addWidget(self.input_dir_edit)
        input_dir_layout.addWidget(self.input_dir_btn)
        layout.addRow('Директория исходных файлов:', input_dir_layout)

        # Поле ввода «Директория для XLSX»
        self.xlsx_dir_edit = QtWidgets.QLineEdit(self.config.get('xlsx_output_dir', ''))
        self.xlsx_dir_btn = QtWidgets.QPushButton('...')
        self.xlsx_dir_btn.clicked.connect(self.select_xlsx_dir)
        xlsx_dir_layout = QtWidgets.QHBoxLayout()
        xlsx_dir_layout.addWidget(self.xlsx_dir_edit)
        xlsx_dir_layout.addWidget(self.xlsx_dir_btn)
        layout.addRow('Директория для сохранения XLSX:', xlsx_dir_layout)

        # Поле ввода «Директория для CSV»
        self.csv_dir_edit = QtWidgets.QLineEdit(self.config.get('csv_output_dir', ''))
        self.csv_dir_btn = QtWidgets.QPushButton('...')
        self.csv_dir_btn.clicked.connect(self.select_csv_dir)
        csv_dir_layout = QtWidgets.QHBoxLayout()
        csv_dir_layout.addWidget(self.csv_dir_edit)
        csv_dir_layout.addWidget(self.csv_dir_btn)
        layout.addRow('Директория для сохранения CSV:', csv_dir_layout)

        # Кнопки OK / Cancel
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def select_input_dir(self):
        dir_path = QtWidgets.QFileDialog.getExistingDirectory(self, "Выберите директорию исходных файлов")
        if dir_path:
            self.input_dir_edit.setText(dir_path)

    def select_xlsx_dir(self):
        dir_path = QtWidgets.QFileDialog.getExistingDirectory(self, "Выберите директорию для сохранения XLSX")
        if dir_path:
            self.xlsx_dir_edit.setText(dir_path)

    def select_csv_dir(self):
        dir_path = QtWidgets.QFileDialog.getExistingDirectory(self, "Выберите директорию для сохранения CSV")
        if dir_path:
            self.csv_dir_edit.setText(dir_path)

    def get_settings(self):
        return {
            'input_dir': self.input_dir_edit.text(),
            'xlsx_output_dir': self.xlsx_dir_edit.text(),
            'csv_output_dir': self.csv_dir_edit.text(),
        }


class App(QtWidgets.QMainWindow):
    """
    Основное окно приложения для выбора и обработки Excel-файлов.
    """
    def __init__(self):
        super().__init__()
        self.config = self.load_config()

        self.headers = [
            'Название лида', 'Имя', 'Источник', 'Рабочий телефон', 'Рабочий e-mail',
            'Название компании', 'Адрес', 'Улица, номер дома', 'Комментарий', 'Сумма',
            'Валюта', 'Номер тендера', 'Регион поставщика', 'ИНН', 'ФЗ',
            'Заказчик', 'Ссылка', 'Часовой пояс поставщика', 'Предмет тендера',
            'Дата протокола', 'НМЦ 2', 'СУММА ОБЕСПЕЧЕНИЯ'
        ]

        # Обязательные заголовки, которые нельзя отключать
        self.mandatory_headers = [
            'Название лида', 'Имя', 'Источник', 'Рабочий телефон',
            'Рабочий e-mail', 'Название компании', 'Адрес', 'Улица, номер дома'
        ]

        self.selected_headers = []
        self.files = []
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Обработка XLSX файлов')
        self.resize(600, 400)

        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        layout = QtWidgets.QVBoxLayout(central_widget)

        # Меню
        self.create_menu()

        # Кнопка выбора файлов
        self.select_files_btn = QtWidgets.QPushButton('Выбрать файлы')
        self.select_files_btn.clicked.connect(self.select_files)
        layout.addWidget(self.select_files_btn)

        # Список выбранных файлов
        self.files_list = QtWidgets.QListWidget()
        layout.addWidget(self.files_list)

        # Чекбоксы заголовков
        headers_label = QtWidgets.QLabel('Выберите заголовки:')
        layout.addWidget(headers_label)

        self.checkboxes_layout = QtWidgets.QGridLayout()
        self.checkboxes = []
        for i, header in enumerate(self.headers):
            chk = QtWidgets.QCheckBox(header)
            if header in self.mandatory_headers:
                chk.setChecked(True)
                chk.setEnabled(False)
            self.checkboxes.append(chk)
            self.checkboxes_layout.addWidget(chk, i // 2, i % 2)
        layout.addLayout(self.checkboxes_layout)

        # Кнопка «Обработать файлы»
        self.process_btn = QtWidgets.QPushButton('Обработать файлы')
        self.process_btn.clicked.connect(self.process_files)
        layout.addWidget(self.process_btn)

    def create_menu(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu('Файл')

        settings_action = QtWidgets.QAction('Настройки', self)
        settings_action.triggered.connect(self.open_settings)
        file_menu.addAction(settings_action)

        exit_action = QtWidgets.QAction('Выход', self)
        exit_action.triggered.connect(QtWidgets.qApp.quit)
        file_menu.addAction(exit_action)

    def open_settings(self):
        dialog = SettingsDialog(self.config, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            new_settings = dialog.get_settings()
            self.config.update(new_settings)
            self.save_config()

    def select_files(self):
        options = QtWidgets.QFileDialog.Options()
        start_dir = self.config.get('input_dir', '')
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Выберите файлы", start_dir, "Excel Files (*.xlsx);;All Files (*)", options=options
        )
        if files:
            self.files = files
            self.files_list.clear()
            self.files_list.addItems(files)

    def process_files(self):
        self.selected_headers = [chk.text() for chk in self.checkboxes if chk.isChecked()]
        if not self.selected_headers:
            QtWidgets.QMessageBox.warning(self, "Нет заголовков", "Пожалуйста, выберите хотя бы один заголовок.")
            return
        if not self.files:
            QtWidgets.QMessageBox.warning(self, "Нет файлов", "Пожалуйста, выберите хотя бы один файл.")
            return

        data_frames = []
        for file in self.files:
            try:
                df = pd.read_excel(file, usecols=self.selected_headers)
                data_frames.append(df)
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Ошибка чтения файла", f"Не удалось прочитать {file}: {e}")
                return
        result_df = pd.concat(data_frames, ignore_index=True)

        self.fill_missing_values(result_df)
        self.fix_emails(result_df)
        self.fix_phones(result_df)

        xlsx_output_dir = self.config.get('xlsx_output_dir', '')
        csv_output_dir = self.config.get('csv_output_dir', '')

        if not xlsx_output_dir or not os.path.isdir(xlsx_output_dir):
            xlsx_output_dir = QtWidgets.QFileDialog.getExistingDirectory(
                self, "Выберите директорию для сохранения XLSX"
            )
            if not xlsx_output_dir:
                QtWidgets.QMessageBox.warning(
                    self, "Нет папки", "Пожалуйста, выберите директорию для XLSX."
                )
                return
            self.config['xlsx_output_dir'] = xlsx_output_dir
            self.save_config()

        if not csv_output_dir or not os.path.isdir(csv_output_dir):
            csv_output_dir = QtWidgets.QFileDialog.getExistingDirectory(
                self, "Выберите директорию для сохранения CSV"
            )
            if not csv_output_dir:
                QtWidgets.QMessageBox.warning(
                    self, "Нет папки", "Пожалуйста, выберите директорию для CSV."
                )
                return
            self.config['csv_output_dir'] = csv_output_dir
            self.save_config()

        xlsx_filename = self.generate_unique_filename(xlsx_output_dir, 'result', '.xlsx')
        csv_filename = self.generate_unique_filename(csv_output_dir, 'result', '.csv')

        xlsx_path = os.path.join(xlsx_output_dir, xlsx_filename)
        csv_path = os.path.join(csv_output_dir, csv_filename)

        try:
            with pd.ExcelWriter(xlsx_path, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False)
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                if 'Рабочий телефон' in result_df.columns:
                    phone_col_idx = result_df.columns.get_loc('Рабочий телефон')
                    text_format = workbook.add_format({'num_format': '@'})
                    worksheet.set_column(phone_col_idx, phone_col_idx, None, text_format)

            result_df.to_csv(csv_path, index=False, sep=',', encoding='utf-8')

            QtWidgets.QMessageBox.information(
                self, "Готово", f"Файлы сохранены:\nXLSX: {xlsx_path}\nCSV: {csv_path}"
            )
        except Exception as e:
            QtWidgets.QMessageBox.warning(
                self, "Ошибка сохранения", f"Не удалось сохранить файлы: {e}"
            )

    def fill_missing_values(self, df: pd.DataFrame):
        """
        Простейшая логика заполнения пустых значений.
        """
        for col in self.selected_headers:
            col_index = df.columns.get_loc(col)
            if col_index > 0:
                left_col = df.columns[col_index - 1]
                df[col] = df[col].fillna(df[left_col].apply(lambda x: 0 if pd.notna(x) else x))
            else:
                df[col] = df[col].fillna(0)

    def fix_emails(self, df: pd.DataFrame):
        """
        Коррекция столбца «Рабочий e-mail» (при необходимости).
        """
        if 'Рабочий e-mail' not in df.columns:
            return

        def fix_email(email):
            if pd.isna(email) or not isinstance(email, str):
                return email
            email = email.strip()
            if '@' not in email:
                email += '@mail.ru'
            else:
                username, domain = email.split('@', 1)
                if domain.endswith('.'):
                    domain += 'ru'
                elif domain.endswith('.r'):
                    domain += 'u'
                elif domain == 'mail' or domain.startswith('mail.'):
                    if not domain.endswith('.ru'):
                        domain += '.ru'
                elif domain == 'gmail' or domain.startswith('gmail.'):
                    if not domain.endswith('.com'):
                        domain += '.com'
                else:
                    if not (domain.endswith('.ru') or domain.endswith('.com')):
                        domain += '.com'
                email = username + '@' + domain
            return email

        df['Рабочий e-mail'] = df['Рабочий e-mail'].apply(fix_email)

    def fix_phones(self, df: pd.DataFrame):
        """
        Коррекция столбца «Рабочий телефон».
        """
        if 'Рабочий телефон' not in df.columns:
            return

        def fix_phone_number(phone):
            if pd.isna(phone):
                return phone
            if isinstance(phone, float):
                phone = str(int(phone))
            elif not isinstance(phone, str):
                return phone

            phones = phone.split(',')
            fixed_phones = []
            for p in phones:
                p = p.strip()
                digits = ''.join(filter(str.isdigit, p))
                if not digits:
                    continue
                if digits.startswith('8'):
                    digits = '+7' + digits[1:]
                elif digits.startswith('7'):
                    digits = '+' + digits
                elif digits.startswith('9'):
                    digits = '+7' + digits
                fixed_phones.append(digits)
            return ', '.join(fixed_phones) if fixed_phones else None

        df['Рабочий телефон'] = df['Рабочий телефон'].apply(fix_phone_number)

    def generate_unique_filename(self, directory, base_name, extension):
        index = 1
        filename = f"{base_name}{extension}"
        while os.path.exists(os.path.join(directory, filename)):
            filename = f"{base_name}_{index}{extension}"
            index += 1
        return filename

    def load_config(self):
        """
        Чтение JSON-конфига.
        """
        if not os.path.exists(CONFIG_PATH):
            os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
            return {}
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)

    def save_config(self):
        """
        Сохранение JSON-конфига.
        """
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)


def main():
    app = QtWidgets.QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
