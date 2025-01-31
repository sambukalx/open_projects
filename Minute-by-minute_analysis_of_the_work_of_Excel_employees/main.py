"""
Все права защищены (c) 2024. 
Данный скрипт предназначен для запуска проекта.
Автор кода не предоставляет прав на использование или распространение данного ПО.
"""

import io
import os
import sys
import psutil
import time
import logging
import shutil
import signal
import json
import pythoncom
import traceback

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject, QUrl
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QGridLayout, QWidget,
    QFileDialog, QProgressBar, QMessageBox, QLineEdit, QTextEdit, QStackedWidget, QScrollArea,
    QHBoxLayout, QListWidget, QListWidgetItem, QInputDialog, QRadioButton, QButtonGroup, QComboBox)
from PyQt6.QtGui import QAction, QKeySequence, QShortcut, QFont, QDesktopServices
import win32com.client

from colorama import Fore, Style, init
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
from myalldata import (
    replacements as default_replacements,
    url_mapping as default_url_mapping,
    employee_timezones as default_employee_timezones,
    company_timezone as default_company_timezone
)
from infoStah_toexcel import *
from format import *
from dost_file import *

from bitrix_normolize import *

init(autoreset=False, wrap=False)


def exception_hook(exctype, value, tb):
    traceback_str = ''.join(traceback.format_exception(exctype, value, tb))
    print(traceback_str)
    QMessageBox.critical(None, "Unhandled Exception", traceback_str)
    sys.exit(1)


sys.excepthook = exception_hook


def get_appdata_dir():
    """
    Возвращает путь к каталогу, в котором хранятся данные приложения.
    Если каталог не существует, он создает его.
    :return: Путь к каталогу AnalyzeOP в папке AppData\Roaming пользователя.
    """
    appdata_dir = os.getenv('APPDATA')
    analyzeop_dir = os.path.join(appdata_dir, 'AnalyzeOP')

    if not os.path.exists(analyzeop_dir):
        os.makedirs(analyzeop_dir)

    return analyzeop_dir


def get_error_log_path():
    """
    Получает полный путь к файлу журнала ошибок.
    :return: Путь к файлу журнала ошибок в виде строки.
    """
    analyzeop_dir = get_appdata_dir()
    return os.path.join(analyzeop_dir, 'errors.log')


def get_config_path():
    """
    :return: Полный путь к файлу конфигурации config.json, расположенному в каталоге данных приложения.
    """
    analyzeop_dir = get_appdata_dir()
    return os.path.join(analyzeop_dir, 'config.json')


logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

console_handler = logging.StreamHandler(sys.__stdout__)
console_handler.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)

logger.addHandler(console_handler)

error_log_path = get_error_log_path()

error_file_handler = logging.FileHandler(error_log_path, encoding='utf-8')
error_file_handler.setLevel(logging.ERROR)
error_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s\n%(exc_info)s')
error_file_handler.setFormatter(error_formatter)
logger.addHandler(error_file_handler)


class BufferHandler(logging.Handler):
    """
    Класс для обработки журналирования путем буферизации записей журнала в памяти.
    Methods
    -------
    __init__():
        Инициализирует буфер для хранения сообщений журнала.
    emit(record):
         Форматирует и добавляет запись журнала в буфер.
    """

    def __init__(self):
        """
        Инициализирует новый экземпляр класса.
        Этот конструктор устанавливает начальное состояние объекта, вызывая метод
        конструктор суперкласса и инициализация пустого списка буферов.
        """
        super().__init__()
        self.buffer = []

    def emit(self, record):
        """
        :param record: Содержит всю информацию, относящуюся к регистрируемому событию.
        :return: None. Этот метод добавляет отформатированную запись журнала в буфер.
        """
        msg = self.format(record)
        self.buffer.append(msg)


class QtHandler(QObject, BufferHandler):
    """
    QtHandler — это специальный обработчик журналирования, который объединяет вывод журнала с текстовым виджетом PyQt5.
    Этот обработчик наследуется как от QObject, так и от BufferHandler, что позволяет ему взаимодействовать.
    с механизмом сигналов и слотов Qt, одновременно выполняя стандартные операции регистрации.
    Attributes:
        write_signal (pyqtSignal): сигнал, используемый для обновления текстового виджета новым сообщением журнала.
        text_widget (QTextEdit): текстовый виджет PyQt5, в котором будут отображаться сообщения журнала.
    Methods:
        __init__(self, text_widget=None):
            Инициализирует QtHandler с помощью необязательного text_widget. Настраивает сигнальное соединение.
        emit(self, record):
            Создает запись журнала и обновляет text_widget, если он предоставлен.
        update_gui(self, msg):
            Добавляет сообщение журнала в text_widget и настраивает
            полосу прокрутки для отображения последнего сообщения.
    """
    write_signal = pyqtSignal(str)

    def __init__(self, text_widget=None):
        """
        :param text_widget: Виджет, который будет отображать текстовый вывод.
        Это может быть любой QWidget, поддерживающий отображение текста.
        """
        QObject.__init__(self)
        BufferHandler.__init__(self)
        self.text_widget = text_widget
        self.write_signal.connect(self.update_gui)

    def emit(self, record):
        """
        :param Record: Запись журнала, которая будет создана.
        Он содержит всю информацию, относящуюся к регистрируемому событию,
        такую как сообщение, уровень, временная метка и т. д.
        :return: None. Эта функция не возвращает никакого значения.
        Он обрабатывает запись журнала и обновляет text_widget, если он доступен.
        """
        BufferHandler.emit(self, record)
        if self.text_widget:
            msg = self.format(record)
            self.write_signal.emit(msg)

    def update_gui(self, msg):
        """
        :param msg: Сообщение, которое будет добавлено к текстовому виджету.
        :return: None
        """
        if self.text_widget:
            self.text_widget.append(msg)
            self.text_widget.verticalScrollBar().setValue(self.text_widget.verticalScrollBar().maximum())


buffer_handler = QtHandler()
buffer_handler.setLevel(logging.DEBUG)
buffer_handler.setFormatter(formatter)
logger.addHandler(buffer_handler)


class EditReplacementsPage(QWidget):
    """
    Этот класс представляет QWidget, который позволяет редактировать замены сотрудников.
    Интерфейс предоставляет текстовый редактор для изменения представления замен в формате JSON.
    а также кнопки для сохранения изменений, возврата к значениям по умолчанию или возврата к
    предыдущий экран.
    Methods
    -------
    __init__(self, parent)
        Инициализирует виджет и его компоненты пользовательского интерфейса.
    init_ui(self)
        Настраивает элементы пользовательского интерфейса, включая макет, текстовый редактор и кнопки.
    save_replacements(self)
        Сохраняет измененные замены в настройках родительского объекта.
    reset_to_default(self)
        Сбрасывает замены к значениям по умолчанию и обновляет текстовый редактор.
    go_back(self)
        Возвращает к предыдущему экрану составного виджета родительского объекта.
    """

    def __init__(self, parent):
        """
        Parameters
        ----------
        parent : object
            Родительский виджет для этого экземпляра класса
        """
        super().__init__(parent)
        self.remove_source_button = None
        self.add_source_button = None
        self.target_list_widget = None
        self.source_list_widget = None
        self.back_button = None
        self.reset_button = None
        self.text_edit = None
        self.save_button = None
        self.add_target_button = None
        self.remove_target_button = None
        self.parent = parent
        self.search_edit = None
        self.all_targets = []
        self.init_ui()

    def init_ui(self):
        """
        Инициализирует пользовательский интерфейс для редактирования сопоставлений сотрудников.
        Этот метод устанавливает основной макет с помощью QVBoxLayout и включает в себя следующие элементы:
        - Ярлык заголовка по центру с жирным стилем.
        — Виджет QTextEdit, загружающий текущий словарь замен сотрудников в формате JSON.
        - Три виджета QPushButton («Сохранить», «Восстановить настройки по умолчанию» и «Назад»)
        с соответствующими макетами и подключениями событий щелчка.
        """
        main_layout = QVBoxLayout()

        # Поле поиска для сотрудников
        search_layout = QHBoxLayout()
        search_label = QLabel("Поиск сотрудника:", self)
        self.search_edit = QLineEdit(self)
        self.search_edit.textChanged.connect(self.filter_targets)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_edit)
        main_layout.addLayout(search_layout)

        # Основной макет
        layout = QHBoxLayout()

        # Список сотрудников (слева)
        self.target_list_widget = QListWidget(self)
        self.target_list_widget.setFixedWidth(200)
        self.target_list_widget.itemClicked.connect(self.display_sources)
        layout.addWidget(self.target_list_widget)

        # Кнопки для управления сотрудниками
        target_buttons_layout = QVBoxLayout()
        self.add_target_button = QPushButton("Добавить сотрудника", self)
        self.remove_target_button = QPushButton("Удалить сотрудника", self)
        target_buttons_layout.addWidget(self.add_target_button)
        target_buttons_layout.addWidget(self.remove_target_button)
        self.add_target_button.clicked.connect(self.add_target)
        self.remove_target_button.clicked.connect(self.remove_target)
        layout.addLayout(target_buttons_layout)

        # Добавляем стрелку
        arrow_label = QLabel("→", self)
        arrow_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        arrow_label.setFixedWidth(30)
        layout.addWidget(arrow_label)

        # Список соответствий (справа)
        self.source_list_widget = QListWidget(self)
        self.source_list_widget.setFixedWidth(200)
        layout.addWidget(self.source_list_widget)

        # Кнопки для управления соответствиями
        source_buttons_layout = QVBoxLayout()
        self.add_source_button = QPushButton("Добавить соответствие", self)
        self.remove_source_button = QPushButton("Удалить соответствие", self)
        source_buttons_layout.addWidget(self.add_source_button)
        source_buttons_layout.addWidget(self.remove_source_button)
        self.add_source_button.clicked.connect(self.add_source)
        self.remove_source_button.clicked.connect(self.remove_source)
        layout.addLayout(source_buttons_layout)

        # Кнопки управления
        control_buttons_layout = QHBoxLayout()
        self.save_button = QPushButton("Сохранить", self)
        self.back_button = QPushButton("Назад", self)
        control_buttons_layout.addWidget(self.save_button)
        control_buttons_layout.addWidget(self.back_button)
        self.save_button.clicked.connect(self.save_replacements)
        self.back_button.clicked.connect(self.go_back)

        main_layout.addLayout(layout)
        main_layout.addLayout(control_buttons_layout)

        self.setLayout(main_layout)

        self.load_targets()

    def load_targets(self):
        self.all_targets = list(self.parent.replacements.keys())
        self.update_target_list()

    def update_target_list(self, filtered_targets=None):
        self.target_list_widget.clear()
        if filtered_targets is None:
            targets_to_show = self.all_targets
        else:
            targets_to_show = filtered_targets
        for target in sorted(targets_to_show):
            self.target_list_widget.addItem(target)
        if self.target_list_widget.count() > 0:
            self.target_list_widget.setCurrentRow(0)
            current_item = self.target_list_widget.currentItem()
            self.display_sources(current_item)
        else:
            self.source_list_widget.clear()

    def filter_targets(self, text):
        text = text.lower()
        filtered_targets = []
        for target in self.all_targets:
            target_lower = target.lower()
            sources_lower = [source.lower() for source in self.parent.replacements[target]]
            if text in target_lower or any(text in source for source in sources_lower):
                filtered_targets.append(target)
        self.update_target_list(filtered_targets)

    def display_sources(self, item):
        if item:
            target_name = item.text()
            self.source_list_widget.clear()
            sources = self.parent.replacements.get(target_name, [])
            for source in sources:
                self.source_list_widget.addItem(source)
        else:
            self.source_list_widget.clear()

    def add_target(self):
        text, ok = QInputDialog.getText(self, 'Добавить сотрудника', 'Введите нового сотрудника:')
        if ok:
            text = text.strip()
            if text:
                if text not in self.parent.replacements:
                    self.parent.replacements[text] = []
                    self.all_targets.append(text)
                    self.update_target_list()
                else:
                    QMessageBox.warning(self, "Предупреждение", "Такой сотрудник уже существует.")
            else:
                QMessageBox.warning(self, "Предупреждение", "Имя сотрудника не может быть пустым.")
        else:
            QMessageBox.warning(self, "Предупреждение", "Операция добавления отменена.")

    def remove_target(self):
        item = self.target_list_widget.currentItem()
        if item:
            target_name = item.text()
            confirmation = QMessageBox.question(
                self,
                "Подтверждение удаления",
                f"Вы уверены, что хотите удалить сотрудника '{target_name}' и все его соответствия?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if confirmation == QMessageBox.StandardButton.Yes:
                try:
                    del self.parent.replacements[target_name]
                    self.all_targets.remove(target_name)
                    self.update_target_list()
                    self.source_list_widget.clear()
                except Exception as e:
                    QMessageBox.critical(self, "Ошибка", f"Не удалось удалить сотрудника: {str(e)}")
            else:
                QMessageBox.information(self, "Отмена", "Удаление отменено.")
        else:
            QMessageBox.warning(self, "Предупреждение", "Выберите сотрудника для удаления.")

    def add_source(self):
        target_item = self.target_list_widget.currentItem()
        if target_item:
            target_name = target_item.text()
            text, ok = QInputDialog.getText(self, 'Добавить соответствие', 'Введите заменяемое имя:')
            if ok:
                text = text.strip()
                if text:
                    if text not in self.parent.replacements[target_name]:
                        self.parent.replacements[target_name].append(text)
                        self.display_sources(target_item)
                    else:
                        QMessageBox.warning(self, "Предупреждение", "Такое соответствие уже существует.")
                else:
                    QMessageBox.warning(self, "Предупреждение", "Имя не может быть пустым.")
            else:
                QMessageBox.warning(self, "Предупреждение", "Операция добавления отменена.")
        else:
            QMessageBox.warning(self, "Предупреждение", "Выберите сотрудника.")

    def remove_source(self):
        target_item = self.target_list_widget.currentItem()
        source_item = self.source_list_widget.currentItem()
        if target_item and source_item:
            target_name = target_item.text()
            source_name = source_item.text()
            confirmation = QMessageBox.question(
                self,
                "Подтверждение удаления",
                f"Вы уверены, что хотите удалить соответствие '{source_name}'?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if confirmation == QMessageBox.StandardButton.Yes:
                try:
                    self.parent.replacements[target_name].remove(source_name)
                    self.display_sources(target_item)
                except ValueError:
                    QMessageBox.warning(self, "Ошибка", f"Соответствие '{source_name}' не найдено.")
                except Exception as e:
                    QMessageBox.critical(self, "Ошибка", f"Не удалось удалить соответствие: {str(e)}")
            else:
                QMessageBox.information(self, "Отмена", "Удаление отменено.")
        else:
            QMessageBox.warning(self, "Предупреждение", "Выберите соответствие для удаления.")

    def save_replacements(self):
        try:
            self.parent.save_settings()
            QMessageBox.information(self, "Сохранение", "Изменения сохранены.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить изменения: {str(e)}")

    def go_back(self):
        self.parent.stacked_widget.setCurrentWidget(self.parent.main_page)
        self.parent.stacked_widget.removeWidget(self)

    def set_disabled(self, disabled):
        self.source_list_widget.setDisabled(disabled)
        self.target_list_widget.setDisabled(disabled)
        self.add_source_button.setDisabled(disabled)
        self.remove_source_button.setDisabled(disabled)
        self.add_target_button.setDisabled(disabled)
        self.remove_target_button.setDisabled(disabled)
        self.save_button.setDisabled(disabled)
        # Кнопку "Назад" оставляем активной


class EditUrlMappingPage(QWidget):
    """
    EditUrlMappingPage — это подкласс QWidget, который предоставляет
    интерфейс для редактирования сопоставлений URL-адресов.
    Methods
    -------
    __init__(self, parent)
        Инициализирует экземпляр EditUrlMappingPage и настраивает компоненты пользовательского интерфейса.
    init_ui (self)
        Настраивает пользовательский интерфейс, включая метки, текстовый редактор и кнопки.
    save_url_mapping (self)
        Сохраняет сопоставления URL-адресов в родительской конфигурации, если входные данные действительны в формате JSON.
    reset_to_default (self)
        Сбрасывает сопоставления URL-адресов к значениям по умолчанию и обновляет текстовый редактор.
    go_back (self
        Возвращает на главную страницу родительского составного виджета.
    """

    def __init__(self, parent):
        """
        Parameters
        ----------
        parent : object
            Родительский компонент для этого элемента пользовательского интерфейса.
        """
        super().__init__()
        self.add_target_button = None
        self.target_list_widget = None
        self.remove_source_button = None
        self.add_source_button = None
        self.source_list_widget = None
        self.remove_target_button = None
        self.save_button = None
        self.reset_button = None
        self.back_button = None
        self.text_edit = None
        self.parent = parent
        self.search_edit = None
        self.all_targets = []
        self.init_ui()

    def init_ui(self):
        """
        Инициализирует пользовательский интерфейс для редактирования переписок сайта.
        Создает макет и добавляет метку заголовка,
        текстовый редактор для сопоставления URL-адресов и кнопки управления.
        Parameters None
        Returns None
        """
        main_layout = QVBoxLayout()

        # Поле поиска для URL
        search_layout = QHBoxLayout()
        search_label = QLabel("Поиск URL:", self)
        self.search_edit = QLineEdit(self)
        self.search_edit.textChanged.connect(self.filter_targets)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_edit)
        main_layout.addLayout(search_layout)

        # Основной макет
        layout = QHBoxLayout()

        # Список целевых URL (слева)
        self.target_list_widget = QListWidget(self)
        self.target_list_widget.setFixedWidth(300)
        self.target_list_widget.itemClicked.connect(self.display_sources)
        layout.addWidget(self.target_list_widget)

        # Кнопки для управления целевыми URL
        target_buttons_layout = QVBoxLayout()
        self.add_target_button = QPushButton("Добавить URL", self)
        self.remove_target_button = QPushButton("Удалить URL", self)
        target_buttons_layout.addWidget(self.add_target_button)
        target_buttons_layout.addWidget(self.remove_target_button)
        self.add_target_button.clicked.connect(self.add_target)
        self.remove_target_button.clicked.connect(self.remove_target)
        layout.addLayout(target_buttons_layout)

        # Добавляем стрелку
        arrow_label = QLabel("→", self)
        arrow_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        arrow_label.setFixedWidth(30)
        layout.addWidget(arrow_label)

        # Список исходных URL (справа)
        self.source_list_widget = QListWidget(self)
        self.source_list_widget.setFixedWidth(300)
        layout.addWidget(self.source_list_widget)

        # Кнопки для управления исходными URL
        source_buttons_layout = QVBoxLayout()
        self.add_source_button = QPushButton("Добавить URL", self)
        self.remove_source_button = QPushButton("Удалить URL", self)
        source_buttons_layout.addWidget(self.add_source_button)
        source_buttons_layout.addWidget(self.remove_source_button)
        self.add_source_button.clicked.connect(self.add_source)
        self.remove_source_button.clicked.connect(self.remove_source)
        layout.addLayout(source_buttons_layout)

        # Кнопки управления
        control_buttons_layout = QHBoxLayout()
        self.save_button = QPushButton("Сохранить", self)
        self.back_button = QPushButton("Назад", self)
        control_buttons_layout.addWidget(self.save_button)
        control_buttons_layout.addWidget(self.back_button)
        self.save_button.clicked.connect(self.save_url_mapping)
        self.back_button.clicked.connect(self.go_back)

        main_layout.addLayout(layout)
        main_layout.addLayout(control_buttons_layout)

        self.setLayout(main_layout)

        self.load_targets()

    def load_targets(self):
        self.all_targets = list(self.parent.url_mapping.keys())
        self.update_target_list()

    def update_target_list(self, filtered_targets=None):
        self.target_list_widget.clear()
        if filtered_targets is None:
            targets_to_show = self.all_targets
        else:
            targets_to_show = filtered_targets
        for target in sorted(targets_to_show):
            self.target_list_widget.addItem(target)
        if self.target_list_widget.count() > 0:
            self.target_list_widget.setCurrentRow(0)
            current_item = self.target_list_widget.currentItem()
            self.display_sources(current_item)
        else:
            self.source_list_widget.clear()

    def filter_targets(self, text):
        text = text.lower()
        filtered_targets = []
        for target in self.all_targets:
            target_lower = target.lower()
            sources_lower = [source.lower() for source in self.parent.url_mapping[target]]
            if text in target_lower or any(text in source for source in sources_lower):
                filtered_targets.append(target)
        self.update_target_list(filtered_targets)

    def display_sources(self, item):
        if item:
            target_name = item.text()
            self.source_list_widget.clear()
            sources = self.parent.url_mapping.get(target_name, [])
            for source in sources:
                self.source_list_widget.addItem(source)
        else:
            self.source_list_widget.clear()

    def add_target(self):
        text, ok = QInputDialog.getText(self, 'Добавить URL', 'Введите новый URL:')
        if ok:
            text = text.strip()
            if text:
                if text not in self.parent.url_mapping:
                    self.parent.url_mapping[text] = []
                    self.all_targets.append(text)
                    self.update_target_list()
                else:
                    QMessageBox.warning(self, "Предупреждение", "Такой URL уже существует.")
            else:
                QMessageBox.warning(self, "Предупреждение", "URL не может быть пустым.")
        else:
            QMessageBox.warning(self, "Предупреждение", "Операция добавления отменена.")

    def remove_target(self):
        item = self.target_list_widget.currentItem()
        if item:
            target_name = item.text()
            confirmation = QMessageBox.question(
                self, "Подтверждение удаления",
                f"Вы уверены, что хотите удалить URL '{target_name}' и все его соответствия?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if confirmation == QMessageBox.StandardButton.Yes:
                try:
                    del self.parent.url_mapping[target_name]
                    self.all_targets.remove(target_name)
                    self.update_target_list()
                    self.source_list_widget.clear()
                except Exception as e:
                    QMessageBox.critical(self, "Ошибка", f"Не удалось удалить URL: {str(e)}")
            else:
                QMessageBox.information(self, "Отмена", "Удаление отменено.")
        else:
            QMessageBox.warning(self, "Предупреждение", "Выберите URL для удаления.")

    def add_source(self):
        target_item = self.target_list_widget.currentItem()
        if target_item:
            target_name = target_item.text()
            text, ok = QInputDialog.getText(self, 'Добавить URL', 'Введите URL для замены:')
            if ok:
                text = text.strip()
                if text:
                    if text not in self.parent.url_mapping[target_name]:
                        self.parent.url_mapping[target_name].append(text)
                        self.display_sources(target_item)
                    else:
                        QMessageBox.warning(self, "Предупреждение", "Такой URL уже существует.")
                else:
                    QMessageBox.warning(self, "Предупреждение", "URL не может быть пустым.")
            else:
                QMessageBox.warning(self, "Предупреждение", "Операция добавления отменена.")
        else:
            QMessageBox.warning(self, "Предупреждение", "Выберите целевой URL.")

    def remove_source(self):
        target_item = self.target_list_widget.currentItem()
        source_item = self.source_list_widget.currentItem()
        if target_item and source_item:
            target_name = target_item.text()
            source_name = source_item.text()
            try:
                self.parent.url_mapping[target_name].remove(source_name)
                self.display_sources(target_item)
            except ValueError:
                QMessageBox.warning(self, "Ошибка", f"URL '{source_name}' не найден.")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось удалить URL: {str(e)}")
        else:
            QMessageBox.warning(self, "Предупреждение", "Выберите URL для удаления.")

    def save_url_mapping(self):
        try:
            self.parent.save_settings()
            QMessageBox.information(self, "Сохранение", "Изменения сохранены.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить изменения: {str(e)}")

    def go_back(self):
        self.parent.stacked_widget.setCurrentWidget(self.parent.main_page)
        self.parent.stacked_widget.removeWidget(self)


class EditEmployeeTimezonesPage(QWidget):
    """
    Класс QWidget для редактирования часовых поясов сотрудников.
    Parameters
    ----------
    parent: QWidget
        Родительский виджет для этого компонента пользовательского интерфейса.
    """

    def __init__(self, parent):
        """
        Parameters
        ----------
        parent : object
            Родительский виджет, которому принадлежит этот виджет.
        """
        super().__init__(parent)
        self.save_timezone_button = None
        self.timezone_combo = None
        self.remove_employee_button = None
        self.add_employee_button = None
        self.employee_list_widget = None
        self.search_edit = None
        self.all_employees = None
        self.back_button = None
        self.reset_button = None
        self.text_edit = None
        self.save_button = None
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        """
        Инициализирует пользовательский интерфейс для редактирования часовых поясов сотрудников.
        Этот метод устанавливает макет,
        включая метку заголовка, текстовый редактор и кнопки для сохранения, сброса и возврата.
        Текстовый редактор заполняется текущими часовыми поясами сотрудников, сериализованными в формат JSON.
        Attributes
        ----------
        layout : QVBoxLayout
            Основная вертикальная компоновка компонентов пользовательского интерфейса.
        title_label: QLabel
            Ярлык, отображающий название редактора.
        text_edit: QTextEdit
            Текстовый редактор для отображения и редактирования JSON-представления часовых поясов сотрудников.
        button_layout: QHBoxLayout
            Горизонтальное расположение кнопок действий.
        save_button : QPushButton
            Кнопка сохранения изменений часовых поясов сотрудников.
        reset_button: QPushButton
            Кнопка сброса часовых поясов сотрудников на значения по умолчанию.
        back_button : QPushButton
            Кнопка возврата к предыдущему экрану.
        """
        main_layout = QVBoxLayout()

        title_label = QLabel("Редактирование часовых поясов сотрудников")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title_label)

        # Поле поиска
        search_layout = QHBoxLayout()
        search_label = QLabel("Поиск:", self)
        self.search_edit = QLineEdit(self)
        self.search_edit.textChanged.connect(self.filter_employees)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_edit)
        main_layout.addLayout(search_layout)

        # Создаем основной макет
        layout = QHBoxLayout()

        # Список сотрудников (слева)
        self.employee_list_widget = QListWidget(self)
        self.employee_list_widget.setFixedWidth(200)
        self.employee_list_widget.itemClicked.connect(self.display_timezone)
        layout.addWidget(self.employee_list_widget)

        # Кнопки для управления сотрудниками
        employee_buttons_layout = QVBoxLayout()
        self.add_employee_button = QPushButton("Добавить сотрудника", self)
        self.remove_employee_button = QPushButton("Удалить сотрудника", self)
        employee_buttons_layout.addWidget(self.add_employee_button)
        employee_buttons_layout.addWidget(self.remove_employee_button)
        self.add_employee_button.clicked.connect(self.add_employee)
        self.remove_employee_button.clicked.connect(self.remove_employee)
        layout.addLayout(employee_buttons_layout)

        # Добавляем стрелку
        arrow_label = QLabel("→", self)
        arrow_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        arrow_label.setFixedWidth(30)
        layout.addWidget(arrow_label)

        # Выпадающий список для выбора часового пояса (справа)
        self.timezone_combo = QComboBox(self)
        timezones = [str(i) for i in range(-12, 15)] + [f"{i}.5" for i in range(-12, 14)]
        timezones = sorted(set(timezones), key=lambda x: float(x))
        self.timezone_combo.addItems(timezones)
        self.timezone_combo.setEnabled(False)
        layout.addWidget(self.timezone_combo)

        # Кнопка для сохранения часового пояса
        self.save_timezone_button = QPushButton("Сохранить часовой пояс", self)
        self.save_timezone_button.setEnabled(False)
        self.save_timezone_button.clicked.connect(self.save_timezone)

        # Кнопки управления
        control_buttons_layout = QHBoxLayout()
        self.save_button = QPushButton("Сохранить все изменения", self)
        self.reset_button = QPushButton("Вернуть по умолчанию", self)
        self.back_button = QPushButton("Назад", self)
        control_buttons_layout.addWidget(self.save_button)
        control_buttons_layout.addWidget(self.reset_button)
        control_buttons_layout.addWidget(self.back_button)
        self.save_button.clicked.connect(self.save_employee_timezones)
        self.reset_button.clicked.connect(self.reset_to_default)
        self.back_button.clicked.connect(self.go_back)

        main_layout.addLayout(layout)
        main_layout.addWidget(self.save_timezone_button)
        main_layout.addLayout(control_buttons_layout)

        self.setLayout(main_layout)

        self.load_employees()

    def load_employees(self):
        """
        Загружает список сотрудников из настроек и отображает их в списке.
        """
        try:
            if self.parent.people:
                try:
                    with open(self.parent.people, 'r', encoding='utf-8') as file:
                        content = file.read()
                        employee_names = [line.strip() for line in content.splitlines() if
                                          line.strip() and not line.strip().endswith(':')]
                        for name in employee_names:
                            if self.parent.employee_timezones is None:
                                self.parent.employee_timezones = {}
                            if name not in self.parent.employee_timezones:
                                self.parent.employee_timezones[name] = 0
                except Exception as e:
                    QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить файл сотрудников: {str(e)}")

            if self.parent.employee_timezones is None:
                self.parent.employee_timezones = {}
            self.all_employees = list(self.parent.employee_timezones.keys())
            self.update_employee_list()
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить сотрудников: {str(e)}")

    def update_employee_list(self, filtered_employees=None):
        """
        Обновляет отображение списка сотрудников.
        """
        self.employee_list_widget.clear()
        if filtered_employees is None:
            employees_to_show = self.all_employees
        else:
            employees_to_show = filtered_employees
        for employee in sorted(employees_to_show):
            self.employee_list_widget.addItem(employee)

    def filter_employees(self, text):
        """
        Фильтрует список сотрудников по введенному тексту.
        """
        filtered_employees = [name for name in self.all_employees if text.lower() in name.lower()]
        self.update_employee_list(filtered_employees)

    def display_timezone(self, item):
        """
        Отображает часовой пояс выбранного сотрудника в комбобоксе.
        """
        employee_name = item.text()
        timezone = self.parent.employee_timezones.get(employee_name, 0)
        index = self.timezone_combo.findText(str(timezone))
        if index != -1:
            self.timezone_combo.setCurrentIndex(index)
        else:
            self.timezone_combo.setCurrentIndex(self.timezone_combo.findText("0"))
        self.timezone_combo.setEnabled(True)
        self.save_timezone_button.setEnabled(True)

    def add_employee(self):
        """
        Добавляет нового сотрудника в список.
        """
        text, ok = QInputDialog.getText(self, 'Добавить сотрудника', 'Введите имя сотрудника:')
        if ok and text:
            if text.strip() == "":
                QMessageBox.warning(self, "Предупреждение", "Имя сотрудника не может быть пустым.")
                return
            if text not in self.parent.employee_timezones:
                self.parent.employee_timezones[text] = 0  # Устанавливаем часовой пояс по умолчанию
                self.all_employees.append(text)
                self.update_employee_list()
            else:
                QMessageBox.warning(self, "Предупреждение", "Такой сотрудник уже существует.")

    def remove_employee(self):
        """
        Удаляет выбранного сотрудника из списка.
        """
        item = self.employee_list_widget.currentItem()
        if item:
            employee_name = item.text()
            del self.parent.employee_timezones[employee_name]
            self.all_employees.remove(employee_name)
            self.update_employee_list()
            self.timezone_combo.setEnabled(False)
            self.save_timezone_button.setEnabled(False)
        else:
            QMessageBox.warning(self, "Предупреждение", "Выберите сотрудника для удаления.")

    def save_timezone(self):
        """
        Сохраняет выбранный часовой пояс для сотрудника.
        """
        item = self.employee_list_widget.currentItem()
        if item:
            employee_name = item.text()
            timezone = float(self.timezone_combo.currentText())
            self.parent.employee_timezones[employee_name] = timezone
            QMessageBox.information(self, "Сохранение", f"Часовой пояс для {employee_name} сохранен.")
        else:
            QMessageBox.warning(self, "Предупреждение", "Выберите сотрудника для сохранения часового пояса.")

    def save_employee_timezones(self):
        """
        Сохраняет все изменения в настройках.
        """
        self.parent.save_settings()
        QMessageBox.information(self, "Сохранение", "Изменения сохранены.")

    def reset_to_default(self):
        """
        Сбрасывает часовые пояса сотрудников к значениям по умолчанию.
        """
        self.parent.employee_timezones = default_employee_timezones.copy()
        self.all_employees = list(self.parent.employee_timezones.keys())
        self.update_employee_list()
        self.timezone_combo.setEnabled(False)
        self.save_timezone_button.setEnabled(False)
        self.parent.save_settings()
        QMessageBox.information(self, "Сброс", "Часовые пояса сотрудников сброшены по умолчанию.")

    def go_back(self):
        """
        Возвращается на главную страницу.
        """
        self.parent.stacked_widget.setCurrentWidget(self.parent.main_page)
        self.parent.stacked_widget.removeWidget(self)


# Класс для выполнения длительных операций в отдельном потоке
class ProcessingThread(QThread):
    """
    Класс, представляющий поток обработки для обработки файлов в отдельном потоке.
    Этот класс предназначен для использования с PyQt для управления долго выполняющимися задачами
    и соответствующего обновления пользовательского интерфейса.
    Attributes
    ----------
    progress : pyqtSignal(int)
        Сигнал к обновлению индикатора выполнения.
    result_signal : pyqtSignal(object)
        Сигнал для передачи результата обработки.
    status_signal : pyqtSignal(str)
        Сигнал на обновление статуса обработки.
    error_signal : pyqtSignal(str)
        Сигнал для отображения ошибок.
    Methods
    -------
    __init__(self, people, zip_path, zvonki, output_excel_path)
        Метод конструктора для инициализации потока с необходимыми параметрами.
    chst_kill(self)
        Функция, позволяющая проверить, активирован ли сигнал остановки
        и потенциально вызвать исключение StopProcessing.
    run(self)
        Основной метод, содержащий основной код для выполнения в потоке.
    """
    progress = pyqtSignal(int)
    result_signal = pyqtSignal(object)
    status_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)

    # Определяем пользовательское исключение
    class StopProcessing(Exception):
        """
        Вызывается исключение для остановки обработки в приложении.
        Это исключение можно использовать, чтобы сигнализировать о том, что определенное условие или
        произошло событие, которое требует остановки текущей обработки
        немедленно. Его можно поймать и использовать для выполнения любых необходимых действий.
        деятельность по очистке или регистрации.
        """
        pass

    def __init__(
            self,
            people,
            zip_path,
            zvonki,
            output_excel_path,
            replacements,
            url_mapping,
            employee_timezones,
            company_timezone
    ):
        """
        :param people: Список или коллекция, содержащая информацию о людях.
        :param zip_path: Путь к ZIP-архиву, который необходимо обработать.
        :param zvonki: Данные или коллекции, связанные со звонками или уведомлениями.
        :param output_excel_path: Путь к файлу, в котором будет сохранен результирующий файл Excel.
        ...
        """
        super().__init__()
        self.timepath = None
        self.people = people
        self.zip_path = zip_path
        self.zvonki = zvonki
        self.output_excel_path = output_excel_path
        self.stop_signal = False
        self.excel_app = None
        self.replacements = replacements
        self.url_mapping = url_mapping
        self.employee_timezones = employee_timezones
        self.company_timezone = company_timezone

    # Функция для сокращения
    def chst_kill(self):
        """
        Проверяет, следует ли остановить процесс, и выдает исключение StopProcessing, если оно истинно.
        :return: None
        :raises StopProcessing: если условие остановки выполнено
        """
        if self.check_stop():
            logger.warning("YOU KILL ME! --<3---|-")
            raise self.StopProcessing()

    # Метод, содержащий основной код выполнения в потоке
    def run(self):
        """
        Выполняет последовательность операций, включающих создание файла Excel,
        извлечение данных и манипулирование файлами.
        :return: None
        """
        pythoncom.CoInitialize()
        excel_app = None
        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
            self.excel_app = excel_app

            # Проверка наличия и доступности файлов перед началом обработки
            self.check_file(self.people, "people", ".txt")
            self.check_file(self.zip_path, "zip-файл", ".zip")
            self.check_file(self.zvonki, "звонки", ".xlsx")
            self.check_folder(self.output_excel_path, "output Excel-файл")
            if not os.path.isdir(self.output_excel_path):
                logger.error("Путь для сохранения Excel-файла не является директорией.")
                raise ValueError("Некорректный путь для сохранения Excel-файла.")

            # =================================================================
            self.progress.emit(4)
            # 1. Парсинг данных сотрудников и отделов
            self.chst_kill()
            logger.info("\n\n#1\nЗапуск parse_departments в sotrudniki.py")
            mas_sotrudniki = parse_departments(self.people)  # sotrudniki.py
            logger.info(f"mas_sotrudniki: {mas_sotrudniki}")
            if not mas_sotrudniki:
                logger.error("mas_sotrudniki пустой или None.")
                raise ValueError("Список сотрудников пустой.")
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(7)
            self.chst_kill()
            # 2. Создание Excel-файла с отделами и сотрудниками
            logger.info("\n\n#2\nЗапуск create_department_employee_excel в excel_create.py")
            output_xlsx_path = create_department_employee_excel(mas_sotrudniki,
                                                                self.output_excel_path)  # excel_create.py
            logger.info(f"Проверка output_xlsx_path: {output_xlsx_path}")
            if output_xlsx_path is None:
                logger.error("Функция create_department_employee_excel вернула None.")
                raise ValueError("Не удалось создать Excel-файл с отделами и сотрудниками.")
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(11)
            # 3. Распаковка zip-файла
            self.chst_kill()
            logger.info("\n\n#3\nЗапуск unzip_file в zip_file.py")
            start_path = unzip_file(self.zip_path)  # zip_file.py
            logger.info(f"Проверка output_xlsx_path: {output_xlsx_path}")
            if start_path is None:
                logger.error("Функция unzip_file вернула None.")
                raise ValueError("Не удалось создать распакованный архив Стахановца.")
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(14)
            # 4. Удаление файлов PNG и других
            self.chst_kill()
            logger.info("\n\n#4\nЗапуск delete_png_files в clearPath.py")
            delete_png_files(start_path)  # clearPath.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(17)
            # 5. Удаление папок с малым количеством файлов
            self.chst_kill()
            logger.info("\n\n#5\nЗапуск delete_small_folders в clearPath.py")
            delete_small_folders(start_path)  # clearPath.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(20)
            # 6. Конвертация файлов XLS в XLSX
            self.chst_kill()
            logger.info("\n\n#6\nЗапуск convert_xls_to_xlsx в clearPath.py")
            convert_xls_to_xlsx(start_path)  # clearPath.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(23)
            # 7. Удаление файлов по определенным условиям
            self.chst_kill()
            logger.info("\n\n#7\nЗапуск delete_x_files в clearPath.py")
            delete_x_files(start_path)  # clearPath.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(26)
            # 8. Удаление папок на основе C2
            self.chst_kill()
            logger.info("\n\n#8\nЗапуск delete_folders_based_on_C2_recursive в clearPath.py")
            delete_folders_based_on_C2_recursive(start_path)  # clearPath.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(29)
            # 9. Удаление PACS файлов
            self.chst_kill()
            logger.info("\n\n#9\nЗапуск delete_pacs в clearPath.py")
            delete_pacs(start_path)  # clearPath.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(32)
            # 10. Поиск XML-файлов
            self.chst_kill()
            logger.info("\n\n#10\nЗапуск find_xml_files в search_file.py")
            found_paths = find_xml_files(start_path)  # search_file.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(35)
            # 11. Извлечение данных о программах из отчета
            self.chst_kill()
            logger.info("\n\n#11\nЗапуск extract_report_data_prog в siteNprog_normolize.py")
            output_file_path_prog = extract_report_data_prog(found_paths)  # siteNprog_normolize.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(39)
            # 12. Извлечение данных о сайтах из отчета
            self.chst_kill()
            logger.info("\n\n#12\nЗапуск extract_report_data_site в siteNprog_normolize.py")
            output_file_path_site = extract_report_data_site(found_paths)  # siteNprog_normolize.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(42)
            # 13. Загрузить данные о программах из XML файла
            self.chst_kill()
            logger.info("\n\n#13\nЗапуск load_program_data в siteNprog_toexcel.py")
            program_data = load_program_data(output_file_path_prog)  # siteNprog_toexcel.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(45)
            # 14. Загрузить данные о сайтах из XML файла
            self.chst_kill()
            logger.info("\n\n#14\nЗапуск load_site_data в siteNprog_toexcel.py")
            site_data = load_site_data(output_file_path_site)  # siteNprog_toexcel.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(48)
            # 15. Обновление Excel-файла с данными сотрудников
            self.chst_kill()
            logger.info("\n\n#15\nЗапуск update_employee_sheets в siteNprog_toexcel.py")
            update_employee_sheets(output_xlsx_path, program_data, site_data)  # siteNprog_toexcel.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(51)
            # 16. Обработка и сохранение данных Excel
            self.chst_kill()
            logger.info("\n\n#16\nЗапуск process_excel в inExcel_site.py")
            output_file_excelSite = os.path.join(self.output_excel_path,
                                                 'Отчеты_отделов_и_сотрудников_обновленный.xlsx')
            process_excel(output_xlsx_path, self.url_mapping,
                          output_file_excelSite)  # Передаем желаемый путь для сохранения
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(55)
            # 17. Создание копии файла со звонками
            self.chst_kill()
            logger.info("\n\n#17\nЗапуск create_file_copy в zvonki_normolize.py")
            copy_file_path_zv = create_file_copy(self.zvonki)
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(58)
            # 18. Обработка и сохранение данных о звонках
            self.chst_kill()
            logger.info("\n\n#18\nЗапуск process_and_save_calls_data в zvonki_normolize.py")
            process_and_save_calls_data(copy_file_path_zv, self.replacements)  # zvonki_normolize.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(61)
            # 19. Обновление отчета по звонкам
            self.chst_kill()
            logger.info("\n\n#19\nЗапуск zvonkiExcel в zvonki_toexcel.py")
            zvonkiExcel(output_file_excelSite, copy_file_path_zv)  # zvonki_toexcel.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(64)
            # 20. Чистка файлов Стахановца и переименование
            self.chst_kill()
            logger.info("\n\n#20\nЗапуск process_folders в stahName.py")
            process_folders(start_path, self.replacements)  # stahName.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(69)
            # 21. Удаление ненужных файлов/папок
            self.chst_kill()
            logger.info("\n\n#21\nЗапуск rename_folders_from_excel_cell в stahName.py")
            rename_folders_from_excel_cell(start_path)  # stahName.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(74)
            # 22. Объединение всех сотрудников в один список
            self.chst_kill()
            logger.info("\n\n#22\nЗапуск get_all_employees в infoWork_stah.py")
            all_employees = get_all_employees(mas_sotrudniki)  # infoWork_stah.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(77)
            # 23. Сканирование всех подкаталогов
            self.chst_kill()
            logger.info("\n\n#23\nЗапуск scan_folders в infoWork_stah.py")
            scan_folders(start_path, all_employees)  # infoWork_stah.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(80)
            # 24. Вставка значений из Стахановца в Excel
            self.chst_kill()
            logger.info("\n\n#24\nЗапуск update_excel_with_employee_data в infoStah_toexcel.py")
            update_excel_with_employee_data(info_work_stah, output_file_excelSite)  # infoStah_toexcel.py
            print_pretty_info_work_stah1(info_work_stah)
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(83)
            # 25. Преобразование файла битрикс xls в xlsx
            self.chst_kill()
            logger.info("\n\n#25\nЗапуск convert_html_to_xlsx в bitrix_normolize.py")
            file_path_bit = convert_html_to_xlsx(bitrix_path)  # bitrix_normolize.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(86)
            # 26. Форматирование файла Битрикса
            self.chst_kill()
            logger.info("\n\n#26\nЗапуск replace_values_in_xlsx в bitrix_normolize.py")
            replace_values_in_xlsx(file_path_bit, self.replacements)  # bitrix_normolize.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(89)
            # 27. Форматирование файла отчета xlsx
            self.chst_kill()
            logger.info("\n\n#27\nЗапуск format_excel_file в format.py")
            format_excel_file(output_file_excelSite)  # format.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(95)
            # 28. Переименовывание основных файлов Стахановца
            self.chst_kill()
            logger.info("\n\n#28\nЗапуск frename_xlsx_files в dost_file.py")
            frename_xlsx_files(start_path)  # dost_file.py
            self.chst_kill()
            # =================================================================

            # =================================================================
            self.progress.emit(99)
            # 29. Перенос временных файлов
            self.chst_kill()
            logger.info("\n\n#29\nЗапуск fileto_log в dost_file.py")
            timefiles = [copy_file_path_zv, start_path, output_xlsx_path]
            timepath = fileto_log(timefiles, self.output_excel_path)  # dost_file.py
            self.timepath = timepath
            self.chst_kill()
            # =================================================================

            self.progress.emit(100)
            logger.info("Все функции успешно выполнены!")
            self.result_signal.emit(output_file_excelSite)


        except self.StopProcessing:
            # Обрабатываем остановку процесса
            logger.warning("Процесс был остановлен пользователем.")
            self.status_signal.emit("Процесс остановлен.")
            self.progress.emit(0)
            return
        except Exception as e:
            logger.exception("Произошла ошибка в ProcessingThread")
            error_message = f"Произошла ошибка: {str(e)}"
            self.error_signal.emit(error_message)
            gen_py_path = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'Temp', 'gen_py')
            if os.path.exists(gen_py_path):
                shutil.rmtree(gen_py_path)
        finally:
            if excel_app is not None:
                try:
                    logger.info("Закрытие Excel приложения")
                    excel_app.Quit()
                except Exception as e:
                    logger.exception(f"Ошибка при завершении Excel: {str(e)}")
                finally:
                    del excel_app
            pythoncom.CoUninitialize()

    # Проверяет существование, доступность и формат файла
    def check_file(self, file_path, file_description, expected_extension):
        """
        :param file_path: Путь к проверяемому файлу.
        :param file_description: Описание проверяемого файла.
        :param expected_extension: Ожидаемое расширение файла.
        :return: None
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"{file_description} не найден: {file_path}")
        if not os.access(file_path, os.R_OK):
            raise PermissionError(f"Нет прав на чтение {file_description}: {file_path}")
        if not file_path.endswith(expected_extension):
            raise ValueError(
                f"Неправильный формат для {file_description}. Ожидается {expected_extension}, получено {os.path.splitext(file_path)[1]}")
        logger.info(f"{file_description} успешно проверен: {file_path}")

    # Проверяет существование и формат папки
    def check_folder(self, folder_path, folder_description):
        """
        :param folder_path: Путь к папке, которую необходимо проверить.
        :param folder_description: Описание папки, используемое в сообщениях журнала и обработке ошибок.
        :return: None
        """
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"{folder_description} не найден: {folder_path}")
        if not os.path.isdir(folder_path):
            raise ValueError(f"{folder_description} не является папкой: {folder_path}")
        logger.info(f"{folder_description} успешно проверен: {folder_path}")

    # Проверяет, был ли отправлен сигнал остановки
    def check_stop(self):
        """
        :return: Возвращает True, если обнаружен stop_signal и процесс остановлен, в противном случае — False.
        """
        if self.stop_signal:
            logger.debug("stop_signal обнаружен как True")
            self.status_signal.emit("Процесс остановлен.")
            self.progress.emit(0)
            return True
        return False

    # Останавливаем выполнение потока
    def stop(self):
        """
        Устанавливает атрибут stop_signal в значение True, указывая, что процесс следует остановить.
        Этот метод обычно вызывается для корректного завершения запущенного процесса.
        :return: None
        """
        logger.debug("Метод stop() вызван, устанавливаем stop_signal в True")
        self.stop_signal = True


# Класс окна консоли для отображения логов
class ConsoleWindow(QWidget):
    """
    Подкласс QWidget, создающий окно консоли для отображения текстового вывода.
    Этот класс используется для создания консольного окна виджета с помощью QTextEdit.
    для вывода. Заголовок окна установлен на «Консоль» (Консоль на кириллице) и
    геометрия установлена на (150, 150, 800, 400). Он использует «DejaVu Sans Mono».
    шрифт для поддержки символов кириллицы и применения темной темы с помощью QSS.
    Methods:
        __init__(): Инициализирует окно консоли с заголовком, геометрией,
                    настройки шрифта, таблица стилей и макет. Также соединяет
                    обработчик буфера для виджета консоли и отображает
                    любые накопленные сообщения журнала.
    """

    def __init__(self):
        """
        Инициализирует окно консоли с определенными настройками.
        Окно настроено на отображение вывода консоли шрифтом, поддерживающим кириллицу.
        Он реализует следующие функции:
        - Устанавливает заголовок окна «Консоль» и размеры 800x400 пикселей.
        — Настраивает виджет QTextEdit для отображения вывода консоли,
          инициализируя его как доступный только для чтения и применяя шрифт «DejaVu Sans Mono».
        - Применяет таблицу стилей к QWidget и QTextEdit для установки цветов фона и текста,
          семейства шрифтов, размера шрифта и свойств границ.
        — Создает и устанавливает QVBoxLayout, добавляя к нему QTextEdit.
        — Привязывает обработчик буфера к текстовому виджету вывода консоли,
          позволяя отображать накопленные сообщения журнала.
        — Перебирает сообщения в буфере и добавляет их к выводу консоли, гарантируя,
          что для полосы прокрутки установлено максимальное значение.
        """
        super().__init__()
        self.setWindowTitle("Консоль")
        self.setGeometry(150, 150, 800, 400)

        # Устанавливаем шрифт, поддерживающий кириллицу
        font = QFont("DejaVu Sans Mono", 10)
        self.console_output = QTextEdit(self)
        self.console_output.setReadOnly(True)
        self.console_output.setFont(font)

        self.setStyleSheet("""
            QWidget {
                background-color: #2b2b2b;
                color: #f0f0f0;
            }
            QTextEdit {
                background-color: #2b2b2b;
                color: #f0f0f0;
                font-family: 'DejaVu Sans Mono', monospace;
                font-size: 12px;
                border: none;
            }
        """)

        layout = QVBoxLayout()
        layout.addWidget(self.console_output)
        self.setLayout(layout)

        # Привязываем обработчик к виджету консоли
        buffer_handler.text_widget = self.console_output

        # Отображаем накопленные записи журнала
        for msg in buffer_handler.buffer:
            self.console_output.append(msg)
        self.console_output.verticalScrollBar().setValue(
            self.console_output.verticalScrollBar().maximum()
        )


# Класс основной страницы (главная)
class MainPage(QWidget):
    """
    Главная страница
    Класс, производный от QWidget, который представляет главную страницу приложения с выбором файла.
    кнопки действий и индикаторы прогресса.
    Methods
    -------
    __init__(self, parent)
        Инициализирует экземпляр MainPage, настраивая элементы пользовательского интерфейса.
    init_ui(self)
        Создает и упорядочивает компоненты пользовательского интерфейса,
        такие как кнопки, метки и индикаторы выполнения, а также
        связывает кнопки с соответствующими функциями.
    check_excel_processes(self)
        Проверяет наличие всех запущенных процессов Excel и предлагает пользователю завершить их перед
        запуск нового процесса.
    terminate_excel_processes(self)
        Завершает все запущенные процессы Excel на компьютере.
    """

    def __init__(self, parent):
        """
        :param parent: Родительский виджет для этого компонента пользовательского интерфейса.
                       Обычно это главное окно или другое содержащее виджет.
        """
        super().__init__()
        self.button_open_report = None
        self.button_select_zvonki = None
        self.file_zip_label = None
        self.progress_bar = None
        self.result_label = None
        self.button_run_process = None
        self.button_stop = None
        self.button_reset = None
        self.folder_output_label = None
        self.file_zvonki_label = None
        self.file_people_label = None
        self.button_select_output = None
        self.button_select_zip = None
        self.button_select_people = None
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        """
        Инициализирует компоненты пользовательского интерфейса (UI) приложения.
        :return: None
        """
        # Виджеты выбора файлов
        self.button_select_people = QPushButton("Файл с сотрудниками", self)
        self.button_select_zip = QPushButton("ZIP файл", self)
        self.button_select_zvonki = QPushButton("Звонки АТС", self)
        self.button_select_output = QPushButton("Папка для Excel", self)

        # Поля для отображения выбранных файлов
        self.file_people_label = QLineEdit(self)
        self.file_zip_label = QLineEdit(self)
        self.file_zvonki_label = QLineEdit(self)
        self.folder_output_label = QLineEdit(self)
        for label in [self.file_people_label, self.file_zip_label, self.file_zvonki_label, self.folder_output_label]:
            label.setReadOnly(True)

        if self.parent.output_excel_path:
            folder_name = os.path.basename(self.parent.output_excel_path)
            self.folder_output_label.setText(folder_name)

        # Кнопки управления
        self.button_reset = QPushButton("Сброс", self)
        self.button_reset.setFixedSize(80, 35)
        self.button_stop = QPushButton("Стоп", self)
        self.button_run_process = QPushButton("Запустить обработку", self)
        self.button_run_process.setEnabled(False)
        self.button_stop.setEnabled(False)

        # Определение прогресс-бара
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Компоновка выбора файлов
        file_selection_layout = QGridLayout()
        file_selection_layout.addWidget(self.button_select_people, 0, 0)
        file_selection_layout.addWidget(self.file_people_label, 0, 1)
        file_selection_layout.addWidget(self.button_select_zip, 0, 2)
        file_selection_layout.addWidget(self.file_zip_label, 0, 3)
        file_selection_layout.addWidget(self.button_select_zvonki, 1, 0)
        file_selection_layout.addWidget(self.file_zvonki_label, 1, 1)
        file_selection_layout.addWidget(self.button_select_output, 2, 0)
        file_selection_layout.addWidget(self.folder_output_label, 2, 1, 1, 3)

        # Подключение кнопок к функциям
        self.button_select_people.clicked.connect(self.parent.select_people_file)
        self.button_select_zip.clicked.connect(self.parent.select_zip_file)
        self.button_select_zvonki.clicked.connect(self.parent.select_zvonki_file)
        self.button_select_output.clicked.connect(self.parent.select_output_folder)
        self.button_run_process.clicked.connect(self.check_excel_processes)
        self.button_stop.clicked.connect(self.parent.stop_thread)
        self.button_reset.clicked.connect(self.parent.reset_all)

        # Компоновка кнопок действий
        action_buttons_layout = QVBoxLayout()
        action_buttons_layout.addWidget(self.button_reset, alignment=Qt.AlignmentFlag.AlignLeft)
        action_buttons_layout.addWidget(self.progress_bar)
        action_buttons_layout.addWidget(self.button_stop)
        action_buttons_layout.addWidget(self.button_run_process)

        self.result_label = QLabel("", self)
        action_buttons_layout.addWidget(self.result_label)

        # Основной макет страницы
        main_layout = QGridLayout()
        main_layout.addLayout(file_selection_layout, 0, 0)
        main_layout.addLayout(action_buttons_layout, 1, 0)

        self.button_open_report = QPushButton("Открыть готовый отчет", self)
        self.button_open_report.setEnabled(False)

        # Подключение кнопки к методу в MainWindow
        self.button_open_report.clicked.connect(self.parent.open_ready_report)

        # Добавьте кнопку в макет
        action_buttons_layout.addWidget(self.button_open_report)

        self.setLayout(main_layout)

    # Проверяем наличие запущенных процессов Excel и предлагаем действия пользователю
    def check_excel_processes(self):
        """
        Проверяет наличие запущенных процессов Excel и отображает диалоговое окно
        с предупреждением, если таковое обнаружено.
        Если обнаружены процессы Excel, откроется диалоговое окно с опциями завершения.
        все процессы Excel и продолжить дальнейшие операции или отменить
        показана операция. Если пользователь решает завершить процессы,
        он вызовет метод `terminate_excel_processes`, чтобы остановить эти
        процессы и впоследствии вызвать метод `run_all_functions` из
        родительский объект.
        Если процессы Excel не найдены, он немедленно вызывает `run_all_functions`
        метод из родительского объекта.
        :return: None
        """
        excel_processes = [proc for proc in psutil.process_iter(['name']) if
                           proc.info['name'] and 'excel' in proc.info['name'].lower()]

        if excel_processes:
            message_box = QMessageBox(self)
            message_box.setWindowTitle("Запущенные процессы Excel")
            message_box.setText("Перед запуском советуем закрыть все Excel файлы.\nВыберите действие:")
            message_box.setIcon(QMessageBox.Icon.Warning)

            # Добавляем стандартные кнопки
            terminate_button = message_box.addButton("Завершить все процессы и начать обработку",
                                                     QMessageBox.ButtonRole.AcceptRole)
            cancel_button = message_box.addButton("Отмена", QMessageBox.ButtonRole.RejectRole)

            # Показываем окно и обрабатываем выбор
            message_box.exec()

            if message_box.clickedButton() == terminate_button:
                self.terminate_excel_processes()
                self.parent.run_all_functions()
            else:
                return
        else:
            self.parent.run_all_functions()

    # Завершаем все процессы Excel на компьютере
    def terminate_excel_processes(self):
        """
        Завершает все запущенные процессы Excel.
        :return: None
        """
        for proc in psutil.process_iter(['name']):
            try:
                process_name = proc.info['name']
                if process_name and 'excel' in process_name.lower():
                    logger.info(f"Найден процесс Excel: PID {proc.pid}, имя '{process_name}'")
                    proc.terminate()
                    proc.wait()
                    logger.info(f"Процесс Excel с PID {proc.pid} завершен.")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
            except Exception as e:
                logger.error(f"Ошибка при завершении процесса Excel с PID {proc.pid}: {str(e)}")
            except psutil.AccessDenied:
                logger.error(f"Нет прав на завершение процесса Excel с PID {proc.pid}")
                QMessageBox.warning(self, "Недостаточно прав",
                                    f"Нет прав на завершение процесса Excel с PID {proc.pid}. Запустите приложение от имени администратора.")

    def set_file_selection_buttons_enabled(self, enabled):
        self.button_select_people.setEnabled(enabled)
        self.button_select_zip.setEnabled(enabled)
        self.button_select_zvonki.setEnabled(enabled)
        self.button_select_output.setEnabled(enabled)


# Класс окна с примерами
class ExamplesPage(QWidget):
    """
    Класс, реализующий QWidget для отображения страницы примеров с прокручиваемой областью,
    содержащей необходимые файлы и дополнительную информацию.
    Methods
    -------
    __init__():
        Инициализирует экземпляр SamplesPage и настраивает пользовательский интерфейс.
    init_ui():
        Настраивает элементы пользовательского интерфейса, включая область прокрутки, метки
        и текстовые поля для отображения списка необходимых файлов, дополнительной информации и примеров файлов.
    """

    def __init__(self):
        """
        Инициализирует экземпляр класса.
        Этот метод отвечает за инициализацию пользовательского интерфейса (UI) класса путем вызова метода init_ui.
        Сначала он вызывает конструктор родительского класса, используя `super().__init__()`.
        Methods:
            init_ui: Настраивает компоненты пользовательского интерфейса.
        """
        super().__init__()
        self.init_ui()

    def init_ui(self):
        """
        Инициализирует пользовательский интерфейс для отображения списка необходимых файлов и примеров.
        :return: None
        """
        # Создаем область прокрутки
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        # Создаем виджет для содержимого
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)

        # Заголовок
        title_label = QLabel("Список необходимых файлов:")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title_label)

        # Список необходимых файлов
        files_list = QLabel(
            "1. Файл с сотрудниками (текстовый файл .txt)\n"
            "2. ZIP-файл с данными Стахановца (.zip)\n"
            "3. Файл звонков из АТС (.xlsx)\n"
            "4. Папка для сохранения отчета (Excel)"
        )
        layout.addWidget(files_list)

        # Добавляем дополнительную информацию
        additional_label = QLabel("\nДополнительная информация")
        additional_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        info_files = QLabel(
            "1. Пример файла с сотрудниками вы можете найти ниже.\n\n"
            "2. Данные из Стахановца, вам понадобится:\n"
            "   Глобальный отчет, выбираете нужный диапазон дат. Все отчеты должны быть По каждому.\n"
            "   Отчеты, которые вам понадобятся:\n"
            "       1. Табель УРВ\n"
            "       2. Сайты\n"
            "       3. Программы\n"
            "       4. Пользовательское время\n"
            "   Вы должны создать отчет и выгрузить его в .zip формате.\n\n"
            "3. Файл звонков из АТС:\n"
            "   Вы должны выгрузить отчет Excel за нужный вам диапазон, без каких либо фильтров более.\n\n"
            "4. Папка для сохранения отчета (Excel):\n"
            "   Выберите любую удобную папку, в которую вы хотите сохранить готовый Excel отчет.\n"
        )

        layout.addWidget(additional_label)
        layout.addWidget(info_files)

        # Пример файла с сотрудниками
        example_label = QLabel("\nПример файла с сотрудниками:")
        example_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(example_label)

        # Текст с примером сотрудников
        example_text = (
            "Отдел продаж (ОП):\n"
            "Менеджер 1\n"
            "Менеджер 2\n"
            "Менеджер 3\n"
            "Менеджер 4\n"
            "Менеджер 5\n"
            "Менеджер 6\n\n"
            "Отдел продаж 2 (ОП2):\n"
            "Менеджер 1\n"
            "Менеджер 2\n"
            "Менеджер 3\n"
            "Менеджер 4\n"
            "Менеджер 5\n"
            "Менеджер 6\n\n"
        )

        # Используем QTextEdit для отображения текста с темно-серым фоном
        example_text_edit = QTextEdit()
        example_text_edit.setReadOnly(True)
        example_text_edit.setPlainText(example_text)
        example_text_edit.setStyleSheet("""
            QTextEdit {
                background-color: #2b2b2b; /* Темно-серый фон */
                color: #f0f0f0; /* Цвет текста */
                border: 1px solid #cccccc;
                border-radius: 10px;
                padding: 10px;
            }
        """)
        example_text_edit.setFixedHeight(270)  # Устанавливаем фиксированную высоту в 200 пикселей
        layout.addWidget(example_text_edit)

        additional_label = QLabel("\n\nСоответствия")
        additional_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        info_files = QLabel(
            "1. Соответствия сотрудников:\n"
            "   В левой части вы прописываете имя сотрудника так, как оно указано в файле .TXT с именами сотрудников.\n"
            "   В правой части, нажав на имя сотрудника, вы прописываете все имена, которые используются в Стахановце, Битриксе или АТС.\n"
            "   Для чего это нужно?:\n"
            "       Эти соответствия необходимы, так как имена сотрудников в разных отчетах и системах пишутся по-разному.\n"
            "       Разберем пример: У нас есть Василий Васечкин, который работает в Отделе продаж 1, у него есть:\n"
            "       2 аккаунта Битрикс (Основной + пробивной)\n"
            "       2 аккаунта в АТС (Основной + пробивной)\n"
            "       1 аккаунт Стахановец\n"
            "       И везде Вася назван по-разному, поэтому и нужны соответствия, которые будут выглядеть следующим образом:\n\n"
            "                                          _______________Васечкин ОП *битрикс аккаунт основной*\n"
            "                                         /_______________Васечкин ОП ПРОБИВ *битрикс аккаунт пробивной*\n"
            "       Василий Васечкин_/_______________Василий Васечкин *АТС аккаунт основной*\n"
            "                                        \_______________Василий Васечкин ПРОБИВ *АТС аккаунт пробивной*\n"
            "                                         \_______________Василий Васечкин ОП *Стахановец*\n\n"
            "2. Соответствия сайтов:\n"
            "   В левой части вы прописываете название сайта, которое хотите видеть в отчете .xlsx.\n"
            "   В правой части, нажав на сайт, вы указываете все его альтернативные названия, которые используются для замены.\n"
            "   Для чего это нужно?:\n"
            "       Эти соответствия нужны для того, чтобы быстрее понимать, какие сайты посещал сотрудник, и для экономии места в отчете.\n\n"
            "3. Часовые пояса сотрудников:\n"
            "   В левой части вы выбираете сотрудника, для которого хотите настроить часовой пояс.\n"
            "   Часовой пояс прописывается с указанием разницы по сравнению с Гринвичем: МСК - 3, Иркутск - 8.\n\n"
            "4. Часовой пояс компании:\n"
            "   Часовой пояс компании задается для расчета данных и также указывается относительно Гринвича.\n\n"
            "Все примеры можно увидеть в уже готовых списках.\n"
            "При любом изменении не забудьте сохранить данные, иначе они будут утеряны.\n"
            "Кнопка 'Вернуть по умолчанию' предназначена для сброса всех ваших изменений.\n"
        )

        layout.addWidget(additional_label)
        layout.addWidget(info_files)

        # Устанавливаем компоновку для виджета содержимого
        content_widget.setLayout(layout)

        # Устанавливаем виджет содержимого в область прокрутки
        scroll_area.setWidget(content_widget)

        # Основная компоновка страницы ExamplesPage
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(scroll_area)
        self.setLayout(main_layout)


# Класс страницы для редактирования файла сотрудников
class EditEmployeesPage(QWidget):
    """
    Класс EditEmployeesPage для редактирования файла сотрудника.
    Класс предоставляет текстовый редактор для изменения файла
    и кнопки для сохранения изменений или возврата на главную страницу.
    :param parent: The parent widget.
    :type parent: QWidget
    :param file_path: Путь к файлу с данными о сотрудниках.
    :type file_path: str
    """

    def __init__(self, parent, people_file):
        """
        :param parent: Родительский виджет или окно для этого экземпляра.
        :param file_path: Путь к файлу, связанному с этим экземпляром.
        """
        super().__init__()
        self.save_button = None
        self.back_button = None
        self.text_edit = None
        self.parent = parent
        self.people_file = people_file
        self.init_ui()

    def init_ui(self):
        """
        Инициализирует пользовательский интерфейс для редактирования файла сотрудника.
        Этот метод устанавливает вертикальный макет с меткой заголовка, QTextEdit для
        редактирование содержимого файла и кнопки «Сохранить» и «Назад». Он пытается загрузить
        содержимое файла в текстовый редактор и подключает кнопки к их
        соответствующие методы.
        :return: None
        """
        layout = QVBoxLayout()

        # Добавляем заголовок
        title_label = QLabel("Редактирование файла сотрудников")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title_label)

        # Создаем текстовый редактор
        self.text_edit = QTextEdit(self)
        layout.addWidget(self.text_edit)

        # Загружаем содержимое файла
        try:
            with open(self.people_file, 'r', encoding='utf-8') as file:
                content = file.read()
                self.text_edit.setPlainText(content)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл: {str(e)}")

        # Кнопки "Сохранить" и "Назад"
        button_layout = QHBoxLayout()
        self.save_button = QPushButton("Сохранить", self)
        self.back_button = QPushButton("Назад", self)
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.back_button)

        #layout.addWidget(self.save_button)

        layout.addLayout(button_layout)

        # Подключаем кнопки к методам
        self.save_button.clicked.connect(self.save_file)
        self.back_button.clicked.connect(self.go_back)

        self.setLayout(layout)

    def load_employees(self):
        try:
            with open(self.people_file, 'r', encoding='utf-8') as file:
                content = file.read()
                self.text_edit.setPlainText(content)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл сотрудников: {str(e)}")

    # Сохраняет содержимое текстового редактора в файл
    def save_file(self):
        """
        Сохраните содержимое text_edit в файл, указанный в file_path.
        :return: None
        """
        try:
            content = self.text_edit.toPlainText()
            with open(self.people_file, 'w', encoding='utf-8') as file:
                file.write(content)
            QMessageBox.information(self, "Сохранение", "Файл успешно сохранен.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {str(e)}")

    # Возвращается на главную страницу
    def go_back(self):
        """
        Вернитесь на главную страницу, установив текущий виджет в составном виджете на главную страницу.
        Удаляет страницу редактора из сложенного виджета.
        :return: None
        """
        self.parent.stacked_widget.setCurrentWidget(self.parent.main_page)
        # Удаляем страницу редактора из стека
        self.parent.stacked_widget.removeWidget(self)

    def set_disabled(self, disabled):
        self.text_edit.setReadOnly(disabled)
        self.save_button.setDisabled(disabled)
        # Кнопку "Назад" можно оставить активной


# Класс страницы натсроек
class SettingsPage(QWidget):
    """
    страница настроек класса (QWidget):
    SettingsPage — это QWidget, который позволяет пользователям настраивать параметры приложения,
    такие как выбор темы и изменение пути сохранения отчетов.
    Methods:
        __init__(self, parent):
            Инициализирует виджет SettingsPage с необходимыми компонентами и родительским виджетом.
        init_ui(self):
            Настраивает элементы пользовательского интерфейса для страницы SettingsPage,
            включая метки, поля ввода, кнопки и переключатели выбора темы.
        change_save_path(self):
            Открывает диалоговое окно файла для выбора нового каталога для сохранения отчетов,
            обновляет путь сохранения и обновляет настройки родительского виджета.
        change_theme(self):
            Изменяет тему приложения в зависимости от выбора пользователя
            (темная или светлая тема) и обновляет настройки родительского виджета.
    """

    def __init__(self, parent):
        """
        :param parent:  Родительский виджет для этого виджета настроек.
        Этот параметр используется для установки отношений родитель-потомок в иерархии виджетов Qt.
        """
        super().__init__()
        self.save_company_timezone_button = None
        self.company_timezone_edit = None
        self.light_theme_radio = None
        self.dark_theme_radio = None
        self.theme_group = None
        self.change_path_button = None
        self.save_path_edit = None
        self.save_path_label = None
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        """
        Инициализирует и настраивает пользовательский интерфейс.
        :return: None
        """
        layout = QVBoxLayout()

        # Метка и поле для редактирования пути сохранения
        self.save_path_label = QLabel("Путь для сохранения отчетов:", self)
        self.save_path_edit = QLineEdit(self)
        self.save_path_edit.setReadOnly(True)
        if self.parent.output_excel_path:
            self.save_path_edit.setText(self.parent.output_excel_path)

        # Кнопка для изменения пути
        self.change_path_button = QPushButton("Изменить путь", self)
        self.change_path_button.clicked.connect(self.change_save_path)

        # Добавляем виджеты в макет
        layout.addWidget(self.save_path_label)
        layout.addWidget(self.save_path_edit)
        layout.addWidget(self.change_path_button)

        # Добавляем раздел для выбора темы
        theme_label = QLabel("Тема приложения:", self)
        layout.addWidget(theme_label)
        self.theme_group = QButtonGroup(self)
        self.dark_theme_radio = QRadioButton("Тёмная тема", self)
        self.light_theme_radio = QRadioButton("Светлая тема", self)
        self.theme_group.addButton(self.dark_theme_radio)
        self.theme_group.addButton(self.light_theme_radio)

        # Устанавливаем выбранную тему по настройкам
        if self.parent.theme == 'dark':
            self.dark_theme_radio.setChecked(True)
        else:
            self.light_theme_radio.setChecked(True)

        # Подключаем изменение темы
        self.dark_theme_radio.toggled.connect(self.change_theme)
        self.light_theme_radio.toggled.connect(self.change_theme)

        layout.addWidget(self.dark_theme_radio)
        layout.addWidget(self.light_theme_radio)

        # Добавляем раздел для редактирования company_timezone
        company_timezone_label = QLabel("\nЧасовой пояс компании:", self)
        layout.addWidget(company_timezone_label)

        self.company_timezone_edit = QLineEdit(self)
        self.company_timezone_edit.setText(str(self.parent.company_timezone))
        layout.addWidget(self.company_timezone_edit)

        # Добавляем кнопку для сохранения
        self.save_company_timezone_button = QPushButton("Сохранить часовой пояс", self)
        self.save_company_timezone_button.clicked.connect(self.save_company_timezone)
        layout.addWidget(self.save_company_timezone_button)

        layout.addStretch()

        self.setLayout(layout)

    # Сохранияем выбранный путь для excel файлов
    def change_save_path(self):
        """
        Предлагает пользователю выбрать путь к каталогу и
        обновляет состояние приложения в соответствии с выбранным путем.
        Выбранный путь устанавливается в качестве нового места сохранения отчетов,
        и различные элементы пользовательского интерфейса обновляются с учетом этого изменения.
        :return: None
        """
        folder = QFileDialog.getExistingDirectory(self, "Выбрать папку для сохранения отчетов")
        if folder:
            self.parent.output_excel_path = folder
            self.save_path_edit.setText(folder)
            self.parent.save_settings()
            folder_name = os.path.basename(folder)
            self.parent.main_page.folder_output_label.setText(folder_name)

    # Выбираем тему
    def change_theme(self):
        """
        Устанавливает тему приложения в зависимости от состояния переключателя темной темы.
        Переключает тему на «темную», если установлен переключатель «Темная тема»,
        в противном случае переключает тему на «светлую». Затем он применяет новую таблицу стилей
        и сохраняет текущие настройки.
        :return: None
        """
        if self.dark_theme_radio.isChecked():
            self.parent.theme = 'dark'
        else:
            self.parent.theme = 'light'
        self.parent.apply_stylesheet()
        self.parent.save_settings()

    def save_company_timezone(self):
        """
        Сохраняет настройку часового пояса компании из поля ввода в настройки родительского объекта.
        Эта функция пытается прочитать целочисленное значение из поля редактирования часового пояса компании.
        сохраните его в атрибуте Company_timezone родительского объекта, а затем сохраните настройки.
        Если входные данные не являются допустимым целым числом, пользователю будет показано предупреждающее сообщение.
        Parameters
        ----------
        self : object
            Экземпляр класса, содержащий этот метод.
        Raises
        ------
        ValueError
            если введенный часовой пояс не является допустимым целым числом.
        """
        try:
            tz = int(self.company_timezone_edit.text())
            self.parent.company_timezone = tz
            self.parent.save_settings()
            QMessageBox.information(self, "Сохранение", "Часовой пояс компании сохранен.")
        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Введите корректное целое число для часового пояса.")

    def set_change_path_button_enabled(self, enabled):
        self.change_path_button.setEnabled(enabled)

    def update_theme_selection(self):
        if self.parent.theme == 'dark':
            self.dark_theme_radio.setChecked(True)
        else:
            self.light_theme_radio.setChecked(True)


# Класс прошлых отчетов
class PastReportsPage(QWidget):
    """
    Класс PastReportsPage — это QWidget, который отображает список прошлых отчетов
    и позволяет пользователю открыть выбранный отчет или вернуться на главную страницу.
    :param parent: Родительский виджет, обычно главное окно приложения.
    """

    def __init__(self, parent):
        """
        :param parent:  Родительский виджет или компонент, частью которого является этот класс.
                        Он используется для инициализации родительского атрибута и важен
                        для иерархической структуры пользовательского интерфейса.
        """
        super().__init__()
        self.back_button = None
        self.reports_list = None
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        """
        Инициализирует пользовательский интерфейс для отображения прошлых отчетов.
        :return: None
        """
        layout = QVBoxLayout()

        # Заголовок
        title_label = QLabel("Прошлые отчеты")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title_label)

        # Список отчетов
        self.reports_list = QListWidget()

        # Загрузка отчетов
        self.load_reports()

        # Двойной клик для открытия отчета
        self.reports_list.itemDoubleClicked.connect(self.open_report)

        layout.addWidget(self.reports_list)

        # Кнопка "Назад"
        self.back_button = QPushButton("Назад", self)
        self.back_button.clicked.connect(self.go_back)
        layout.addWidget(self.back_button)

        self.setLayout(layout)

    # Загрузка прошлых отчетов
    def load_reports(self):
        """
        Загружает отчеты из указанного пути вывода и заполняет список отчетов.
        :return: None
        """
        self.reports_list.clear()
        output_path = self.parent.output_excel_path
        if output_path and os.path.exists(output_path):
            reports = [f for f in os.listdir(output_path) if f.endswith('.xlsx')]
            for report in reports:
                item = QListWidgetItem(report)
                self.reports_list.addItem(item)
        else:
            QMessageBox.warning(self, "Предупреждение", "Путь для сохранения отчетов не задан или недоступен.")

    # Открытие отчета
    def open_report(self, item):
        """
        :param item: Элемент, представляющий открываемый отчет.
        :return: None
        """
        report_name = item.text()
        report_path = os.path.join(self.parent.output_excel_path, report_name)
        if os.path.exists(report_path):
            try:
                os.startfile(report_path)
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось открыть файл: {str(e)}")
        else:
            QMessageBox.warning(self, "Предупреждение", "Файл отчета не найден.")

    # Возвращение назад
    def go_back(self):
        """
        Устанавливает текущий виджет составного виджета на главную страницу и удаляет текущий виджет.
        :return: None
        """
        self.parent.stacked_widget.setCurrentWidget(self.parent.main_page)
        self.parent.stacked_widget.removeWidget(self)


# Основное окно приложения
class MainWindow(QMainWindow):
    """
    Класс главного окна приложения обработки данных сотрудников
    Stylesheets
    -----------
    dark_theme_stylesheet
        Таблица стилей для темной темы интерфейса приложения.
    light_theme_stylesheet
        Таблица стилей для светлой темы интерфейса приложения.
    Methods
    -------
    __init__()
        Инициализирует класс MainWindow и настраивает главное окно приложения.
    init_ui()
        Настраивает пользовательский интерфейс, включая строку меню и боковую панель с кнопками навигации.
    keyPressEvent(event)
        Обрабатывает события нажатия клавиш, в частности определяя нажатие обеих клавиш Shift.
    keyReleaseEvent(event)
        Обрабатывает события отпускания клавиш для управления состоянием клавиш Shift.
    show_temp_files_button()
        Отображает кнопку «Временные файлы» при нажатии обеих клавиш Shift.
    """
    # Темная тема приложения
    dark_theme_stylesheet = """
    QWidget {
        background-color: #1e1e2e;
        color: #f0f0f0;
        font-family: Arial, sans-serif;
    }
    QPushButton {
        background-color: #2b2b3d;
        color: #ffffff;
        border-radius: 8px;
        padding: 10px;
        font-size: 14px;
    }
    QPushButton:hover {
        background-color: #3d3d5c;
    }
    QPushButton:pressed {
        background-color: #454563;
    }
    QLineEdit {
        background-color: #2b2b3d;
        color: #ffffff;
        padding: 5px;
        border: 1px solid #444;
        border-radius: 5px;
    }
    QProgressBar {
        border: 2px solid #3d3d5c;
        border-radius: 10px;
        text-align: center;
        background-color: #2b2b3d;
        color: #ffffff;
    }
    QProgressBar::chunk {
        background-color: #6a0dad;
        border-radius: 10px;
    }
    """

    # Светлая тема приложения
    light_theme_stylesheet = """
    QWidget {
        background-color: #ffffff;
        color: #000000;
        font-family: Arial, sans-serif;
    }
    QPushButton {
        background-color: #f0f0f0;
        color: #000000;
        border-radius: 8px;
        padding: 10px;
        font-size: 14px;
    }
    QPushButton:hover {
        background-color: #e0e0e0;
    }
    QPushButton:pressed {
        background-color: #d0d0d0;
    }
    QLineEdit {
        background-color: #ffffff;
        color: #000000;
        padding: 5px;
        border: 1px solid #ccc;
        border-radius: 5px;
    }
    QProgressBar {
        border: 2px solid #ccc;
        border-radius: 10px;
        text-align: center;
        background-color: #f0f0f0;
        color: #000000;
    }
    QProgressBar::chunk {
        background-color: #0078d7;
        border-radius: 10px;
    }
    """

    def __init__(self):
        """
        __init__ метод для инициализации основного класса приложения.
        Инициализирует различные компоненты пользовательского интерфейса,
        настройки приложения и другие необходимые атрибуты.
        Он задает заголовок и геометрию окна, а также инициализирует настройки и необходимые страницы.
        Attributes:
        - button_temp_files: кнопка, связанная с временными файлами.
        - theme: строка, представляющая тему приложения (по умолчанию «темная»).
        - view_employees_page: страница или виджет для просмотра сотрудников.
        - edit_employees_page: страница или виджет для редактирования сотрудников.
        - Past_reports_page: страница или виджет для прошлых отчетов.
        - output_file_excelSite: заполнитель для выходного файла сайта Excel.
        - example_page: страница или виджет для примеров.
        - button_main: виджет главной кнопки.
        - stacked_widget: сложенный виджет для переключения между разными страницами.
        - button_past_reports: кнопка, связанная с прошлыми отчетами.
        - button_view_employees: кнопка для просмотра сотрудников.
        - button_settings: Кнопка доступа к настройкам.
        - main_page: главная страница приложения.
        - left_shift_pressed: логическое значение для отслеживания нажатия левой клавиши Shift.
        - right_shift_pressed: логическое значение для отслеживания нажатия правой клавиши Shift.
        - processing_completed: логическое значение, указывающее, завершена ли обработка.
        - timepath: атрибут функциональности timepath.
        - settings: словарь для хранения настроек приложения.
        - settings_page: переменная для хранения виджета страницы настроек.
        - people: переменная для хранения информации о людях.
        - zip_path: переменная для хранения пути к zip-файлам.
        - zvonki: переменная для хранения данных «звонков».
        - output_excel_path: переменная для хранения пути к выходному файлу Excel.
        - thread: переменная для хранения информации о потоках.
        - console_window: переменная для окна консоли, создаваемая только при необходимости.
        - example_window: переменная для окна примеров.
        Инициализирует пользовательский интерфейс, загружает настройки и применяет таблицу стилей.
        """
        super().__init__()
        self.processing = None
        self.edit_replacements_action = None
        self.edit_url_mapping_action = None
        self.edit_employee_timezones_action = None
        self.edit_employee_timezones_page = None
        self.edit_replacements_page = None
        self.edit_url_mapping_page = None
        self.edit_employees_action = None
        self.replacements = default_replacements.copy()
        self.url_mapping = default_url_mapping.copy()
        self.employee_timezones = default_employee_timezones.copy()
        self.company_timezone = default_company_timezone
        self.button_temp_files = None
        self.theme = 'dark'
        self.view_employees_page = None
        self.edit_employees_page = None
        self.past_reports_page = None
        self.output_file_excelSite = None
        self.examples_page = None
        self.button_main = None
        self.stacked_widget = None
        self.button_past_reports = None
        self.button_view_employees = None
        self.button_settings = None
        self.main_page = None
        self.left_shift_pressed = False
        self.right_shift_pressed = False
        self.processing_completed = False
        self.timepath = None
        self.settings = {}
        self.settings_page = None
        self.setWindowTitle("Приложение для обработки данных сотрудников")

        # Размер приложения
        self.setGeometry(100, 100, 1200, 600)
        self.setFixedSize(1200, 400)

        # Переменные для передачи информации
        self.people = None
        self.zip_path = None
        self.zvonki = None
        self.output_excel_path = None
        self.thread = None

        # Переменные для консоли
        self.console_window = None  # Создаем объект консоли только при необходимости

        # Переменная для окна примеров
        self.examples_window = None

        self.load_settings()
        self.init_ui()
        self.apply_stylesheet()
        self.update_ui_with_settings()

    def init_ui(self):
        """
        Инициализирует пользовательский интерфейс приложения. Создает и настраивает пункты меню,
        соединяет сигналы для различных действий, настраивает кнопки боковой панели и определяет основной макет.
        :return: None
        """
        # Виджеты для меню
        menubar = self.menuBar()
        about_menu = menubar.addMenu("О программе")

        examples_action = QAction("Примеры файлов", self)
        version_action = QAction("Версия", self)
        console_action = QAction("Консоль", self)

        about_menu.addAction(examples_action)
        about_menu.addAction(version_action)
        about_menu.addAction(console_action)

        # Подключаем действия к методам
        examples_action.triggered.connect(self.show_examples)
        version_action.triggered.connect(self.show_version)
        console_action.triggered.connect(self.toggle_terminal)

        # Меню "Править соответствия"
        edit_menu = menubar.addMenu("Править соответствия")

        # Действие для редактирования replacements
        self.edit_replacements_action = QAction("Править соответствия сотрудников", self)
        edit_menu.addAction(self.edit_replacements_action)
        self.edit_replacements_action.triggered.connect(self.edit_replacements)

        # Действие для редактирования url_mapping
        self.edit_url_mapping_action = QAction("Править соответствия сайтов", self)
        edit_menu.addAction(self.edit_url_mapping_action)
        self.edit_url_mapping_action.triggered.connect(self.edit_url_mapping)

        # Добавляем действие для редактирования employee_timezones
        self.edit_employee_timezones_action = QAction("Править часовые пояса сотрудников", self)
        edit_menu.addAction(self.edit_employee_timezones_action)
        self.edit_employee_timezones_action.triggered.connect(self.edit_employee_timezones)

        # Действие для редактирования файла сотрудников
        self.edit_employees_action = QAction("Редактировать файл сотрудников", self)
        edit_menu.addAction(self.edit_employees_action)
        self.edit_employees_action.triggered.connect(self.edit_employees)

        # Виджеты для работы с приложением
        self.button_main = QPushButton("Главная", self)
        self.button_past_reports = QPushButton("Прошлые отчеты", self)
        self.button_view_employees = QPushButton("Просмотр файла с сотрудниками", self)
        self.button_settings = QPushButton("Настройки", self)

        # Добавляем кнопку "Временные файлы"
        self.button_temp_files = QPushButton("Временные файлы", self)
        self.button_temp_files.setVisible(False)
        self.button_temp_files.setEnabled(False)
        self.button_temp_files.clicked.connect(self.open_temp_files_folder)

        # Подключаем кнопки боковой панели
        self.button_main.clicked.connect(self.show_main_page)
        self.button_past_reports.clicked.connect(self.show_past_reports)
        self.button_view_employees.clicked.connect(self.view_employees)
        self.button_settings.clicked.connect(self.show_settings)

        # Боковая панель
        sidebar_layout = QVBoxLayout()
        sidebar_widget = QWidget()
        sidebar_widget.setLayout(sidebar_layout)
        sidebar_layout.addWidget(self.button_main)
        sidebar_layout.addWidget(self.button_past_reports)
        sidebar_layout.addWidget(self.button_view_employees)
        sidebar_layout.addWidget(self.button_temp_files)
        sidebar_layout.addStretch()
        sidebar_layout.addWidget(self.button_settings)

        # Создаем QStackedWidget для переключения между страницами
        self.stacked_widget = QStackedWidget()

        # Создаем страницы
        self.main_page = MainPage(self)
        self.examples_page = ExamplesPage()
        # Добавляем страницы в QStackedWidget
        self.stacked_widget.addWidget(self.main_page)
        self.stacked_widget.addWidget(self.examples_page)

        # Определение сочетания клавиш для открытия консоли
        toggle_terminal_shortcut = QShortcut(QKeySequence("Shift+Tab"), self)
        toggle_terminal_shortcut.activated.connect(self.toggle_terminal)

        # Основной макет
        main_layout = QGridLayout()
        main_layout.addWidget(sidebar_widget, 0, 0, 1, 1)
        main_layout.addWidget(self.stacked_widget, 0, 1, 1, 4)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def keyPressEvent(self, event):
        """
        :param event: Объект события, содержащий информацию о событии нажатия клавиши.
        :return: None
        """
        scancode = event.nativeScanCode()
        if scancode == 42:  # Левый Shift
            self.left_shift_pressed = True
        elif scancode == 54:  # Правый Shift
            self.right_shift_pressed = True

        if self.left_shift_pressed and self.right_shift_pressed:
            self.show_temp_files_button()
        super().keyPressEvent(event)

    def keyReleaseEvent(self, event):
        """
        :param event: Объект QKeyEvent, содержащий информацию о событии выпуска ключа.
        :return: None
        """
        scancode = event.nativeScanCode()
        if scancode == 42:  # Левый Shift
            self.left_shift_pressed = False
        elif scancode == 54:  # Правый Shift
            self.right_shift_pressed = False
        super().keyReleaseEvent(event)

    def update_ui_with_settings(self):
        if self.output_excel_path:
            folder_name = os.path.basename(self.output_excel_path)
            self.main_page.folder_output_label.setText(folder_name)
        if self.settings_page:
            self.settings_page.update_theme_selection()

    def disable_mapping_editing(self):
        # Отключаем пункты меню
        self.edit_replacements_action.setEnabled(False)
        self.edit_url_mapping_action.setEnabled(False)
        self.edit_employee_timezones_action.setEnabled(False)
        self.button_view_employees.setEnabled(False)
        if self.stacked_widget.currentWidget() in [
            self.edit_replacements_page,
            self.edit_url_mapping_page,
            self.edit_employee_timezones_page
        ]:
            self.stacked_widget.currentWidget().set_disabled(True)

    def enable_mapping_editing(self):
        self.edit_replacements_action.setEnabled(True)
        self.edit_url_mapping_action.setEnabled(True)
        self.edit_employee_timezones_action.setEnabled(True)
        self.button_view_employees.setEnabled(True)
        if self.stacked_widget.currentWidget() in [
            self.edit_replacements_page,
            self.edit_url_mapping_page,
            self.edit_employee_timezones_page
        ]:
            self.stacked_widget.currentWidget().set_disabled(False)

    def convert_replacements_format(self, old_replacements):
        """
        Parameters
        ----------
        old_replacements : dict
            Словарь, в котором ключи — это исходные строки, а значения — строки замены.
        Returns
        -------
        dict
           Словарь, в котором ключи представляют собой строки замены, а значения — списки исходных строк.
        """
        new_format = {}
        for source, target in old_replacements.items():
            if target not in new_format:
                new_format[target] = []
            new_format[target].append(source)
        return new_format

    def show_temp_files_button(self):
        """
        Показывает кнопку временных файлов, если она не видна, и включает ее, если обработка завершена.
        :return: None
        """
        if not self.button_temp_files.isVisible():
            self.button_temp_files.setVisible(True)
            logger.debug("Кнопка 'Временные файлы' теперь видима.")
        if self.processing_completed:
            self.button_temp_files.setEnabled(True)  # Активируем кнопку, если обработка завершена

    # Применение темы
    def apply_stylesheet(self):
        """
        Применяет соответствующую таблицу стилей на основе текущей темы.
        Если для темы установлено значение «темная», применяется таблица стилей темной темы.
        Если для темы установлено значение «светлая», применяется таблица стилей светлой темы.
        :return: None
        """
        if self.theme == 'dark':
            self.setStyleSheet(self.dark_theme_stylesheet)
        elif self.theme == 'light':
            self.setStyleSheet(self.light_theme_stylesheet)

    # Загрузка настроек из файла
    def load_settings(self):
        """
        Загружает настройки приложения из файла конфигурации.
        :return: None
        """
        config_path = get_config_path()
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                self.settings = json.load(f)
                self.output_excel_path = self.settings.get('output_excel_path', None)
                self.theme = self.settings.get('theme', 'dark')
                flat_replacements = self.settings.get('replacements', default_replacements.copy())
                self.replacements = self.expand_replacements(flat_replacements)
                flat_url_mapping = self.settings.get('url_mapping', default_url_mapping.copy())
                self.url_mapping = self.expand_replacements(flat_url_mapping)
                self.employee_timezones = self.settings.get('employee_timezones', default_employee_timezones.copy())
                self.company_timezone = self.settings.get('company_timezone', default_company_timezone)
        except FileNotFoundError:
            self.settings = {}
            self.theme = 'dark'
            self.replacements = default_replacements.copy()
            self.url_mapping = default_url_mapping.copy()
            self.employee_timezones = default_employee_timezones.copy()
            self.company_timezone = default_company_timezone
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить настройки: {str(e)}")

    # Сохранение настроек
    def save_settings(self):
        """
        Сохраняет текущие настройки в файл конфигурации.
        :return: None
        """
        config_path = get_config_path()
        try:
            self.settings['output_excel_path'] = self.output_excel_path
            self.settings['theme'] = self.theme
            flat_replacements = self.flatten_replacements(self.replacements)  # Теперь корректно возвращает flat
            self.settings['replacements'] = flat_replacements
            flat_url_mapping = self.flatten_replacements(self.url_mapping)  # Аналогично для url_mapping
            self.settings['url_mapping'] = flat_url_mapping
            self.settings['employee_timezones'] = self.employee_timezones
            self.settings['company_timezone'] = self.company_timezone
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить настройки: {str(e)}")

    def flatten_replacements(self, replacements):
        flat = {}
        for target, sources in replacements.items():
            if isinstance(sources, list):
                for source in sources:
                    flat[source] = target
            else:
                flat[sources] = target
        return flat

    def expand_replacements(self, flat_replacements):
        expanded = {}
        for source, target in flat_replacements.items():
            if target not in expanded:
                expanded[target] = []
            expanded[target].append(source)
        return expanded

    # Кнопка Править сотрудников
    def edit_employees(self):
        """
        Проверяет, пуст ли список людей. Если это так, отображается предупреждающее сообщение.
        Если список людей не пуст, создается EditEmployeesPage и переключается
        текущий сложенный виджет.
        :return: None
        """
        if not self.people:
            QMessageBox.warning(self, "Предупреждение", "Файл с сотрудниками не выбран.")
            return
        else:
            self.edit_employees_page = EditEmployeesPage(self, self.people)
            self.stacked_widget.addWidget(self.edit_employees_page)
            self.stacked_widget.setCurrentWidget(self.edit_employees_page)

    # Метод для переключения страниц (Главная)
    def show_main_page(self):
        """
        Переключает текущий отображаемый виджет на главную страницу.
        :return: None
        """
        self.stacked_widget.setCurrentWidget(self.main_page)

    # Метод для переключения страниц (Примеры)
    def show_examples(self):
        """
        :return: None
            Этот метод устанавливает текущий виджет stacked_widget на example_page.
        """
        self.stacked_widget.setCurrentWidget(self.examples_page)

    # Открытие готового отчета
    def open_ready_report(self):
        """
        Пытается открыть файл отчета, указанный атрибутом output_file_excelSite,
        с помощью приложения по умолчанию, связанного с этим типом файла.
        Отображает сообщение об ошибке, если файл не может быть открыт,
        или предупреждение, если файл не найден.
        :return: None
        """
        if self.output_file_excelSite and os.path.exists(self.output_file_excelSite):
            try:
                os.startfile(self.output_file_excelSite)
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось открыть файл: {str(e)}")
        else:
            QMessageBox.warning(self, "Предупреждение", "Файл отчета не найден.")

    # Переименование готового отчета
    def handle_results(self, output_file_path):
        """
        :param output_file_path: Путь к выходному файлу, созданному потоком.
        :return: None.
        """
        # Получаем timepath из потока
        self.timepath = self.thread.timepath
        logger.info(f"Получены результаты из потока: {output_file_path}")
        file_name, ok = QInputDialog.getText(self, "Сохранение отчета", "Введите имя для отчета:")
        if ok and file_name:
            if not file_name.endswith('.xlsx'):
                file_name += '.xlsx'
            new_file_path = os.path.join(self.output_excel_path, file_name)
            try:
                shutil.move(output_file_path, new_file_path)
                self.output_file_excelSite = new_file_path
                self.main_page.button_open_report.setEnabled(True)
                QMessageBox.information(self, "Сохранение", f"Отчет сохранен как {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {str(e)}")
                self.output_file_excelSite = output_file_path
        else:
            self.output_file_excelSite = output_file_path
            self.main_page.button_open_report.setEnabled(True)
            QMessageBox.information(self, "Сохранение", "Отчет сохранен с исходным именем.")

    # Показ версии приложения
    def show_version(self):
        """
        Отображает информацию о версии приложения в окне сообщения.
        :return: None
        """
        version_text = (
            "Версия приложения: 1.0.0\n"
            "Дата выпуска: 18.10.2024\n"
            "Автор: sambuka_lx"
        )
        QMessageBox.information(self, "Версия", version_text)

    # Что-то там с териминалом
    def toggle_terminal(self):
        """
        Переключает видимость окна консоли. Если окно консоли не
        инициализирован, он инициализирует новый экземпляр ConsoleWindow. Если окно консоли
        в настоящее время виден, он скрывает его; в противном случае отображается окно.
        :return: None
        """
        if not self.console_window:
            self.console_window = ConsoleWindow()
        if self.console_window.isVisible():
            self.console_window.hide()
        else:
            self.console_window.show()

    # Выбор .txt файла с сотрудниками
    def select_people_file(self):
        """
        Обрабатывает выбор TXT-файла с помощью диалогового окна файла. Обновляет атрибут people и
        соответствующая метка на главной_странице, если файл выбран. Также обновляет состояние
        кнопку запуска в зависимости от нового состояния.
        Отображает окно сообщения об ошибке, если возникает исключение.
        :return: None
        """
        try:
            file, _ = QFileDialog.getOpenFileName(self, "Выбрать файл с сотрудниками", "", "Text Files (*.txt)")
            if file:
                self.people = file
                self.main_page.file_people_label.setText(os.path.basename(file))
            self.update_run_button_state()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось выбрать файл: {str(e)}")

    # Выбор .zip файла Стахановца
    def select_zip_file(self):
        """
        Обрабатывает выбор ZIP-файла с помощью диалогового окна файла. Обновляет атрибут zip_path и
        соответствующая метка на главной_странице, если файл выбран. Также обновляет состояние
        кнопку запуска в зависимости от нового состояния.
        Отображает окно сообщения об ошибке, если возникает исключение.
        :return: None
        """
        try:
            file, _ = QFileDialog.getOpenFileName(self, "Выбрать ZIP файл Стахановца", "", "ZIP Files (*.zip)")
            if file:
                self.zip_path = file
                self.main_page.file_zip_label.setText(os.path.basename(file))
            self.update_run_button_state()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось выбрать файл: {str(e)}")

    # Выбор .xlsx файла с звонками из АТС
    def select_zvonki_file(self):
        """
        Открывает диалоговое окно файла для выбора файла Excel, обновляет метку именем файла,
        и обновляет состояние кнопки запуска.
        :return: None
        """
        try:
            file, _ = QFileDialog.getOpenFileName(self, "Выбрать файл звонки из АТС", "", "Excel Files (*.xlsx)")
            if file:
                self.zvonki = file
                self.main_page.file_zvonki_label.setText(os.path.basename(file))
            self.update_run_button_state()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось выбрать файл: {str(e)}")

    # Выбор папки для сохранения
    def select_output_folder(self):
        """
        Предлагает пользователю выбрать каталог для сохранения отчетов. Если выбран каталог, обновляется выходной путь,
        обновляет метку пользовательского интерфейса с указанием выбранного имени папки,
        сохраняет настройки и обновляет состояние кнопки запуска.
        :return:
        """
        try:
            folder = QFileDialog.getExistingDirectory(self, "Выбрать папку для сохранения отчетов")
            if folder:
                self.output_excel_path = folder
                folder_name = os.path.basename(folder)
                self.main_page.folder_output_label.setText(folder_name)
                self.save_settings()
            self.update_run_button_state()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось выбрать папку: {str(e)}")

    # Активируем кнопку запуска, если выбраны все файлы
    def update_run_button_state(self):
        """
        Обновляет состояние кнопки «Запустить процесс» в зависимости от наличия необходимых путей к файлам.
        Кнопка активна только в том случае, если выбраны все необходимые файлы.
        :return: None
        """
        if all([self.people, self.zip_path, self.zvonki, self.output_excel_path]):
            self.main_page.button_run_process.setEnabled(True)
        else:
            self.main_page.button_run_process.setEnabled(False)

    # Запуск всех функций
    def run_all_functions(self):
        """
        Запускает поток обработки данных, если он еще не запущен,
        и соответствующим образом обновляет пользовательский интерфейс.
        :return: None
        """
        if self.thread is None or not self.thread.isRunning():
            try:
                self.processing = True
                self.disable_file_selection_buttons()
                self.main_page.button_run_process.setEnabled(False)
                self.main_page.button_stop.setEnabled(True)
                self.main_page.button_reset.setEnabled(False)
                self.disable_mapping_editing()
                self.disable_employee_editing()
                flat_replacements = self.flatten_replacements(self.replacements)
                flat_url_mapping = self.flatten_replacements(self.url_mapping)
                self.thread = ProcessingThread(
                    self.people,
                    self.zip_path,
                    self.zvonki,
                    self.output_excel_path,
                    flat_replacements,
                    flat_url_mapping,
                    self.employee_timezones,
                    self.company_timezone
                )
                self.thread.progress.connect(self.update_progress_bar)
                self.thread.result_signal.connect(self.handle_results)
                self.thread.status_signal.connect(self.update_status)
                self.thread.error_signal.connect(self.show_error_message)
                self.thread.finished.connect(self.on_thread_finished)
                self.thread.start()
            except ValueError as e:
                self.show_error_message(f"Ошибка валидации данных: {str(e)}")
            except Exception as e:
                self.show_error_message(f"Произошла неизвестная ошибка: {str(e)}")

    # Обновление прогресс бара
    def update_progress_bar(self, value):
        """
        :param value: Текущее значение прогресса, которое будет установлено на индикаторе выполнения.
        Это должно быть целое число от 0 до 100 включительно.
        :return: None. Этот метод обновляет индикатор выполнения и не возвращает никакого значения.
        """
        if value == 100:
            self.main_page.progress_bar.setStyleSheet("QProgressBar::chunk { background-color: green; }")
            self.main_page.result_label.setText("Готово!")
        else:
            self.main_page.progress_bar.setStyleSheet("")
        self.main_page.progress_bar.setValue(value)

    # Обновление статуса приложения
    def update_status(self, message):
        """
        :param message: Сообщение о состоянии, которое необходимо указать на метке результата.
        :return: None
        """
        self.main_page.result_label.setText(message)

    # Сброс всех выбранных файлов
    def reset_all(self):
        """
        Сбрасывает все выбранные файлы и элементы пользовательского интерфейса
        в состояние по умолчанию.
        :return: None
        """
        self.main_page.file_people_label.clear()
        self.main_page.file_zip_label.clear()
        self.main_page.file_zvonki_label.clear()
        self.main_page.folder_output_label.clear()
        self.main_page.progress_bar.setValue(0)
        self.main_page.progress_bar.setStyleSheet("")
        self.main_page.result_label.setText("")
        self.people = None
        self.zip_path = None
        self.zvonki = None
        self.output_excel_path = None
        self.thread = None
        self.update_run_button_state()
        self.main_page.button_open_report.setEnabled(False)
        self.output_file_excelSite = None

    # Остановка процесса
    def stop_thread(self):
        """
        Останавливает текущий выполняющийся поток,
        ждет его завершения и соответствующим образом обновляет пользовательский интерфейс.
        :return: None
        """
        try:
            if self.thread and self.thread.isRunning():
                self.thread.stop()
                self.thread.wait(5000)
                if self.thread.isRunning():
                    logger.warning("Поток не остановился вовремя, принудительно завершаем.")
                    self.thread.terminate()
                self.main_page.button_stop.setEnabled(False)
                self.main_page.progress_bar.setStyleSheet("QProgressBar::chunk { background-color: red; }")
                self.main_page.result_label.setText("Процесс остановлен.")
                self.enable_mapping_editing()
                self.enable_employee_editing()
                self.on_thread_finished()
                self.on_thread_finished()
                logger.warning("YOU KILL ME! --<3---|-")
        except Exception as e:
            self.show_error_message(f"Ошибка при остановке процесса: {str(e)}")

    # Обработчик завершения потока
    def on_thread_finished(self):
        """
        Включает и отключает кнопки пользовательского интерфейса
        в зависимости от завершения процесса потока и перемещает файл журнала ошибок.
        :return: None
        """
        self.processing = False  # Сбрасываем флаг, что процесс завершен
        self.main_page.button_run_process.setEnabled(True)
        self.main_page.button_stop.setEnabled(False)
        self.main_page.button_reset.setEnabled(True)
        self.processing_completed = True
        if self.button_temp_files.isVisible():
            self.button_temp_files.setEnabled(True)
        self.enable_mapping_editing()
        self.enable_employee_editing()
        self.enable_file_selection_buttons()
        self.move_errors_log()

    def move_errors_log(self):
        """
        Перемещает файл error.log в указанный каталог в timepath, если он существует.
        Закрывает открытые обработчики файлов журналов, связанные с error.log,
        перед попыткой перемещения файла.
        Регистрирует информационное сообщение после успешного перемещения
        или сообщение об ошибке в случае возникновения исключения.
        :return: None
        """
        if self.timepath:
            try:
                for handler in logger.handlers[:]:
                    if isinstance(handler, logging.FileHandler) and handler.baseFilename.endswith('errors.log'):
                        handler.close()
                        logger.removeHandler(handler)
                errors_log_path = os.path.join(os.getcwd(), 'errors.log')
                if os.path.exists(errors_log_path):
                    shutil.move(errors_log_path, self.timepath)
                    logger.info(f"errors.log перемещен в {self.timepath}")
            except Exception as e:
                logger.error(f"Ошибка при перемещении errors.log: {str(e)}")

    def open_temp_files_folder(self):
        """
        Открывает папку временных файлов, если она существует.
        Пытается открыть каталог, указанный в `self.timepath`,
        используя файловый менеджер по умолчанию.
        Если каталог не существует или не может быть открыт,
        выводится соответствующее сообщение.
        :return: None
        """
        if self.timepath and os.path.exists(self.timepath):
            try:
                os.startfile(self.timepath)
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось открыть папку: {str(e)}")
        else:
            QMessageBox.warning(self, "Предупреждение", "Папка временных файлов не найдена.")

    # Показ ошибок
    def show_error_message(self, message):
        """
        :param message: Сообщение об ошибке, которое будет отображено.
        :return: None
        """
        error_dialog = QMessageBox()
        error_dialog.setIcon(QMessageBox.Icon.Critical)
        error_dialog.setWindowTitle("Ошибка")
        error_dialog.setText(message)
        error_dialog.exec()
        self.main_page.button_open_report.setEnabled(False)  # Отключаем кнопку

    # Дополнительные методы для других кнопок боковой панели
    def show_past_reports(self):
        """
        Управляет отображением прошлых отчетов, проверяя,
        установлен ли путь вывода для сохранения отчетов.
        Если путь не задан, отображается предупреждающее сообщение.
        В противном случае он инициализирует и отображает
        Виджет PastReportsPage.
        :return: None
        """
        if not self.output_excel_path:
            QMessageBox.warning(self, "Предупреждение", "Путь для сохранения отчетов не задан.")
            return
        else:
            self.past_reports_page = PastReportsPage(self)
            self.stacked_widget.addWidget(self.past_reports_page)
            self.stacked_widget.setCurrentWidget(self.past_reports_page)

    # Просмотр фала сотрудников
    def view_employees(self):
        """
        :param peple: Файл TXT
        :return: None
        """
        if not self.people:
            QMessageBox.warning(self, "Предупреждение", "Файл с сотрудниками не выбран.")
            return
        else:
            self.view_employees_page = EditEmployeesPage(self, self.people)
            self.view_employees_page.text_edit.setReadOnly(True)
            self.stacked_widget.addWidget(self.view_employees_page)
            self.stacked_widget.setCurrentWidget(self.view_employees_page)

    # Просмотр настроек
    def show_settings(self):
        """
        Нужно сконфигурировать и отобразить страницу настроек в приложении.
        :return: None
        """
        if not self.settings_page:
            self.settings_page = SettingsPage(self)
        self.settings_page.set_change_path_button_enabled(not self.processing)
        self.stacked_widget.addWidget(self.settings_page)
        self.stacked_widget.setCurrentWidget(self.settings_page)

    def edit_replacements(self):
        """
        Инициализируйте и отобразите страницу редактирования замен в пользовательском интерфейсе.
        Этот метод инициализирует новый экземпляр EditReplacementsPage и добавляет его в составной виджет.
        установка его в качестве текущего виджета для отображения страницы редактирования замен.
        Attributes
        ----------
        self : object
            Экземпляр класса, содержащий составной виджет и страницу редактирования замен.
        """
        self.edit_replacements_page = EditReplacementsPage(self)
        self.stacked_widget.addWidget(self.edit_replacements_page)
        self.stacked_widget.setCurrentWidget(self.edit_replacements_page)

    def edit_url_mapping(self):
        """
        Инициализирует EditUrlMappingPage и устанавливает его в качестве текущего виджета.
        в сложенном виджете.
        Parameters None
        Returns None
        """
        self.edit_url_mapping_page = EditUrlMappingPage(self)
        self.stacked_widget.addWidget(self.edit_url_mapping_page)
        self.stacked_widget.setCurrentWidget(self.edit_url_mapping_page)

    def edit_employee_timezones(self):
        """
        Переход на страницу «Редактирование часовых поясов сотрудников».
        Attributes
        ----------
        self : object
            Reference to the current instance of the class.
        Methods
        -------
        edit_employee_timezones_page : method
            Инициализирует страницу редактирования часовых поясов сотрудников.
        stacked_widget.addWidget : method
            Добавляет страницу редактирования часовых поясов сотрудников в составной виджет.
        stacked_widget.setCurrentWidget : method
            Устанавливает текущий виджет в составном виджете на страницу редактирования часовых поясов сотрудников.
        """
        self.edit_employee_timezones_page = EditEmployeeTimezonesPage(self)
        self.stacked_widget.addWidget(self.edit_employee_timezones_page)
        self.stacked_widget.setCurrentWidget(self.edit_employee_timezones_page)

    def disable_employee_editing(self):
        self.edit_employees_action.setEnabled(False)
        self.button_view_employees.setEnabled(False)
        if self.stacked_widget.currentWidget() == self.edit_employees_page:
            self.edit_employees_page.set_disabled(True)

    def enable_employee_editing(self):
        self.edit_employees_action.setEnabled(True)
        self.button_view_employees.setEnabled(True)
        if self.stacked_widget.currentWidget() == self.edit_employees_page:
            self.edit_employees_page.set_disabled(False)

    def disable_file_selection_buttons(self):
        self.main_page.set_file_selection_buttons_enabled(False)
        if self.settings_page:
            self.settings_page.set_change_path_button_enabled(False)

    def enable_file_selection_buttons(self):
        self.main_page.set_file_selection_buttons_enabled(True)
        if self.settings_page:
            self.settings_page.set_change_path_button_enabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
