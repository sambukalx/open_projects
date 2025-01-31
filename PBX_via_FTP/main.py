"""
Все права защищены (c) 2024.
Данный скрипт представляет собой FTP-систему АТС для отдела аналитики.
Код предназначен исключительно для ознакомления.
Любое распространение и/или модификация без согласия автора запрещены.
"""


import os
import sys
import re
import json
import ctypes
import tempfile
import openpyxl
import qdarkstyle

from datetime import datetime, timedelta, date
from ftplib import FTP
from pydub import AudioSegment
from openpyxl.styles import Alignment
from PyQt5.QtGui import QMovie
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QUrl, QDate, QObject
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QTableWidget,
    QTableWidgetItem, QLabel, QLineEdit, QDialog, QDialogButtonBox, QProgressDialog,
    QComboBox, QSlider, QMessageBox, QToolBar, QAction, QFileDialog, QInputDialog,
    QHBoxLayout, QDateEdit, QFormLayout, QHeaderView
)

# Критерии, по которым аналитики делают отметки
CRITERIA = [
    "XXX",
    "YYY",
    "ZZZ"
]

CONFIG_DIR = os.path.join(os.getenv("APPDATA"), "ProsluskaZV")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")
LOG_PATH = os.path.join(CONFIG_DIR, "app.log")

tempfile.tempdir = os.path.join(CONFIG_DIR, "Temp")
os.makedirs(tempfile.tempdir, exist_ok=True)

def write_log(message: str):
    """
    Запись текстового сообщения в лог-файл.
    """
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().isoformat()} - {message}\n")


def load_config():
    """
    Загружает JSON-конфигурацию из CONFIG_PATH.
    Если файл не найден или поврежден, создает конфиг по умолчанию.
    """
    if not os.path.exists(CONFIG_PATH):
        os.makedirs(CONFIG_DIR, exist_ok=True)
        write_log("Config not found, creating default config.")
        return {
            "login": "",
            "folder_info": {},
            "downloads": {},
            "download_path": "Загрузки",
            "highlight_threshold": 40,
            "account_mapping": {}
        }
    try:
        with open(CONFIG_PATH, "r", encoding='utf-8') as f:
            data = json.load(f)
        if "download_path" not in data:
            data["download_path"] = "Загрузки"
        if "highlight_threshold" not in data:
            data["highlight_threshold"] = 40
        if "account_mapping" not in data:
            data["account_mapping"] = {}
        return data
    except json.JSONDecodeError as e:
        write_log(f"Ошибка чтения конфигурации: {e}")
        return {
            "login": "",
            "folder_info": {},
            "downloads": {},
            "download_path": "Загрузки",
            "highlight_threshold": 40,
            "account_mapping": {}
        }


def save_config(data):
    """
    Сохраняет текущую конфигурацию в файл JSON.
    """
    try:
        with open(CONFIG_PATH, "w", encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        write_log("Конфигурация сохранена.")
    except Exception as e:
        write_log(f"Ошибка сохранения конфигурации: {e}")


class DownloadThread(QThread):
    """
    Фоновый поток для скачивания файлов (записей звонков) с FTP-сервера.
    """
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, config, host, user, passwd, days_to_download=360, parent=None):
        super().__init__(parent)
        self.config = config
        self.host = host
        self.user = user
        self.passwd = passwd
        self.days_to_download = days_to_download
        self.running = True

    def run(self):
        write_log("Starting download thread.")
        try:
            ftp = FTP(self.host)
            ftp.set_pasv(True)
            ftp.login(self.user, self.passwd)
            self.status.emit("Подключение к FTP выполнено.")
            write_log("Connected to FTP.")
        except Exception as e:
            self.status.emit(f"Ошибка подключения к FTP: {e}")
            write_log(f"FTP connection error: {e}")
            self.finished.emit()
            return

        target_date = datetime.now() - timedelta(days=self.days_to_download)
        try:
            ftp.cwd("/recordings")
            folders = ftp.nlst()
        except Exception as e:
            self.status.emit(f"Ошибка получения списка папок: {e}")
            write_log(f"Error getting folder list: {e}")
            ftp.quit()
            self.finished.emit()
            return

        # Фильтрация папок по дате
        folders_to_download = [
            f for f in folders
            if self.is_date_format(f) and f >= target_date.strftime("%Y-%m-%d")
        ]

        zvonki_dir = os.path.join(CONFIG_DIR, "Zvonki")
        os.makedirs(zvonki_dir, exist_ok=True)

        total_files = 0
        files_downloaded = 0
        for folder in folders_to_download:
            try:
                folder_files = ftp.nlst(f"/recordings/{folder}")
                folder_files = [f for f in folder_files if f.endswith('.mp3')]
                total_files += len(folder_files)
            except Exception as e:
                write_log(f"Error accessing folder {folder}: {e}")

        for folder in folders_to_download:
            if not self.running:
                break
            try:
                ftp.cwd(f"/recordings/{folder}")
                files = ftp.nlst()
                files = [f for f in files if f.endswith('.mp3')]
                for file in files:
                    if not self.running:
                        break
                    local_path = os.path.join(zvonki_dir, file)
                    if not os.path.exists(local_path):
                        with open(local_path, 'wb') as f_local:
                            ftp.retrbinary(f"RETR {file}", f_local.write)
                        self.config["downloads"][file] = local_path
                        save_config(self.config)
                        write_log(f"Downloaded file: {file}")
                    files_downloaded += 1
                    progress = int((files_downloaded / total_files) * 100)
                    self.progress.emit(progress)
                    self.status.emit(f"Загружено {files_downloaded} из {total_files} файлов")
            except Exception as e:
                write_log(f"File download error: {e}")

        ftp.quit()
        self.status.emit("Загрузка завершена.")
        write_log("Download finished.")
        self.finished.emit()

    def is_date_format(self, folder_name):
        try:
            datetime.strptime(folder_name, "%Y-%m-%d")
            return True
        except ValueError:
            return False


class DurationLoaderThread(QThread):
    """
    Фоновый поток для загрузки и определения длительности mp3-файлов,
    если они еще не были определены при предыдущих анализах.
    """
    updated = pyqtSignal()

    def __init__(self, config):
        super().__init__()
        self.config = config
        self.running = True

    def run(self):
        write_log("Starting duration loader thread.")
        changed = False
        for date_key, folder_data in self.config.get("folder_info", {}).items():
            calls = folder_data.get("calls", [])
            for call in calls:
                if not self.running:
                    break
                if call["duration"] in ("Неизвестно", "Ошибка"):
                    filename = call["filename"]
                    local_path = self.config["downloads"].get(filename)
                    if local_path and os.path.exists(local_path):
                        try:
                            audio = AudioSegment.from_file(local_path, format="mp3")
                            duration_seconds = len(audio) / 1000.0
                            duration_str = self.format_duration(duration_seconds)
                            call["duration"] = duration_str
                            changed = True
                            write_log(f"Duration updated for {filename}: {duration_str}")
                        except Exception as e:
                            write_log(f"Error getting duration for {filename}: {e}")
                            call["duration"] = "Неизвестно"
            if changed:
                save_config(self.config)
                changed = False
                self.updated.emit()

    def format_duration(self, seconds):
        seconds = int(seconds)
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        return f"{h:02d}:{m:02d}:{s:02d}"


class LoginDialog(QDialog):
    """
    Диалоговое окно для ввода логина и пароля к FTP.
    """
    def __init__(self, current_login=""):
        super().__init__()
        self.setWindowTitle("Введите логин и пароль")
        self.setFixedSize(400, 200)

        layout = QVBoxLayout()
        self.login_input = QLineEdit(self)
        self.login_input.setPlaceholderText("Логин")
        self.login_input.setText(current_login)

        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText("Пароль")
        self.password_input.setEchoMode(QLineEdit.Password)

        layout.addWidget(QLabel("Логин:"))
        layout.addWidget(self.login_input)
        layout.addWidget(QLabel("Пароль:"))
        layout.addWidget(self.password_input)

        self.toggle_password_button = QPushButton("Показать пароль")
        self.toggle_password_button.clicked.connect(self.toggle_password_visibility)
        layout.addWidget(self.toggle_password_button)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def toggle_password_visibility(self):
        if self.password_input.echoMode() == QLineEdit.Password:
            self.password_input.setEchoMode(QLineEdit.Normal)
            self.toggle_password_button.setText("Скрыть пароль")
        else:
            self.password_input.setEchoMode(QLineEdit.Password)
            self.toggle_password_button.setText("Показать пароль")

    def get_credentials(self):
        return self.login_input.text(), self.password_input.text()


class SettingsDialog(QDialog):
    """
    Окно «Настройки»: путь для сохранения файлов, порог длительности, таблица соответствий аккаунтов.
    """
    def __init__(self, config):
        super().__init__()
        self.setWindowTitle("Настройки")
        self.setFixedSize(600, 400)
        self.config = config

        from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QGroupBox,
                                     QTableWidget, QTableWidgetItem, QPushButton)

        main_layout = QVBoxLayout(self)
        general_group = QGroupBox("Общие настройки")
        general_layout = QFormLayout()
        general_group.setLayout(general_layout)

        self.path_line_edit = QLineEdit(self)
        self.path_line_edit.setText(self.config.get("download_path", "Загрузки"))
        choose_button = QPushButton("Выбрать...")
        choose_button.clicked.connect(self.choose_path)
        path_layout = QHBoxLayout()
        path_layout.addWidget(self.path_line_edit)
        path_layout.addWidget(choose_button)

        self.threshold_line_edit = QLineEdit(self)
        self.threshold_line_edit.setText(str(self.config.get("highlight_threshold", 40)))

        general_layout.addRow("Путь для сохранения файлов:", path_layout)
        general_layout.addRow("Порог длительности (сек.):", self.threshold_line_edit)
        main_layout.addWidget(general_group)

        mapping_group = QGroupBox("Соответствия Аккаунт/Номер → Имя сотрудника")
        mapping_layout = QVBoxLayout()
        mapping_group.setLayout(mapping_layout)

        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(2)
        self.mapping_table.setHorizontalHeaderLabels(["Аккаунт/Номер", "Имя сотрудника"])
        self.mapping_table.setEditTriggers(QTableWidget.DoubleClicked)
        self.load_mappings()

        btn_add = QPushButton("Добавить")
        btn_add.clicked.connect(self.add_mapping)
        btn_del = QPushButton("Удалить")
        btn_del.clicked.connect(self.delete_mapping)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_del)
        btn_layout.addStretch()

        mapping_layout.addWidget(self.mapping_table)
        mapping_layout.addLayout(btn_layout)
        main_layout.addWidget(mapping_group)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.save_settings)
        buttons.rejected.connect(self.reject)
        main_layout.addWidget(buttons)
        self.setLayout(main_layout)

    def choose_path(self):
        directory = QFileDialog.getExistingDirectory(self, "Выберите папку для загрузок", self.path_line_edit.text())
        if directory:
            self.path_line_edit.setText(directory)

    def load_mappings(self):
        account_mapping = self.config.get("account_mapping", {})
        self.mapping_table.setRowCount(0)
        for i, (acc, name) in enumerate(account_mapping.items()):
            self.mapping_table.insertRow(i)
            self.mapping_table.setItem(i, 0, QTableWidgetItem(acc))
            self.mapping_table.setItem(i, 1, QTableWidgetItem(name))

    def add_mapping(self):
        row = self.mapping_table.rowCount()
        self.mapping_table.insertRow(row)
        self.mapping_table.setItem(row, 0, QTableWidgetItem(""))
        self.mapping_table.setItem(row, 1, QTableWidgetItem(""))

    def delete_mapping(self):
        selected = self.mapping_table.selectionModel().selectedRows()
        for s in reversed(selected):
            self.mapping_table.removeRow(s.row())

    def save_settings(self):
        new_path = self.path_line_edit.text().strip()
        try:
            new_threshold = int(self.threshold_line_edit.text().strip())
        except:
            new_threshold = 40

        self.config["download_path"] = new_path
        self.config["highlight_threshold"] = new_threshold

        new_mapping = {}
        for row in range(self.mapping_table.rowCount()):
            acc_item = self.mapping_table.item(row, 0)
            name_item = self.mapping_table.item(row, 1)
            if acc_item and name_item:
                acc = acc_item.text().strip()
                name = name_item.text().strip()
                if acc and name:
                    new_mapping[acc] = name

        self.config["account_mapping"] = new_mapping
        save_config(self.config)
        self.accept()


class RebuildWorker(QObject):
    """
    Фоновый объект для пересборки (rebuild) структуры folder_info.
    """
    finished = pyqtSignal()

    def __init__(self, config):
        super().__init__()
        self.config = config

    def run(self):
        self._rebuild_folder_info()
        self.finished.emit()

    def format_duration(self, seconds):
        seconds = int(seconds)
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        return f"{h:02d}:{m:02d}:{s:02d}"

    def _rebuild_folder_info(self):
        old_folder_info = self.config.get("folder_info", {})
        new_folder_info = {}
        last_week_date = datetime.now().date() - timedelta(days=7)

        for filename, local_path in self.config.get("downloads", {}).items():
            parsed = self.parse_filename(filename)
            if not parsed:
                continue

            call_date_str = parsed["date"]
            call_date = datetime.strptime(call_date_str, "%Y-%m-%d").date()
            day_of_week = datetime.strptime(call_date_str, "%Y-%m-%d").strftime("%A")

            if call_date_str not in new_folder_info:
                new_folder_info[call_date_str] = {
                    "day": day_of_week,
                    "incoming": 0,
                    "outgoing": 0,
                    "calls": []
                }

            call_type = "Входящий" if parsed["type"] == "in" else "Исходящий"
            number = re.sub(r"\D", "", parsed["number"])
            call_time = f"{parsed['date']} {parsed['time']}"

            if call_date > last_week_date:
                try:
                    audio = AudioSegment.from_file(local_path, format="mp3")
                    duration_seconds = len(audio) / 1000.0
                    duration_str = self.format_duration(duration_seconds)
                except Exception as e:
                    write_log(f"Ошибка определения длительности {filename}: {e}")
                    duration_str = "Неизвестно"
            else:
                duration_str = "Неизвестно"

            old_marks = {}
            for old_date, old_data in old_folder_info.items():
                for old_call in old_data.get("calls", []):
                    if old_call.get("filename") == filename:
                        old_marks = old_call.get("marks", {})
                        break

            call_data = {
                "filename": filename,
                "type": call_type,
                "number": number,
                "account": parsed["account"],
                "datetime": call_time,
                "duration": duration_str,
                "marks": old_marks
            }

            if parsed["type"] == "in":
                new_folder_info[call_date_str]["incoming"] += 1
            else:
                new_folder_info[call_date_str]["outgoing"] += 1

            new_folder_info[call_date_str]["calls"].append(call_data)

        self.config["folder_info"] = new_folder_info
        save_config(self.config)

    def parse_filename(self, filename):
        pattern = (r"(?P<account>[\w\d]+)_(?P<type>in|out)_(?P<date>\d{4}_\d{2}_\d{2})-"
                   r"(?P<time>\d{2}_\d{2}_\d{2})_(?P<number>[\d]+)")
        match = re.match(pattern, filename)
        if match:
            data = match.groupdict()
            data["date"] = data["date"].replace("_", "-")
            data["time"] = data["time"].replace("_", ":")
            data["number"] = re.sub(r"\D", "", data["number"])
            return data
        else:
            return None


class FTPApp(QMainWindow):
    """
    Главный класс GUI.
    """
    def __init__(self):
        super().__init__()
        self.config = load_config()
        self.ftplog = self.config.get("login", "")
        self.pas = None

        self.setWindowTitle(f"АТС Система - {self.ftplog}")
        self.setGeometry(200, 200, 1500, 700)

        # Тулбар
        self.toolbar = QToolBar("Main Toolbar")
        self.addToolBar(self.toolbar)

        settings_action = QAction("Настройки", self)
        settings_action.triggered.connect(self.open_settings)
        self.toolbar.addAction(settings_action)

        export_action = QAction("Выгрузить в XLSX", self)
        export_action.triggered.connect(self.export_to_xlsx)
        self.toolbar.addAction(export_action)

        # Основной лэйаут
        self.layout_main = QVBoxLayout()
        central = QWidget()
        central.setLayout(self.layout_main)
        self.setCentralWidget(central)

        # Поля для фильтрации
        self.search_number_input = QLineEdit()
        self.search_number_input.setPlaceholderText("Поиск по номеру")

        self.from_number_input = QLineEdit()
        self.from_number_input.setPlaceholderText("Откуда звонили")

        self.to_number_input = QLineEdit()
        self.to_number_input.setPlaceholderText("Куда звонили")

        self.period_box = QComboBox()
        periods = ["Все", "Сегодня", "Вчера", "Текущая неделя", "Прошлая неделя",
                   "Текущий месяц", "Прошлый месяц", "Произвольный период"]
        for p in periods:
            self.period_box.addItem(p)

        self.start_date_edit = QDateEdit(calendarPopup=True)
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.hide()

        self.end_date_edit = QDateEdit(calendarPopup=True)
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.hide()

        self.period_box.currentIndexChanged.connect(self.on_period_changed)

        self.call_type_box = QComboBox()
        self.call_type_box.addItems(["Все", "Входящий", "Исходящий", "Пропущенный"])

        self.duration_box = QComboBox()
        durations = ["Все", "От 5 секунд", "От 10 секунд", "От 20 секунд", "От 30 секунд",
                     "От 40 секунд", "От 60 секунд (1 мин)", "Другая..."]
        for d in durations:
            self.duration_box.addItem(d)

        self.custom_duration_input = QLineEdit()
        self.custom_duration_input.setPlaceholderText("Другая длительность (сек)")
        self.custom_duration_input.hide()
        self.duration_box.currentIndexChanged.connect(self.duron)

        self.route_input = QLineEdit()
        self.route_input.setPlaceholderText("Маршрут (account)")

        self.apply_button = QPushButton("Применить")
        self.apply_button.clicked.connect(self.apply_filters)

        self.reset_button = QPushButton("Сбросить")
        self.reset_button.clicked.connect(self.reset_filters)

        self.criteria_filter_box = QComboBox()
        self.criteria_filter_box.addItem("Без фильтра по критерию")
        for c in CRITERIA:
            self.criteria_filter_box.addItem(c)

        self.criteria_color_box = QComboBox()
        self.criteria_color_box.addItem("Все")
        self.criteria_color_box.addItem("Зеленый")
        self.criteria_color_box.addItem("Красный")

        # Верхняя панель фильтров
        self.search_widget = QWidget()
        self.search_main_layout = QVBoxLayout()
        self.search_widget.setLayout(self.search_main_layout)

        top_line_layout = QHBoxLayout()
        top_line_layout.addWidget(QLabel("Поиск:"))
        top_line_layout.addWidget(self.search_number_input)
        top_line_layout.addWidget(QLabel("Откуда:"))
        top_line_layout.addWidget(self.from_number_input)
        top_line_layout.addWidget(QLabel("Куда:"))
        top_line_layout.addWidget(self.to_number_input)
        top_line_layout.addWidget(QLabel("Период:"))
        top_line_layout.addWidget(self.period_box)
        top_line_layout.addWidget(self.start_date_edit)
        top_line_layout.addWidget(self.end_date_edit)
        top_line_layout.addWidget(QLabel("Тип:"))
        top_line_layout.addWidget(self.call_type_box)
        top_line_layout.addWidget(QLabel("Длит.:"))
        top_line_layout.addWidget(self.duration_box)
        top_line_layout.addWidget(self.custom_duration_input)
        top_line_layout.addWidget(QLabel("Маршрут:"))
        top_line_layout.addWidget(self.route_input)

        self.search_main_layout.addLayout(top_line_layout)

        # Вторая строка фильтров (критерии)
        criteria_line_layout = QHBoxLayout()
        criteria_line_layout.addWidget(QLabel("Критерий:"))
        criteria_line_layout.addWidget(self.criteria_filter_box)
        criteria_line_layout.addWidget(QLabel("Состояние:"))
        criteria_line_layout.addWidget(self.criteria_color_box)
        self.search_main_layout.addLayout(criteria_line_layout)

        # Третья строка (кнопки)
        buttons_line_layout = QHBoxLayout()
        buttons_line_layout.addWidget(self.apply_button)
        buttons_line_layout.addWidget(self.reset_button)
        self.search_main_layout.addLayout(buttons_line_layout)

        self.layout_main.addWidget(self.search_widget)

        # Информация о процессе
        self.info_label = QLabel("Приложение загружает данные...")
        self.layout_main.addWidget(self.info_label)

        self.back_button = QPushButton("Назад")
        self.back_button.setEnabled(False)
        self.back_button.clicked.connect(self.go_back)
        self.layout_main.addWidget(self.back_button)

        # Таблица «папок» (дат), когда звонки сделаны
        self.folder_table = QTableWidget()
        self.folder_table.setColumnCount(5)
        self.folder_table.setHorizontalHeaderLabels(["Дата", "День недели", "Входящих", "Исходящих", "Всего"])
        self.folder_table.cellDoubleClicked.connect(self.open_folder_from_table)
        self.folder_table.cellClicked.connect(self.display_folder_info_from_table)
        self.layout_main.addWidget(self.folder_table)

        self.info_label_folder = QLabel("Информация о папке:")
        self.layout_main.addWidget(self.info_label_folder)

        # Таблица с деталями звонков
        base_columns = 8
        total_columns = base_columns + len(CRITERIA)
        self.call_table = QTableWidget()
        self.call_table.setColumnCount(total_columns)
        headers = ["№", "Тип", "Номер", "Аккаунт", "Длит.", "Время", "Прослушать", "Скачать"] + CRITERIA
        self.call_table.setHorizontalHeaderLabels(headers)
        self.call_table.hide()
        self.layout_main.addWidget(self.call_table)

        # Медiaplayer + кнопки
        self.player = QMediaPlayer(self)
        self.player.durationChanged.connect(self.on_duration_changed2)
        self.player.positionChanged.connect(self.on_position_changed)
        self.player.stateChanged.connect(self.on_state_changed)

        controls_layout = QHBoxLayout()
        self.play_button = QPushButton("▶/⏸")
        self.play_button.clicked.connect(self.toggle_play_pause)
        controls_layout.addWidget(self.play_button)

        self.stop_button = QPushButton("■")
        self.stop_button.clicked.connect(self.stop_playback)
        controls_layout.addWidget(self.stop_button)

        self.position_slider = QSlider(Qt.Horizontal)
        self.position_slider.sliderMoved.connect(self.set_player_position)
        self.layout_main.addWidget(self.position_slider)

        self.speed_box = QComboBox()
        for sp in ["0.5x", "0.7x", "1.0x", "1.2x", "1.5x", "1.7x", "2.0x", "3.0x"]:
            self.speed_box.addItem(sp)
        self.speed_box.setCurrentText("1.0x")
        self.speed_box.currentTextChanged.connect(self.change_speed)
        controls_layout.addWidget(QLabel("Скорость:"))
        controls_layout.addWidget(self.speed_box)

        self.layout_main.addLayout(controls_layout)

        self.time_label = QLabel("00:00:00 / 00:00:00")
        self.layout_main.addWidget(self.time_label)

        self.current_path = "/recordings"
        self.previous_paths = []
        self.download_thread = None
        self.duration_thread = None
        self.progress_dialog = None
        self.total_duration = 0
        self.current_playing_row = None
        self.current_playing_button = None

        self.check_password()

    # -------------------------- Методы GUI -------------------------- #
    def on_duration_changed2(self, duration):
        self.total_duration = duration
        self.position_slider.setRange(0, duration)

    def stop_playback(self):
        if self.player.state() != QMediaPlayer.StoppedState:
            self.player.stop()

    def on_period_changed(self):
        p = self.period_box.currentText()
        if p == "Произвольный период":
            self.start_date_edit.show()
            self.end_date_edit.show()
        else:
            self.start_date_edit.hide()
            self.end_date_edit.hide()

    def duron(self):
        d = self.duration_box.currentText()
        if d == "Другая...":
            self.custom_duration_input.show()
        else:
            self.custom_duration_input.hide()

    def check_password(self):
        dialog = LoginDialog(current_login=self.ftplog)
        if dialog.exec_() == QDialog.Accepted:
            self.ftplog, self.pas = dialog.get_credentials()
            self.config["login"] = self.ftplog
            save_config(self.config)
            QTimer.singleShot(0, self.start_initial_download)
        else:
            QApplication.quit()

    def start_initial_download(self):
        self.download_thread = DownloadThread(self.config, "XXX", self.ftplog, self.pas, 360)
        self.progress_dialog = QProgressDialog("Идет загрузка файлов...", "Отмена", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setValue(0)
        self.progress_dialog.show()

        self.download_thread.progress.connect(self.progress_dialog.setValue)
        self.download_thread.status.connect(self.info_label.setText)
        self.progress_dialog.canceled.connect(self.cancel_download)
        self.download_thread.finished.connect(self.finish_initial_download)
        self.download_thread.start()

    def cancel_download(self):
        if self.download_thread and self.download_thread.isRunning():
            self.download_thread.running = False
            self.download_thread.wait()
            self.info_label.setText("Загрузка отменена пользователем.")
        if self.progress_dialog:
            self.progress_dialog.close()

    def finish_initial_download(self):
        self.info_label.setText("Загрузка завершена.")
        if self.progress_dialog:
            self.progress_dialog.close()

        self.download_thread = None
        self.show_custom_blocker("Обновление данных...")

        self.rebuild_thread = QThread()
        self.rebuild_worker = RebuildWorker(self.config)
        self.rebuild_worker.moveToThread(self.rebuild_thread)

        self.rebuild_thread.started.connect(self.rebuild_worker.run)
        self.rebuild_worker.finished.connect(self.rebuild_thread.quit)
        self.rebuild_worker.finished.connect(self.rebuild_thread.deleteLater)
        self.rebuild_worker.finished.connect(self.on_rebuild_finished)

        self.rebuild_thread.start()

    def on_rebuild_finished(self):
        self.update_folder_table_from_config()
        self.start_duration_loading()
        self.hide_custom_blocker()

    def start_duration_loading(self):
        self.duration_thread = DurationLoaderThread(self.config)
        self.duration_thread.updated.connect(self.on_duration_updated)
        self.duration_thread.start()

    def on_duration_updated(self):
        if self.call_table.isVisible():
            parts = self.current_path.split("/")
            current_folder = parts[-1] if len(parts) > 1 else None
            if current_folder and current_folder in self.config.get("folder_info", {}):
                calls = self.config["folder_info"][current_folder].get("calls", [])
                self.update_call_table_from_config(calls)

    def apply_filters(self):
        calls = self.get_all_calls_from_config()
        filtered = self.filter_calls(calls)
        self.show_calls(filtered)

    def reset_filters(self):
        self.search_number_input.clear()
        self.from_number_input.clear()
        self.to_number_input.clear()
        self.period_box.setCurrentIndex(0)
        self.call_type_box.setCurrentIndex(0)
        self.duration_box.setCurrentIndex(0)
        self.custom_duration_input.clear()
        self.custom_duration_input.hide()
        self.route_input.clear()
        self.update_folder_table_from_config()

    def get_all_calls_from_config(self):
        all_calls = []
        folder_info = self.config.get("folder_info", {})
        for date_key, data in folder_info.items():
            calls = data.get("calls", [])
            for c in calls:
                c["date"] = date_key
                if "marks" not in c:
                    c["marks"] = {}
                all_calls.append(c)
        return all_calls

    def filter_calls(self, calls):
        search_number = self.search_number_input.text().strip()
        from_number = self.from_number_input.text().strip()
        to_number = self.to_number_input.text().strip()
        route = self.route_input.text().strip()
        account_mapping = self.config.get("account_mapping", {})
        period = self.period_box.currentText()
        start_date, end_date = self.get_period_dates(period)
        call_type = self.call_type_box.currentText()

        duration_filter = self.duration_box.currentText()
        custom_duration = None
        if duration_filter == "Другая...":
            try:
                custom_duration = int(self.custom_duration_input.text().strip())
            except:
                custom_duration = None

        selected_criterion = self.criteria_filter_box.currentText()
        selected_color = self.criteria_color_box.currentText()

        filtered = []
        for call in calls:
            call_date = datetime.strptime(call["date"], "%Y-%m-%d").date()
            if start_date and end_date and not (start_date <= call_date <= end_date):
                continue

            if call_type == "Входящий" and call["type"] != "Входящий":
                continue
            if call_type == "Исходящий" and call["type"] != "Исходящий":
                continue
            if call_type == "Пропущенный":
                dur_sec = self.duration_to_seconds(call["duration"]) or 0
                if call["type"] != "Входящий" or dur_sec > 0:
                    continue

            dur_sec = self.duration_to_seconds(call["duration"]) or 0
            if duration_filter == "От 5 секунд" and dur_sec < 5:
                continue
            if duration_filter == "От 10 секунд" and dur_sec < 10:
                continue
            if duration_filter == "От 20 секунд" and dur_sec < 20:
                continue
            if duration_filter == "От 30 секунд" and dur_sec < 30:
                continue
            if duration_filter == "От 40 секунд" and dur_sec < 40:
                continue
            if duration_filter == "От 60 секунд (1 мин)" and dur_sec < 60:
                continue
            if duration_filter == "Другая..." and custom_duration is not None and dur_sec < custom_duration:
                continue

            if search_number:
                if search_number not in call["number"] and search_number not in call["account"]:
                    continue
            if from_number and from_number not in call["number"]:
                continue
            if to_number and to_number not in call["account"]:
                continue

            acc = call["account"]
            displayed_account = account_mapping.get(acc, acc)
            if route:
                # Проверяем текст для фильтра по маршруту
                if route not in acc and route not in displayed_account:
                    continue

            if selected_criterion != "Без фильтра по критерию":
                mark_color = call["marks"].get(selected_criterion, None)
                if selected_color == "Зеленый" and mark_color != "green":
                    continue
                elif selected_color == "Красный" and mark_color != "red":
                    continue

            filtered.append(call)
        return filtered

    def get_period_dates(self, period):
        today = date.today()
        if period == "Все":
            return None, None
        elif period == "Сегодня":
            return today, today
        elif period == "Вчера":
            yest = today - timedelta(days=1)
            return yest, yest
        elif period == "Текущая неделя":
            monday = today - timedelta(days=today.weekday())
            return monday, today
        elif period == "Прошлая неделя":
            monday_last = (today - timedelta(days=today.weekday())) - timedelta(days=7)
            sunday_last = monday_last + timedelta(days=6)
            return monday_last, sunday_last
        elif period == "Текущий месяц":
            first_day = date(today.year, today.month, 1)
            return first_day, today
        elif period == "Прошлый месяц":
            first_day_current = date(today.year, today.month, 1)
            last_day_last_month = first_day_current - timedelta(days=1)
            first_day_last_month = date(last_day_last_month.year, last_day_last_month.month, 1)
            return first_day_last_month, last_day_last_month
        elif period == "Произвольный период":
            start = self.start_date_edit.date().toPyDate()
            end = self.end_date_edit.date().toPyDate()
            return start, end
        return None, None

    def show_calls(self, calls):
        if not calls:
            QMessageBox.information(self, "Результаты поиска", "Нет результатов")
            return
        self.update_call_table_from_config(calls, direct_list=True)

    def open_settings(self):
        dlg = SettingsDialog(self.config)
        if dlg.exec_() == QDialog.Accepted:
            self.config = load_config()

    def change_speed(self, text):
        try:
            rate = float(text.replace("x", ""))
            self.player.setPlaybackRate(rate)
        except ValueError:
            pass

    def set_player_position(self, position):
        self.player.setPosition(position)

    def on_position_changed(self, position):
        if not self.position_slider.isSliderDown():
            self.position_slider.setValue(position)
        self.update_time_label(position, self.total_duration)

    def update_time_label(self, position, duration):
        def fmt(ms):
            s = ms // 1000
            h = s // 3600
            m = (s % 3600) // 60
            s = s % 60
            return f"{h:02d}:{m:02d}:{s:02d}"
        current_str = fmt(position)
        total_str = fmt(duration)
        self.time_label.setText(f"{current_str} / {total_str}")

    def on_state_changed(self, state):
        if state == QMediaPlayer.StoppedState:
            self.clear_current_playing_highlight()

    def toggle_play_pause(self):
        if self.player.state() == QMediaPlayer.PlayingState:
            self.player.pause()
        else:
            self.player.play()

    def open_folder_from_table(self, row, column):
        folder_name = self.folder_table.item(row, 0).text()
        self.show_custom_blocker("Открытие папки...")
        self.open_folder_by_name(folder_name)

    def open_folder_by_name(self, folder_name):
        folder_data = self.config["folder_info"].get(folder_name)
        if not folder_data:
            QMessageBox.warning(self, "Ошибка", f"Нет данных о папке {folder_name}.")
            self.hide_custom_blocker()
            return
        calls = folder_data.get("calls", [])
        self.update_call_table_from_config(calls)
        self.previous_paths.append(self.current_path)
        self.current_path = f"{self.current_path}/{folder_name}"
        self.back_button.setEnabled(True)
        self.hide_custom_blocker()

    def create_mark_widget(self, call, criterion):
        """
        Виджет с двумя кнопками (зеленая/красная) для «отметок» по критерию.
        """
        widget = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        green_btn = QPushButton("🟢")
        red_btn = QPushButton("🔴")
        green_btn.setFixedSize(70, 20)
        red_btn.setFixedSize(70, 20)

        def set_green():
            call["marks"][criterion] = "green"
            save_config(self.config)
            update_buttons()

        def set_red():
            call["marks"][criterion] = "red"
            save_config(self.config)
            update_buttons()

        def update_buttons():
            state = call["marks"].get(criterion)
            green_btn.setStyleSheet("")
            red_btn.setStyleSheet("")
            if state == "green":
                green_btn.setStyleSheet("background-color: lightgreen;")
            elif state == "red":
                red_btn.setStyleSheet("background-color: lightcoral;")

        green_btn.clicked.connect(set_green)
        red_btn.clicked.connect(set_red)
        layout.addWidget(green_btn)
        layout.addWidget(red_btn)
        widget.setLayout(layout)

        if criterion in call["marks"]:
            state = call["marks"][criterion]
            if state == "green":
                green_btn.setStyleSheet("background-color: lightgreen;")
            elif state == "red":
                red_btn.setStyleSheet("background-color: lightcoral;")

        return widget

    def update_call_table_from_config(self, calls, direct_list=False):
        """
        Отображает звонки в self.call_table.
        """
        self.call_table.setRowCount(0)
        highlight_threshold = self.config.get("highlight_threshold", 40)
        account_mapping = self.config.get("account_mapping", {})

        # Сохраняем текущий список звонков, чтобы при экспорте в XLSX можно было выгрузить именно то, что видим.
        self.current_displayed_calls = calls[:]

        base_columns = 8
        for row_position, call in enumerate(calls):
            self.call_table.insertRow(row_position)
            item_num = QTableWidgetItem(str(row_position + 1))
            duration_seconds = self.duration_to_seconds(call.get("duration", "Неизвестно"))
            if duration_seconds is not None:
                if duration_seconds > highlight_threshold:
                    item_num.setBackground(Qt.green)
                elif duration_seconds < 10:
                    item_num.setBackground(Qt.red)
            self.call_table.setItem(row_position, 0, item_num)

            self.call_table.setItem(row_position, 1, QTableWidgetItem(call["type"]))

            # Отображаем подмену по account_mapping
            displayed_number = account_mapping.get(call["number"], call["number"])
            displayed_account = account_mapping.get(call["account"], call["account"])

            self.call_table.setItem(row_position, 2, QTableWidgetItem(displayed_number))
            self.call_table.setItem(row_position, 3, QTableWidgetItem(displayed_account))

            duration = call["duration"] if call["duration"] else "Идет загрузка..."
            self.call_table.setItem(row_position, 4, QTableWidgetItem(duration))
            self.call_table.setItem(row_position, 5, QTableWidgetItem(call["datetime"]))

            play_button = QPushButton("▶")
            play_button.clicked.connect(lambda checked, r=row_position, c=calls: self.play_call(r, c))
            self.call_table.setCellWidget(row_position, 6, play_button)

            download_button = QPushButton("Скачать")
            download_button.clicked.connect(lambda checked, r=row_position, c=calls: self.download_call(r, c))
            self.call_table.setCellWidget(row_position, 7, download_button)

            if "marks" not in call:
                call["marks"] = {}

            for i, crit in enumerate(CRITERIA):
                column_index = base_columns + i
                widget = self.create_mark_widget(call, crit)
                self.call_table.setCellWidget(row_position, column_index, widget)

        self.folder_table.hide()
        self.call_table.show()

    def display_folder_info_from_table(self, row, column):
        folder_name = self.folder_table.item(row, 0).text()
        info = self.config["folder_info"].get(folder_name, {})
        day = info.get("day", "Неизвестно")
        incoming = info.get("incoming", 0)
        outgoing = info.get("outgoing", 0)
        total_calls = incoming + outgoing
        self.info_label_folder.setText(
            f"Дата: {folder_name} | День недели: {day} | Входящих: {incoming} | Исходящих: {outgoing} | Всего: {total_calls}"
        )

    def update_folder_table_from_config(self):
        folder_info = self.config.get("folder_info", {})
        folders = sorted(folder_info.keys(), key=lambda x: datetime.strptime(x, "%Y-%m-%d"), reverse=True)
        self.folder_table.setRowCount(0)
        for i, folder in enumerate(folders):
            info = folder_info[folder]
            day = info.get("day", "Неизвестно")
            incoming = info.get("incoming", 0)
            outgoing = info.get("outgoing", 0)
            total_calls = incoming + outgoing
            self.folder_table.insertRow(i)
            self.folder_table.setItem(i, 0, QTableWidgetItem(folder))
            self.folder_table.setItem(i, 1, QTableWidgetItem(day))
            self.folder_table.setItem(i, 2, QTableWidgetItem(str(incoming)))
            self.folder_table.setItem(i, 3, QTableWidgetItem(str(outgoing)))
            self.folder_table.setItem(i, 4, QTableWidgetItem(str(total_calls)))

        self.call_table.hide()
        self.folder_table.show()

    def duration_to_seconds(self, duration_str):
        if duration_str in ("Неизвестно", "Ошибка"):
            return None
        parts = duration_str.split(":")
        if len(parts) == 3:
            try:
                h, m, s = map(int, parts)
                return h * 3600 + m * 60 + s
            except:
                return None
        return None

    def play_call(self, row, calls):
        if self.player.state() == QMediaPlayer.PlayingState:
            self.player.stop()
        self.clear_current_playing_highlight()

        if calls and 0 <= row < len(calls):
            filename = calls[row]["filename"]
            local_filename = self.config.get("downloads", {}).get(filename)
            if not local_filename or not os.path.exists(local_filename):
                QMessageBox.warning(self, "Ошибка", "Файл не найден локально.")
                return
            try:
                file_url = QUrl.fromLocalFile(local_filename)
                self.player.setMedia(QMediaContent(file_url))
                self.player.play()
                self.highlight_current_playing_button(row)
            except Exception as e:
                QMessageBox.warning(self, "Ошибка воспроизведения", f"Не удалось воспроизвести файл: {e}")

    def highlight_current_playing_button(self, row):
        play_widget = self.call_table.cellWidget(row, 6)
        if isinstance(play_widget, QPushButton):
            play_widget.setStyleSheet("background-color: lightgreen;")
        self.current_playing_row = row
        self.current_playing_button = play_widget

    def clear_current_playing_highlight(self):
        if self.current_playing_button is not None:
            self.current_playing_button.setStyleSheet("")
        self.current_playing_button = None
        self.current_playing_row = None

    def download_call(self, row, calls):
        if calls and 0 <= row < len(calls):
            filename = calls[row]["filename"]
            local_filename = self.config.get("downloads", {}).get(filename)
            if not local_filename or not os.path.exists(local_filename):
                QMessageBox.warning(self, "Ошибка", "Файл не найден локально.")
                return

            initial_dir = self.config.get("download_path", "Загрузки")
            directory = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения", initial_dir)
            if not directory:
                return
            new_name, ok = QInputDialog.getText(self, "Имя файла", "Введите имя файла без расширения:",
                                                QLineEdit.Normal, filename)
            if not ok or not new_name.strip():
                return
            if not new_name.lower().endswith(".mp3"):
                new_name += ".mp3"

            target_path = os.path.join(directory, new_name)
            self.show_custom_blocker("Сохранение файла...")

            try:
                from shutil import copyfile
                copyfile(local_filename, target_path)
                QMessageBox.information(self, "Сохранение", f"Файл успешно сохранен в:\n{target_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось сохранить файл: {e}")
            finally:
                self.hide_custom_blocker()

    def go_back(self):
        if self.previous_paths:
            self.current_path = self.previous_paths.pop()
            parts = self.current_path.strip("/").split("/")
            if len(parts) > 1:
                folder_name = parts[-1]
                if folder_name in self.config.get("folder_info", {}):
                    calls = self.config["folder_info"][folder_name].get("calls", [])
                    self.update_call_table_from_config(calls)
            else:
                self.update_folder_table_from_config()
                self.call_table.hide()
                self.folder_table.show()
            if not self.previous_paths:
                self.back_button.setEnabled(False)

    def export_to_xlsx(self):
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить как...",
            self.config.get("download_path", ""),
            "Excel Files (*.xlsx)"
        )
        if not filename:
            return
        if not hasattr(self, 'current_displayed_calls') or not self.current_displayed_calls:
            QMessageBox.information(self, "Экспорт", "Нет данных для выгрузки.")
            return
        filtered = self.current_displayed_calls
        if not filtered:
            QMessageBox.information(self, "Экспорт", "Нет данных для выгрузки.")
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Звонки"
        headers = ["№", "Тип звонка", "Номер", "Аккаунт", "Длит. звонка", "Время"] + CRITERIA
        ws.append(headers)

        for i, call in enumerate(filtered, start=1):
            row = [
                i, call["type"], call["number"], call["account"],
                call["duration"], call["datetime"]
            ]
            for crit in CRITERIA:
                mark = call["marks"].get(crit, "")
                if mark == "green":
                    val = "Зеленый"
                elif mark == "red":
                    val = "Красный"
                else:
                    val = ""
                row.append(val)
            ws.append(row)

        # Автоматическая подстройка ширины
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                val = cell.value
                if val and isinstance(val, str):
                    length = len(val)
                    if length > max_length:
                        max_length = length
            ws.column_dimensions[column].width = max_length + 2

        try:
            wb.save(filename)
            QMessageBox.information(self, "Экспорт", f"Данные успешно сохранены в {filename}")
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось сохранить файл: {e}")

    def show_custom_blocker(self, text="Подождите, идет загрузка..."):
        self.blocker_dialog = QDialog(self, Qt.FramelessWindowHint)
        self.blocker_dialog.setModal(True)
        self.blocker_dialog.setWindowModality(Qt.WindowModal)
        self.blocker_dialog.setAttribute(Qt.WA_TranslucentBackground)
        self.blocker_dialog.setAttribute(Qt.WA_DeleteOnClose, False)

        layout = QVBoxLayout(self.blocker_dialog)
        label_text = QLabel(text, self.blocker_dialog)
        label_text.setAlignment(Qt.AlignCenter)
        label_gif = QLabel(self.blocker_dialog)
        label_gif.setAlignment(Qt.AlignCenter)

        base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        gif_path = os.path.join(base_path, "baldi-clyde.gif")

        movie = QMovie(gif_path)
        label_gif.setMovie(movie)
        movie.start()

        layout.addWidget(label_text)
        layout.addWidget(label_gif)
        self.blocker_dialog.setLayout(layout)
        self.blocker_dialog.setFixedSize(300, 200)
        self.blocker_dialog.show()

    def hide_custom_blocker(self):
        if hasattr(self, 'blocker_dialog') and self.blocker_dialog is not None:
            self.blocker_dialog.close()
            self.blocker_dialog = None

    def closeEvent(self, event):
        write_log("Application is closing gracefully.")
        super().closeEvent(event)


if __name__ == "__main__":
    if sys.platform == "win32":
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

    try:
        app = QApplication(sys.argv)
        app.setStyleSheet(qdarkstyle.load_stylesheet())
        window = FTPApp()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        write_log(f"{datetime.now().isoformat()} - Critical error on startup: {e}\n")
        raise
