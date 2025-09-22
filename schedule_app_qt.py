import sys
import json
import os
import random
import shutil
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QSortFilterProxyModel, QModelIndex
from PyQt5.QtGui import QFont, QIcon, QColor, QPalette, QStandardItemModel, QStandardItem
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, QLabel, QPushButton,
    QComboBox, QSpinBox, QProgressBar, QStatusBar, QFileDialog, QMessageBox,
    QDialog, QLineEdit, QFormLayout, QDialogButtonBox, QGroupBox, QScrollArea,
    QTreeView, QTableView, QAbstractItemView, QMenu, QAction, QDateEdit, QCalendarWidget,  # <-- ДОБАВЛЕНО QTableView
    QFrame, QSizePolicy, QCheckBox, QSpinBox, QDoubleSpinBox, QMenuBar, QToolBar,
    QTextEdit
)
# Импортируем необходимые модули для работы с датами и временем
import calendar
from datetime import datetime as dt_datetime

# Глобальные стили и цвета
COLORS = {
    'primary': '#4a6fa5',
    'secondary': '#166088',
    'accent': '#47b8e0',
    'success': '#3cba92',
    'warning': '#f4a261',
    'danger': '#e71d36',
    'light': '#f8f9fa',
    'dark': '#343a40',
    'border': '#dee2e6',
    'hover': '#e9ecef'
}

class BellScheduleEditor(QDialog):
    """Класс для редактирования расписания звонков (Qt версия)"""
    def __init__(self, parent, current_schedule):
        super().__init__(parent)
        self.setWindowTitle("Редактор расписания звонков")
        self.setModal(True)
        self.resize(500, 400)
        self.current_schedule = current_schedule
        self.result = None
        self.setup_ui()
        self.load_schedule_from_string(current_schedule)
    def setup_ui(self):
        layout = QVBoxLayout(self)
        # Заголовок
        title_label = QLabel("Уроки:")
        title_label.setFont(QFont('Segoe UI', 10, QFont.Bold))
        layout.addWidget(title_label)
        # Таблица
        self.tree = QTableWidget(0, 3)
        self.tree.setHorizontalHeaderLabels(['№', 'Начало', 'Конец'])
        self.tree.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tree.setSelectionMode(QAbstractItemView.SingleSelection)
        layout.addWidget(self.tree, 1)
        # Кнопки
        buttons_layout = QHBoxLayout()
        self.add_btn = QPushButton("➕ Добавить")
        self.edit_btn = QPushButton("✏️ Редактировать")
        self.delete_btn = QPushButton("🗑️ Удалить")
        self.save_btn = QPushButton("💾 Сохранить")
        self.edit_btn.setEnabled(False)
        self.delete_btn.setEnabled(False)
        self.add_btn.clicked.connect(self.add_interval)
        self.edit_btn.clicked.connect(self.edit_interval)
        self.delete_btn.clicked.connect(self.delete_interval)
        self.save_btn.clicked.connect(self.save_and_close)
        self.tree.itemSelectionChanged.connect(self.on_select)
        buttons_layout.addWidget(self.add_btn)
        buttons_layout.addWidget(self.edit_btn)
        buttons_layout.addWidget(self.delete_btn)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.save_btn)
        layout.addLayout(buttons_layout)
    def load_schedule_from_string(self, schedule_str):
        if not schedule_str.strip():
            return
        intervals = schedule_str.split(',')
        for i, interval in enumerate(intervals):
            parts = interval.strip().split('-')
            if len(parts) == 2:
                start_time, end_time = parts[0].strip(), parts[1].strip()
                row = self.tree.rowCount()
                self.tree.insertRow(row)
                self.tree.setItem(row, 0, QTableWidgetItem(str(i + 1)))
                self.tree.setItem(row, 1, QTableWidgetItem(start_time))
                self.tree.setItem(row, 2, QTableWidgetItem(end_time))
    def add_interval(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавить интервал")
        dialog.setModal(True)
        dialog.resize(300, 150)
        form_layout = QFormLayout(dialog)
        start_var = QLineEdit("8:00")
        end_var = QLineEdit("8:45")
        form_layout.addRow("Начало (ч:мм):", start_var)
        form_layout.addRow("Конец (ч:мм):", end_var)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._save_add_interval(start_var.text(), end_var.text(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _save_add_interval(self, start_time, end_time, dialog):
        if not start_time or not end_time:
            QMessageBox.warning(self, "Предупреждение", "Введите время начала и конца")
            return
        row = self.tree.rowCount()
        self.tree.insertRow(row)
        self.tree.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        self.tree.setItem(row, 1, QTableWidgetItem(start_time))
        self.tree.setItem(row, 2, QTableWidgetItem(end_time))
        dialog.accept()
    def edit_interval(self):
        selected_items = self.tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите интервал для редактирования")
            return
        row = selected_items[0].row()
        start_time = self.tree.item(row, 1).text()
        end_time = self.tree.item(row, 2).text()
        dialog = QDialog(self)
        dialog.setWindowTitle("Редактировать интервал")
        dialog.setModal(True)
        dialog.resize(300, 150)
        form_layout = QFormLayout(dialog)
        start_var = QLineEdit(start_time)
        end_var = QLineEdit(end_time)
        form_layout.addRow("Начало (ч:мм):", start_var)
        form_layout.addRow("Конец (ч:мм):", end_var)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._save_edit_interval(row, start_var.text(), end_var.text(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _save_edit_interval(self, row, start_time, end_time, dialog):
        if not start_time or not end_time:
            QMessageBox.warning(self, "Предупреждение", "Введите время начала и конца")
            return
        self.tree.setItem(row, 1, QTableWidgetItem(start_time))
        self.tree.setItem(row, 2, QTableWidgetItem(end_time))
        dialog.accept()
    def delete_interval(self):
        selected_items = self.tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите интервал для удаления")
            return
        if QMessageBox.question(self, "Подтверждение", "Вы уверены, что хотите удалить этот интервал?") == QMessageBox.Yes:
            row = selected_items[0].row()
            self.tree.removeRow(row)
            self.renumber_intervals()
    def renumber_intervals(self):
        for row in range(self.tree.rowCount()):
            self.tree.setItem(row, 0, QTableWidgetItem(str(row + 1)))
    def on_select(self):
        selected = self.tree.selectedItems()
        has_selection = len(selected) > 0
        self.edit_btn.setEnabled(has_selection)
        self.delete_btn.setEnabled(has_selection)
    def save_and_close(self):
        intervals = []
        for row in range(self.tree.rowCount()):
            start_time = self.tree.item(row, 1).text()
            end_time = self.tree.item(row, 2).text()
            intervals.append(f"{start_time}-{end_time}")
        self.result = ','.join(intervals)
        self.accept()

class TimeSortProxyModel(QSortFilterProxyModel):
    """Класс-наследник QSortFilterProxyModel для правильной сортировки по времени."""
    
    def lessThan(self, left, right):
        """
        Переопределение метода lessThan для сравнения временных интервалов.
        Сравнивает начало каждого интервала.
        """
        # Получаем значение ячейки по индексу
        left_data = self.sourceModel().data(left)
        right_data = self.sourceModel().data(right)
        
        # Преобразуем строку "HH:MM-HH:MM" в время начала
        def parse_time(time_str):
            start_time_str = time_str.split('-')[0].strip()
            return dt_datetime.strptime(start_time_str, '%H:%M')
        
        try:
            left_start = parse_time(left_data)
            right_start = parse_time(right_data)
            return left_start < right_start
        except ValueError:
            # Если парсинг не удался, используем стандартную лексикографическую сортировку
            return left_data < right_data

class ScheduleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🎓 Система автоматического составления расписания")
        self.setGeometry(100, 100, 1400, 900)
        self.setMinimumSize(1200, 800)
        # Данные приложения
        self.settings = {
            'days_per_week': 5,
            'lessons_per_day': 6,
            'weeks': 2,
            'start_date': datetime.now().date().isoformat(),
            'school_name': 'Образовательное учреждение',
            'director': 'Директор',
            'academic_year': f"{datetime.now().year}-{datetime.now().year + 1}",
            'auto_backup': True,
            'backup_interval': 30,
            'max_backups': 10,
            'last_academic_year_update': datetime.now().year
        }
        self.groups = []
        self.teachers = []
        self.classrooms = []
        self.subjects = []
        self.schedule = pd.DataFrame()
        self.substitutions = []
        self.holidays = []
        self.backup_timer = None
        self.last_backup_time = None
        self.next_backup_time = None
        # Создание директории для бэкапов
        self.backup_dir = "backups"
        if not os.path.exists(self.backup_dir):
            os.makedirs(self.backup_dir)
        # Создание директории для архива расписаний
        self.archive_dir = "schedule_archive"
        if not os.path.exists(self.archive_dir):
            os.makedirs(self.archive_dir)
        self.create_widgets()
        self.load_data()
        self.start_auto_backup()
        self.check_and_update_experience()
    def create_widgets(self):
        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        # Меню
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Файл")
        save_action = QAction("💾 Сохранить", self)
        save_action.triggered.connect(self.save_data)
        load_action = QAction("📂 Загрузить", self)
        load_action.triggered.connect(self.load_data)
        settings_action = QAction("⚙️ Настройки", self)
        settings_action.triggered.connect(self.open_settings)
        backup_action = QAction("🛡️ Бэкап", self)
        backup_action.triggered.connect(self.open_backup_manager)
        about_action = QAction("❓ О программе", self)
        about_action.triggered.connect(self.show_about)
        file_menu.addAction(save_action)
        file_menu.addAction(load_action)
        file_menu.addAction(settings_action)
        file_menu.addAction(backup_action)
        file_menu.addAction(about_action)
        # Заголовок
        title_frame = QFrame()
        title_layout = QHBoxLayout(title_frame)
        title_label = QLabel("🎓 Система автоматического составления расписания")
        title_label.setFont(QFont('Segoe UI', 18, QFont.Bold))
        title_label.setStyleSheet(f"color: {COLORS['secondary']};")
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        # Индикатор авто-бэкапа
        backup_indicator_frame = QFrame()
        backup_indicator_layout = QVBoxLayout(backup_indicator_frame)
        self.backup_status_label = QLabel("Авто-бэкап: ВКЛ")
        self.backup_status_label.setFont(QFont('Segoe UI', 10, QFont.Bold))
        self.backup_status_label.setStyleSheet(f"color: {COLORS['success']};")
        self.backup_info_label = QLabel("Следующий: --:--")
        self.backup_info_label.setFont(QFont('Segoe UI', 10, QFont.Bold))
        backup_indicator_layout.addWidget(self.backup_status_label)
        backup_indicator_layout.addWidget(self.backup_info_label)
        title_layout.addWidget(backup_indicator_frame)
        main_layout.addWidget(title_frame)
        # Верхняя панель с настройками
        settings_frame = QGroupBox("⚙️ Настройки расписания")
        settings_layout = QGridLayout(settings_frame)
        # Параметры расписания
        days_label = QLabel("Дней в неделю:")
        self.days_var = QSpinBox()
        self.days_var.setRange(1, 7)
        self.days_var.setValue(self.settings['days_per_week'])
        lessons_label = QLabel("Занятий в день:")
        self.lessons_var = QSpinBox()
        self.lessons_var.setRange(1, 12)
        self.lessons_var.setValue(self.settings['lessons_per_day'])
        weeks_label = QLabel("Недель:")
        self.weeks_var = QSpinBox()
        self.weeks_var.setRange(1, 12)
        self.weeks_var.setValue(self.settings['weeks'])
        settings_layout.addWidget(days_label, 0, 0)
        settings_layout.addWidget(self.days_var, 0, 1)
        settings_layout.addWidget(lessons_label, 0, 2)
        settings_layout.addWidget(self.lessons_var, 0, 3)
        settings_layout.addWidget(weeks_label, 0, 4)
        settings_layout.addWidget(self.weeks_var, 0, 5)
        # Кнопки управления
        buttons_frame = QFrame()
        buttons_layout = QHBoxLayout(buttons_frame)
        generate_btn = QPushButton("🚀 Сгенерировать расписание")
        generate_btn.clicked.connect(self.generate_schedule_thread)
        check_btn = QPushButton("🔍 Проверить конфликт")
        check_btn.clicked.connect(self.check_conflicts)
        optimize_btn = QPushButton("⚡ Оптимизировать")
        optimize_btn.clicked.connect(self.optimize_schedule)
        reports_btn = QPushButton("📊 Отчеты")
        reports_btn.clicked.connect(self.show_reports)
        export_excel_btn = QPushButton("📤 Экспорт в Excel")
        export_excel_btn.clicked.connect(self.export_to_excel)
        export_html_btn = QPushButton("🌐 Экспорт в HTML")
        export_html_btn.clicked.connect(self.export_to_html)
        substitutions_btn = QPushButton("🔄 Журнал замен")
        substitutions_btn.clicked.connect(self.open_substitutions)
        buttons_layout.addWidget(generate_btn)
        buttons_layout.addWidget(check_btn)
        buttons_layout.addWidget(optimize_btn)
        buttons_layout.addWidget(reports_btn)
        buttons_layout.addWidget(export_excel_btn)
        buttons_layout.addWidget(export_html_btn)
        buttons_layout.addWidget(substitutions_btn)
        buttons_layout.addStretch()
        settings_layout.addWidget(buttons_frame, 1, 0, 1, 6)
        # Прогресс-бар
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)  # Indeterminate
        self.progress.hide()
        settings_layout.addWidget(self.progress, 2, 0, 1, 6)
        main_layout.addWidget(settings_frame)
        # Основная область с вкладками
        self.notebook = QTabWidget()
        main_layout.addWidget(self.notebook, 1)
        # Создание вкладок
        self.create_groups_tab()
        self.create_teachers_tab()
        self.create_classrooms_tab()
        self.create_subjects_tab()
        self.create_schedule_tab()
        self.create_reports_tab()
        self.create_holidays_tab()
        self.create_archive_tab()
        # Статусная строка
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.status_var = "Готов к работе"
        self.statusBar.showMessage(self.status_var)
        # Обновляем индикатор бэкапа
        self.update_backup_indicator()
    def create_groups_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        # Кнопки
        btn_frame = QFrame()
        btn_layout = QHBoxLayout(btn_frame)
        add_btn = QPushButton("➕ Добавить группу")
        add_btn.clicked.connect(self.add_group)
        edit_btn = QPushButton("✏️ Редактировать")
        edit_btn.clicked.connect(self.edit_group)
        delete_btn = QPushButton("🗑️ Удалить группу")
        delete_btn.clicked.connect(self.delete_group)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        layout.addWidget(btn_frame)
        # Таблица
        self.groups_tree = QTableWidget(0, 6)
        self.groups_tree.setHorizontalHeaderLabels(['ID', 'Название', 'Тип', 'Студенты', 'Курс', 'Специальность'])
        self.groups_tree.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.groups_tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.groups_tree, 1)
        self.notebook.addTab(tab, "👥 Группы")
    def create_teachers_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        # Кнопки
        btn_frame = QFrame()
        btn_layout = QHBoxLayout(btn_frame)
        add_btn = QPushButton("➕ Добавить преподавателя")
        add_btn.clicked.connect(self.add_teacher)
        edit_btn = QPushButton("✏️ Редактировать")
        edit_btn.clicked.connect(self.edit_teacher)
        delete_btn = QPushButton("🗑️ Удалить преподавателя")
        delete_btn.clicked.connect(self.delete_teacher)
        update_exp_btn = QPushButton("🔄 Обновить стаж")
        update_exp_btn.clicked.connect(self.update_all_experience)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addWidget(update_exp_btn)
        btn_layout.addStretch()
        layout.addWidget(btn_frame)
        # Таблица
        self.teachers_tree = QTableWidget(0, 6)
        self.teachers_tree.setHorizontalHeaderLabels(['ID', 'ФИО', 'Предметы', 'Макс. часов', 'Квалификация', 'Стаж'])
        self.teachers_tree.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.teachers_tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.teachers_tree, 1)
        self.notebook.addTab(tab, "👨‍🏫 Преподаватели")
    def create_classrooms_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        # Кнопки
        btn_frame = QFrame()
        btn_layout = QHBoxLayout(btn_frame)
        add_btn = QPushButton("➕ Добавить аудиторию")
        add_btn.clicked.connect(self.add_classroom)
        edit_btn = QPushButton("✏️ Редактировать")
        edit_btn.clicked.connect(self.edit_classroom)
        delete_btn = QPushButton("🗑️ Удалить аудиторию")
        delete_btn.clicked.connect(self.delete_classroom)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        layout.addWidget(btn_frame)
        # Таблица
        self.classrooms_tree = QTableWidget(0, 6)
        self.classrooms_tree.setHorizontalHeaderLabels(['ID', 'Номер', 'Вместимость', 'Тип', 'Оборудование', 'Расположение'])
        self.classrooms_tree.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.classrooms_tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.classrooms_tree, 1)
        self.notebook.addTab(tab, "🏫 Аудитории")
    def create_subjects_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        # Кнопки
        btn_frame = QFrame()
        btn_layout = QHBoxLayout(btn_frame)
        add_btn = QPushButton("➕ Добавить предмет")
        add_btn.clicked.connect(self.add_subject)
        edit_btn = QPushButton("✏️ Редактировать")
        edit_btn.clicked.connect(self.edit_subject)
        delete_btn = QPushButton("🗑️ Удалить предмет")
        delete_btn.clicked.connect(self.delete_subject)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        layout.addWidget(btn_frame)
        # Таблица
        self.subjects_tree = QTableWidget(0, 6)
        self.subjects_tree.setHorizontalHeaderLabels(['ID', 'Название', 'Тип группы', 'Часов/неделю', 'Форма контроля', 'Кафедра'])
        self.subjects_tree.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.subjects_tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.subjects_tree, 1)
        self.notebook.addTab(tab, "📚 Предметы")
    def create_schedule_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        # Фрейм для фильтров
        filter_frame = QFrame()
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.addWidget(QLabel("Фильтры:"))
        # Неделя
        week_layout = QHBoxLayout()
        week_layout.addWidget(QLabel("Неделя:"))
        self.week_var = QComboBox()
        self.week_var.addItems([f"Неделя {i}" for i in range(1, 13)])
        self.week_var.setCurrentIndex(0)
        self.week_var.currentIndexChanged.connect(self.filter_schedule)
        week_layout.addWidget(self.week_var)
        filter_layout.addLayout(week_layout)
        # Группа
        group_layout = QHBoxLayout()
        group_layout.addWidget(QLabel("Группа:"))
        self.group_filter_var = QComboBox()
        self.group_filter_var.addItem("")
        group_layout.addWidget(self.group_filter_var)
        filter_layout.addLayout(group_layout)
        # Преподаватель
        teacher_layout = QHBoxLayout()
        teacher_layout.addWidget(QLabel("Преподаватель:"))
        self.teacher_filter_var = QComboBox()
        self.teacher_filter_var.addItem("")
        teacher_layout.addWidget(self.teacher_filter_var)
        filter_layout.addLayout(teacher_layout)
        # Аудитория
        classroom_layout = QHBoxLayout()
        classroom_layout.addWidget(QLabel("Аудитория:"))
        self.classroom_filter_var = QComboBox()
        self.classroom_filter_var.addItem("")
        classroom_layout.addWidget(self.classroom_filter_var)
        filter_layout.addLayout(classroom_layout)
        # Кнопка обновления
        refresh_btn = QPushButton("🔄 Обновить")
        refresh_btn.clicked.connect(self.filter_schedule)
        filter_layout.addWidget(refresh_btn)
        filter_layout.addStretch()
        layout.addWidget(filter_frame)
        # Кнопки управления
        schedule_buttons_frame = QFrame()
        buttons_layout = QHBoxLayout(schedule_buttons_frame)
        buttons = [
            ("➕ Добавить занятие", self.add_lesson),
            ("✏️ Редактировать", self.edit_lesson),
            ("🗑️ Удалить занятие", self.delete_lesson),
            ("🔄 Заменить занятие", self.substitute_lesson),
            ("📅 Календарь", self.show_calendar),
            ("🌐 Экспорт в HTML", self.export_to_html),
            ("⏱️ Найти свободное время", self.find_free_slot)
        ]
        for text, command in buttons:
            btn = QPushButton(text)
            btn.clicked.connect(command)
            buttons_layout.addWidget(btn)
        buttons_layout.addStretch()
        layout.addWidget(schedule_buttons_frame)
        # Таблица расписания
        # Создаем модель данных
        self.schedule_model = QStandardItemModel()
        # Устанавливаем заголовки столбцов
        self.schedule_model.setHorizontalHeaderLabels(['Время', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'])
        
        # Создаем прокси-модель для сортировки
        self.schedule_proxy_model = TimeSortProxyModel(self)
        self.schedule_proxy_model.setSourceModel(self.schedule_model)
        
        # Создаем виджет таблицы
        self.schedule_view = QTableView()
        self.schedule_view.setModel(self.schedule_proxy_model)
        # Устанавливаем ширину столбца "Время"
        self.schedule_view.setColumnWidth(0, 100)
        # Увеличиваем высоту строк
        self.schedule_view.verticalHeader().setDefaultSectionSize(100)
        # Разрешаем выбор строк
        self.schedule_view.setSelectionBehavior(QAbstractItemView.SelectItems)
        # Показываем только первый столбец (время) в вертикальном заголовке
        # Разрешаем растягивание столбцов
        self.schedule_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.schedule_view, 1)
        self.notebook.addTab(tab, "📅 Расписание")
    def create_reports_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        # Вкладки отчетов
        reports_notebook = QTabWidget()
        layout.addWidget(reports_notebook, 1)
        # Отчет по нагрузке преподавателей
        teacher_report_frame = QWidget()
        teacher_report_layout = QVBoxLayout(teacher_report_frame)
        self.teacher_report_tree = QTableWidget(0, 4)
        self.teacher_report_tree.setHorizontalHeaderLabels(['Преподаватель', 'Часы', 'Группы', 'Предметы'])
        self.teacher_report_tree.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        teacher_report_layout.addWidget(self.teacher_report_tree, 1)
        reports_notebook.addTab(teacher_report_frame, "👨‍🏫 Нагрузка преподавателей")
        # Отчет по нагрузке групп
        group_report_frame = QWidget()
        group_report_layout = QVBoxLayout(group_report_frame)
        self.group_report_tree = QTableWidget(0, 4)
        self.group_report_tree.setHorizontalHeaderLabels(['Группа', 'Часы', 'Предметы', 'Преподаватели'])
        self.group_report_tree.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        group_report_layout.addWidget(self.group_report_tree, 1)
        reports_notebook.addTab(group_report_frame, "👥 Нагрузка групп")
        # Отчет по конфликтам
        conflicts_frame = QWidget()
        conflicts_layout = QVBoxLayout(conflicts_frame)
        self.conflicts_text = QTextEdit()
        self.conflicts_text.setReadOnly(True)
        conflicts_layout.addWidget(self.conflicts_text, 1)
        reports_notebook.addTab(conflicts_frame, "⚠️ Конфликты")
        # Текстовый отчет
        summary_frame = QWidget()
        summary_layout = QVBoxLayout(summary_frame)
        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        summary_layout.addWidget(self.summary_text, 1)
        reports_notebook.addTab(summary_frame, "📋 Сводка")
        self.notebook.addTab(tab, "📈 Отчеты")
    def create_holidays_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        # Кнопки
        btn_frame = QFrame()
        btn_layout = QHBoxLayout(btn_frame)
        add_btn = QPushButton("➕ Добавить праздник")
        add_btn.clicked.connect(self.add_holiday)
        delete_btn = QPushButton("🗑️ Удалить праздник")
        delete_btn.clicked.connect(self.delete_holiday)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        layout.addWidget(btn_frame)
        # Таблица
        self.holidays_tree = QTableWidget(0, 3)
        self.holidays_tree.setHorizontalHeaderLabels(['Дата', 'Название', 'Тип'])
        self.holidays_tree.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.holidays_tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.holidays_tree, 1)
        self.notebook.addTab(tab, "🎉 Праздники")
    def create_archive_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        # Кнопки
        btn_frame = QFrame()
        btn_layout = QHBoxLayout(btn_frame)
        save_btn = QPushButton("💾 Сохранить текущее расписание")
        save_btn.clicked.connect(self.save_current_schedule)
        load_btn = QPushButton("📂 Загрузить расписание")
        load_btn.clicked.connect(self.load_archived_schedule)
        delete_btn = QPushButton("🗑️ Удалить")
        delete_btn.clicked.connect(self.delete_archived_schedule)
        export_btn = QPushButton("📤 Экспорт в Excel")
        export_btn.clicked.connect(self.export_archived_schedule)
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(load_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(export_btn)
        layout.addWidget(btn_frame)
        # Таблица
        self.archive_tree = QTableWidget(0, 7)
        self.archive_tree.setHorizontalHeaderLabels(['Имя файла', 'Дата создания', 'Группы', 'Преподаватели', 'Аудитории', 'Предметы', 'Занятий'])
        self.archive_tree.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.archive_tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.archive_tree, 1)
        self.notebook.addTab(tab, "💾 Архив расписаний")
    # --- Загрузка данных в таблицы ---
    def load_groups_data(self):
        self.groups_tree.setRowCount(0)
        for group in self.groups:
            row = self.groups_tree.rowCount()
            self.groups_tree.insertRow(row)
            self.groups_tree.setItem(row, 0, QTableWidgetItem(str(group['id'])))
            self.groups_tree.setItem(row, 1, QTableWidgetItem(group['name']))
            self.groups_tree.setItem(row, 2, QTableWidgetItem(group.get('type', 'основное')))
            self.groups_tree.setItem(row, 3, QTableWidgetItem(str(group.get('students', 0))))
            self.groups_tree.setItem(row, 4, QTableWidgetItem(group.get('course', '')))
            self.groups_tree.setItem(row, 5, QTableWidgetItem(group.get('specialty', '')))
        # Обновляем комбобокс фильтра групп
        self.group_filter_var.clear()
        self.group_filter_var.addItem("")
        for group in self.groups:
            self.group_filter_var.addItem(group['name'])
    def load_teachers_data(self):
        self.teachers_tree.setRowCount(0)
        for teacher in self.teachers:
            row = self.teachers_tree.rowCount()
            self.teachers_tree.insertRow(row)
            self.teachers_tree.setItem(row, 0, QTableWidgetItem(str(teacher['id'])))
            self.teachers_tree.setItem(row, 1, QTableWidgetItem(teacher['name']))
            self.teachers_tree.setItem(row, 2, QTableWidgetItem(teacher.get('subjects', '')))
            self.teachers_tree.setItem(row, 3, QTableWidgetItem(str(teacher.get('max_hours', 0))))
            self.teachers_tree.setItem(row, 4, QTableWidgetItem(teacher.get('qualification', '')))
            self.teachers_tree.setItem(row, 5, QTableWidgetItem(str(teacher.get('experience', 0))))
        # Обновляем комбобокс фильтра преподавателей
        self.teacher_filter_var.clear()
        self.teacher_filter_var.addItem("")
        for teacher in self.teachers:
            self.teacher_filter_var.addItem(teacher['name'])
    def load_classrooms_data(self):
        self.classrooms_tree.setRowCount(0)
        for classroom in self.classrooms:
            row = self.classrooms_tree.rowCount()
            self.classrooms_tree.insertRow(row)
            self.classrooms_tree.setItem(row, 0, QTableWidgetItem(str(classroom['id'])))
            self.classrooms_tree.setItem(row, 1, QTableWidgetItem(classroom['name']))
            self.classrooms_tree.setItem(row, 2, QTableWidgetItem(str(classroom.get('capacity', 0))))
            self.classrooms_tree.setItem(row, 3, QTableWidgetItem(classroom.get('type', 'обычная')))
            self.classrooms_tree.setItem(row, 4, QTableWidgetItem(classroom.get('equipment', '')))
            self.classrooms_tree.setItem(row, 5, QTableWidgetItem(classroom.get('location', '')))
        # Обновляем комбобокс фильтра аудиторий
        self.classroom_filter_var.clear()
        self.classroom_filter_var.addItem("")
        for classroom in self.classrooms:
            self.classroom_filter_var.addItem(classroom['name'])
    def load_subjects_data(self):
        self.subjects_tree.setRowCount(0)
        for subject in self.subjects:
            row = self.subjects_tree.rowCount()
            self.subjects_tree.insertRow(row)
            self.subjects_tree.setItem(row, 0, QTableWidgetItem(str(subject['id'])))
            self.subjects_tree.setItem(row, 1, QTableWidgetItem(subject['name']))
            self.subjects_tree.setItem(row, 2, QTableWidgetItem(subject.get('group_type', 'основное')))
            self.subjects_tree.setItem(row, 3, QTableWidgetItem(str(subject.get('hours_per_week', 0))))
            self.subjects_tree.setItem(row, 4, QTableWidgetItem(subject.get('assessment', '')))
            self.subjects_tree.setItem(row, 5, QTableWidgetItem(subject.get('department', '')))
    def load_holidays_data(self):
        self.holidays_tree.setRowCount(0)
        for holiday in self.holidays:
            row = self.holidays_tree.rowCount()
            self.holidays_tree.insertRow(row)
            self.holidays_tree.setItem(row, 0, QTableWidgetItem(holiday.get('date', '')))
            self.holidays_tree.setItem(row, 1, QTableWidgetItem(holiday.get('name', '')))
            self.holidays_tree.setItem(row, 2, QTableWidgetItem(holiday.get('type', 'Государственный')))
    # --- Методы добавления/редактирования/удаления ---
    def add_group(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавить группу")
        dialog.setModal(True)
        dialog.resize(400, 300)
        form_layout = QFormLayout(dialog)
        name_entry = QLineEdit()
        type_combo = QComboBox()
        type_combo.addItems(["основное", "углубленное", "вечернее"])
        students_spin = QSpinBox()
        students_spin.setRange(1, 100)
        students_spin.setValue(25)
        course_entry = QLineEdit()
        specialty_entry = QLineEdit()
        form_layout.addRow("Название:", name_entry)
        form_layout.addRow("Тип:", type_combo)
        form_layout.addRow("Студентов:", students_spin)
        form_layout.addRow("Курс:", course_entry)
        form_layout.addRow("Специальность:", specialty_entry)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._save_group(name_entry.text(), type_combo.currentText(), students_spin.value(), course_entry.text(), specialty_entry.text(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _save_group(self, name, group_type, students, course, specialty, dialog):
        if not name:
            QMessageBox.warning(self, "Предупреждение", "Введите название группы")
            return
        new_id = max([g['id'] for g in self.groups], default=0) + 1
        self.groups.append({
            'id': new_id,
            'name': name,
            'type': group_type,
            'students': students,
            'course': course,
            'specialty': specialty
        })
        self.load_groups_data()
        self.create_backup()
        dialog.accept()
    def edit_group(self):
        selected_items = self.groups_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите группу для редактирования")
            return
        row = selected_items[0].row()
        group_id = int(self.groups_tree.item(row, 0).text())
        group = next((g for g in self.groups if g['id'] == group_id), None)
        if not group:
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Редактировать группу")
        dialog.setModal(True)
        dialog.resize(400, 300)
        form_layout = QFormLayout(dialog)
        name_entry = QLineEdit(group['name'])
        type_combo = QComboBox()
        type_combo.addItems(["основное", "углубленное", "вечернее"])
        type_combo.setCurrentText(group.get('type', 'основное'))
        students_spin = QSpinBox()
        students_spin.setRange(1, 100)
        students_spin.setValue(group.get('students', 25))
        course_entry = QLineEdit(group.get('course', ''))
        specialty_entry = QLineEdit(group.get('specialty', ''))
        form_layout.addRow("Название:", name_entry)
        form_layout.addRow("Тип:", type_combo)
        form_layout.addRow("Студентов:", students_spin)
        form_layout.addRow("Курс:", course_entry)
        form_layout.addRow("Специальность:", specialty_entry)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._update_group(group, name_entry.text(), type_combo.currentText(), students_spin.value(), course_entry.text(), specialty_entry.text(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _update_group(self, group, name, group_type, students, course, specialty, dialog):
        if not name:
            QMessageBox.warning(self, "Предупреждение", "Введите название группы")
            return
        group['name'] = name
        group['type'] = group_type
        group['students'] = students
        group['course'] = course
        group['specialty'] = specialty
        self.load_groups_data()
        self.create_backup()
        dialog.accept()
    def delete_group(self):
        selected_items = self.groups_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите группу для удаления")
            return
        if QMessageBox.question(self, "Подтверждение", "Удалить выбранную группу?") == QMessageBox.Yes:
            row = selected_items[0].row()
            group_id = int(self.groups_tree.item(row, 0).text())
            self.groups = [g for g in self.groups if g['id'] != group_id]
            self.load_groups_data()
            self.create_backup()
    def add_teacher(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавить преподавателя")
        dialog.setModal(True)
        dialog.resize(450, 500)
        form_layout = QFormLayout(dialog)
        name_entry = QLineEdit()
        subjects_entry = QLineEdit()
        max_hours_spin = QSpinBox()
        max_hours_spin.setRange(1, 50)
        max_hours_spin.setValue(20)
        qualification_entry = QLineEdit()
        experience_spin = QSpinBox()
        experience_spin.setRange(0, 50)
        experience_spin.setValue(0)
        contacts_entry = QLineEdit()
        forbidden_days_entry = QLineEdit()
        preferred_days_entry = QLineEdit()
        max_lessons_per_day_spin = QSpinBox()
        max_lessons_per_day_spin.setRange(1, 12)
        max_lessons_per_day_spin.setValue(6)
        form_layout.addRow("ФИО:", name_entry)
        form_layout.addRow("Предметы (через запятую):", subjects_entry)
        form_layout.addRow("Макс. часов/неделю:", max_hours_spin)
        form_layout.addRow("Квалификация:", qualification_entry)
        form_layout.addRow("Стаж (лет):", experience_spin)
        form_layout.addRow("Контакты:", contacts_entry)
        form_layout.addRow("Запрещенные дни:", forbidden_days_entry)
        form_layout.addRow("Предпочтительные дни:", preferred_days_entry)
        form_layout.addRow("Макс. пар в день:", max_lessons_per_day_spin)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._save_teacher(
            name_entry.text(), subjects_entry.text(), max_hours_spin.value(),
            qualification_entry.text(), experience_spin.value(), contacts_entry.text(),
            forbidden_days_entry.text(), preferred_days_entry.text(), max_lessons_per_day_spin.value(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _save_teacher(self, name, subjects, max_hours, qualification, experience, contacts,
                     forbidden_days, preferred_days, max_lessons_per_day, dialog):
        if not name:
            QMessageBox.warning(self, "Предупреждение", "Введите ФИО преподавателя")
            return
        new_id = max([t['id'] for t in self.teachers], default=0) + 1
        self.teachers.append({
            'id': new_id,
            'name': name,
            'subjects': subjects,
            'max_hours': max_hours,
            'qualification': qualification,
            'experience': experience,
            'contacts': contacts,
            'forbidden_days': forbidden_days,
            'preferred_days': preferred_days,
            'max_lessons_per_day': max_lessons_per_day
        })
        self.load_teachers_data()
        self.create_backup()
        dialog.accept()
    def edit_teacher(self):
        selected_items = self.teachers_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите преподавателя для редактирования")
            return
        row = selected_items[0].row()
        teacher_id = int(self.teachers_tree.item(row, 0).text())
        teacher = next((t for t in self.teachers if t['id'] == teacher_id), None)
        if not teacher:
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Редактировать преподавателя")
        dialog.setModal(True)
        dialog.resize(450, 500)
        form_layout = QFormLayout(dialog)
        name_entry = QLineEdit(teacher['name'])
        subjects_entry = QLineEdit(teacher.get('subjects', ''))
        max_hours_spin = QSpinBox()
        max_hours_spin.setRange(1, 50)
        max_hours_spin.setValue(teacher.get('max_hours', 20))
        qualification_entry = QLineEdit(teacher.get('qualification', ''))
        experience_spin = QSpinBox()
        experience_spin.setRange(0, 50)
        experience_spin.setValue(teacher.get('experience', 0))
        contacts_entry = QLineEdit(teacher.get('contacts', ''))
        forbidden_days_entry = QLineEdit(teacher.get('forbidden_days', ''))
        preferred_days_entry = QLineEdit(teacher.get('preferred_days', ''))
        max_lessons_per_day_spin = QSpinBox()
        max_lessons_per_day_spin.setRange(1, 12)
        max_lessons_per_day_spin.setValue(teacher.get('max_lessons_per_day', 6))
        form_layout.addRow("ФИО:", name_entry)
        form_layout.addRow("Предметы (через запятую):", subjects_entry)
        form_layout.addRow("Макс. часов/неделю:", max_hours_spin)
        form_layout.addRow("Квалификация:", qualification_entry)
        form_layout.addRow("Стаж (лет):", experience_spin)
        form_layout.addRow("Контакты:", contacts_entry)
        form_layout.addRow("Запрещенные дни:", forbidden_days_entry)
        form_layout.addRow("Предпочтительные дни:", preferred_days_entry)
        form_layout.addRow("Макс. пар в день:", max_lessons_per_day_spin)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._update_teacher(
            teacher, name_entry.text(), subjects_entry.text(), max_hours_spin.value(),
            qualification_entry.text(), experience_spin.value(), contacts_entry.text(),
            forbidden_days_entry.text(), preferred_days_entry.text(), max_lessons_per_day_spin.value(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _update_teacher(self, teacher, name, subjects, max_hours, qualification, experience, contacts,
                       forbidden_days, preferred_days, max_lessons_per_day, dialog):
        if not name:
            QMessageBox.warning(self, "Предупреждение", "Введите ФИО преподавателя")
            return
        teacher['name'] = name
        teacher['subjects'] = subjects
        teacher['max_hours'] = max_hours
        teacher['qualification'] = qualification
        teacher['experience'] = experience
        teacher['contacts'] = contacts
        teacher['forbidden_days'] = forbidden_days
        teacher['preferred_days'] = preferred_days
        teacher['max_lessons_per_day'] = max_lessons_per_day
        self.load_teachers_data()
        self.create_backup()
        dialog.accept()
    def delete_teacher(self):
        selected_items = self.teachers_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите преподавателя для удаления")
            return
        if QMessageBox.question(self, "Подтверждение", "Удалить выбранного преподавателя?") == QMessageBox.Yes:
            row = selected_items[0].row()
            teacher_id = int(self.teachers_tree.item(row, 0).text())
            self.teachers = [t for t in self.teachers if t['id'] != teacher_id]
            self.load_teachers_data()
            self.create_backup()
    def update_all_experience(self):
        if QMessageBox.question(self, "Подтверждение", "Вы уверены, что хотите увеличить стаж всех преподавателей на 1 год?") == QMessageBox.Yes:
            updated_count = 0
            for teacher in self.teachers:
                teacher['experience'] = teacher.get('experience', 0) + 1
                updated_count += 1
            self.load_teachers_data()
            self.create_backup()
            QMessageBox.information(self, "Успех", f"Стаж обновлен у {updated_count} преподавателей")
    def add_classroom(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавить аудиторию")
        dialog.setModal(True)
        dialog.resize(400, 300)
        form_layout = QFormLayout(dialog)
        name_entry = QLineEdit()
        capacity_spin = QSpinBox()
        capacity_spin.setRange(1, 200)
        capacity_spin.setValue(30)
        type_combo = QComboBox()
        type_combo.addItems(["обычная", "компьютерный класс", "спортзал", "лаборатория", "специализированная"])
        equipment_entry = QLineEdit()
        location_entry = QLineEdit()
        form_layout.addRow("Номер:", name_entry)
        form_layout.addRow("Вместимость:", capacity_spin)
        form_layout.addRow("Тип:", type_combo)
        form_layout.addRow("Оборудование:", equipment_entry)
        form_layout.addRow("Расположение:", location_entry)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._save_classroom(
            name_entry.text(), capacity_spin.value(), type_combo.currentText(),
            equipment_entry.text(), location_entry.text(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _save_classroom(self, name, capacity, room_type, equipment, location, dialog):
        if not name:
            QMessageBox.warning(self, "Предупреждение", "Введите номер аудитории")
            return
        new_id = max([c['id'] for c in self.classrooms], default=0) + 1
        self.classrooms.append({
            'id': new_id,
            'name': name,
            'capacity': capacity,
            'type': room_type,
            'equipment': equipment,
            'location': location
        })
        self.load_classrooms_data()
        self.create_backup()
        dialog.accept()
    def edit_classroom(self):
        selected_items = self.classrooms_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите аудиторию для редактирования")
            return
        row = selected_items[0].row()
        classroom_id = int(self.classrooms_tree.item(row, 0).text())
        classroom = next((c for c in self.classrooms if c['id'] == classroom_id), None)
        if not classroom:
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Редактировать аудиторию")
        dialog.setModal(True)
        dialog.resize(400, 300)
        form_layout = QFormLayout(dialog)
        name_entry = QLineEdit(classroom['name'])
        capacity_spin = QSpinBox()
        capacity_spin.setRange(1, 200)
        capacity_spin.setValue(classroom.get('capacity', 30))
        type_combo = QComboBox()
        type_combo.addItems(["обычная", "компьютерный класс", "спортзал", "лаборатория", "специализированная"])
        type_combo.setCurrentText(classroom.get('type', 'обычная'))
        equipment_entry = QLineEdit(classroom.get('equipment', ''))
        location_entry = QLineEdit(classroom.get('location', ''))
        form_layout.addRow("Номер:", name_entry)
        form_layout.addRow("Вместимость:", capacity_spin)
        form_layout.addRow("Тип:", type_combo)
        form_layout.addRow("Оборудование:", equipment_entry)
        form_layout.addRow("Расположение:", location_entry)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._update_classroom(
            classroom, name_entry.text(), capacity_spin.value(), type_combo.currentText(),
            equipment_entry.text(), location_entry.text(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _update_classroom(self, classroom, name, capacity, room_type, equipment, location, dialog):
        if not name:
            QMessageBox.warning(self, "Предупреждение", "Введите номер аудитории")
            return
        classroom['name'] = name
        classroom['capacity'] = capacity
        classroom['type'] = room_type
        classroom['equipment'] = equipment
        classroom['location'] = location
        self.load_classrooms_data()
        self.create_backup()
        dialog.accept()
    def delete_classroom(self):
        selected_items = self.classrooms_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите аудиторию для удаления")
            return
        if QMessageBox.question(self, "Подтверждение", "Удалить выбранную аудиторию?") == QMessageBox.Yes:
            row = selected_items[0].row()
            classroom_id = int(self.classrooms_tree.item(row, 0).text())
            self.classrooms = [c for c in self.classrooms if c['id'] != classroom_id]
            self.load_classrooms_data()
            self.create_backup()
    def add_subject(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавить предмет")
        dialog.setModal(True)
        dialog.resize(450, 350)
        form_layout = QFormLayout(dialog)
        name_entry = QLineEdit()
        group_type_combo = QComboBox()
        group_type_combo.addItems(["основное", "углубленное", "вечернее"])
        hours_spin = QSpinBox()
        hours_spin.setRange(1, 20)
        hours_spin.setValue(4)
        assessment_combo = QComboBox()
        assessment_combo.addItems(["экзамен", "зачет", "зачет с оценкой", "курсовая работа", "дифференцированный зачет"])
        department_entry = QLineEdit()
        description_entry = QLineEdit()
        form_layout.addRow("Название:", name_entry)
        form_layout.addRow("Тип группы:", group_type_combo)
        form_layout.addRow("Часов/неделю:", hours_spin)
        form_layout.addRow("Форма контроля:", assessment_combo)
        form_layout.addRow("Кафедра:", department_entry)
        form_layout.addRow("Описание:", description_entry)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._save_subject(
            name_entry.text(), group_type_combo.currentText(), hours_spin.value(),
            assessment_combo.currentText(), department_entry.text(), description_entry.text(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _save_subject(self, name, group_type, hours_per_week, assessment, department, description, dialog):
        if not name:
            QMessageBox.warning(self, "Предупреждение", "Введите название предмета")
            return
        new_id = max([s['id'] for s in self.subjects], default=0) + 1
        self.subjects.append({
            'id': new_id,
            'name': name,
            'group_type': group_type,
            'hours_per_week': hours_per_week,
            'assessment': assessment,
            'department': department,
            'description': description
        })
        self.load_subjects_data()
        self.create_backup()
        dialog.accept()
    def edit_subject(self):
        selected_items = self.subjects_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите предмет для редактирования")
            return
        row = selected_items[0].row()
        subject_id = int(self.subjects_tree.item(row, 0).text())
        subject = next((s for s in self.subjects if s['id'] == subject_id), None)
        if not subject:
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Редактировать предмет")
        dialog.setModal(True)
        dialog.resize(450, 350)
        form_layout = QFormLayout(dialog)
        name_entry = QLineEdit(subject['name'])
        group_type_combo = QComboBox()
        group_type_combo.addItems(["основное", "углубленное", "вечернее"])
        group_type_combo.setCurrentText(subject.get('group_type', 'основное'))
        hours_spin = QSpinBox()
        hours_spin.setRange(1, 20)
        hours_spin.setValue(subject.get('hours_per_week', 4))
        assessment_combo = QComboBox()
        assessment_combo.addItems(["экзамен", "зачет", "зачет с оценкой", "курсовая работа", "дифференцированный зачет"])
        assessment_combo.setCurrentText(subject.get('assessment', 'экзамен'))
        department_entry = QLineEdit(subject.get('department', ''))
        description_entry = QLineEdit(subject.get('description', ''))
        form_layout.addRow("Название:", name_entry)
        form_layout.addRow("Тип группы:", group_type_combo)
        form_layout.addRow("Часов/неделю:", hours_spin)
        form_layout.addRow("Форма контроля:", assessment_combo)
        form_layout.addRow("Кафедра:", department_entry)
        form_layout.addRow("Описание:", description_entry)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._update_subject(
            subject, name_entry.text(), group_type_combo.currentText(), hours_spin.value(),
            assessment_combo.currentText(), department_entry.text(), description_entry.text(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _update_subject(self, subject, name, group_type, hours_per_week, assessment, department, description, dialog):
        if not name:
            QMessageBox.warning(self, "Предупреждение", "Введите название предмета")
            return
        subject['name'] = name
        subject['group_type'] = group_type
        subject['hours_per_week'] = hours_per_week
        subject['assessment'] = assessment
        subject['department'] = department
        subject['description'] = description
        self.load_subjects_data()
        self.create_backup()
        dialog.accept()
    def delete_subject(self):
        selected_items = self.subjects_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите предмет для удаления")
            return
        if QMessageBox.question(self, "Подтверждение", "Удалить выбранный предмет?") == QMessageBox.Yes:
            row = selected_items[0].row()
            subject_id = int(self.subjects_tree.item(row, 0).text())
            self.subjects = [s for s in self.subjects if s['id'] != subject_id]
            self.load_subjects_data()
            self.create_backup()
    # --- Генерация расписания ---
    def generate_schedule_thread(self):
        self.progress.show()
        self.statusBar.showMessage("Генерация расписания...")
        # В реальном приложении здесь запускался бы отдельный поток.
        # Для простоты примера вызываем напрямую.
        self.generate_schedule()
    def generate_schedule(self):
        try:
            if not self.groups:
                raise Exception("Необходимо добавить хотя бы одну группу")
            if not self.subjects:
                raise Exception("Необходимо добавить хотя бы один предмет")
            if not self.teachers:
                raise Exception("Необходимо добавить хотя бы одного преподавателя")
            if not self.classrooms:
                raise Exception("Необходимо добавить хотя бы одну аудиторию")
            self.settings['days_per_week'] = self.days_var.value()
            self.settings['lessons_per_day'] = self.lessons_var.value()
            self.settings['weeks'] = self.weeks_var.value()
            holiday_dates = []
            for h in self.holidays:
                try:
                    holiday_dates.append(datetime.strptime(h['date'], '%Y-%m-%d').date())
                except ValueError:
                    continue
            days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
            start_date = datetime.strptime(self.settings['start_date'], '%Y-%m-%d').date()
            all_dates = []
            for week in range(self.settings['weeks']):
                for day_index, day_name in enumerate(days):
                    current_date = start_date + timedelta(weeks=week, days=day_index)
                    if current_date not in holiday_dates:
                        all_dates.append((week+1, day_name, current_date))
            schedule_data = []
            lesson_id = 1
            bell_schedule_str = self.settings.get('bell_schedule', '8:00-8:45,8:55-9:40,9:50-10:35,10:45-11:30,11:40-12:25,12:35-13:20')
            times = [slot.strip() for slot in bell_schedule_str.split(',')]
            for week_num, day_name, _ in all_dates:
                for time_slot in times:
                    for group in self.groups:
                        schedule_data.append({
                            'id': lesson_id,
                            'week': week_num,
                            'day': day_name,
                            'time': time_slot,
                            'group_id': group['id'],
                            'group_name': group['name'],
                            'subject_id': None,
                            'subject_name': '',
                            'teacher_id': None,
                            'teacher_name': '',
                            'classroom_id': None,
                            'classroom_name': '',
                            'status': 'свободно'
                        })
                        lesson_id += 1
            self.schedule = pd.DataFrame(schedule_data)
            self.assign_subjects_to_groups()
            self.assign_teachers_and_classrooms()
            self.on_schedule_generated()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка генерации расписания: {str(e)}")
            self.progress.hide()
            self.statusBar.showMessage("Ошибка генерации")
    def assign_subjects_to_groups(self):
        if not self.subjects or not self.groups:
            return
        for group in self.groups:
            group_subjects = [s for s in self.subjects if s.get('group_type') in [group['type'], 'общий']]
            for subject in group_subjects:
                hours_per_week = subject.get('hours_per_week', 0)
                if hours_per_week <= 0:
                    continue
                free_slots = self.schedule[
                    (self.schedule['group_id'] == group['id']) &
                    (self.schedule['status'] == 'свободно')
                ].index
                if len(free_slots) >= hours_per_week:
                    selected_slots = random.sample(list(free_slots), hours_per_week)
                    for slot in selected_slots:
                        self.schedule.loc[slot, 'subject_id'] = subject['id']
                        self.schedule.loc[slot, 'subject_name'] = subject['name']
                        self.schedule.loc[slot, 'status'] = 'запланировано'
    def assign_teachers_and_classrooms(self):
        for idx, lesson in self.schedule.iterrows():
            if lesson['status'] == 'запланировано':
                subject = next((s for s in self.subjects if s['id'] == lesson['subject_id']), None)
                if subject:
                    available_teachers = [t for t in self.teachers if subject['name'] in t['subjects']]
                    if available_teachers:
                        teacher = random.choice(available_teachers)
                        self.schedule.loc[idx, 'teacher_id'] = teacher['id']
                        self.schedule.loc[idx, 'teacher_name'] = teacher['name']
                        if self.classrooms:
                            classroom = random.choice(self.classrooms)
                            self.schedule.loc[idx, 'classroom_id'] = classroom['id']
                            self.schedule.loc[idx, 'classroom_name'] = classroom['name']
                        self.schedule.loc[idx, 'status'] = 'подтверждено'
    def on_schedule_generated(self):
        self.progress.hide()
        self.statusBar.showMessage("Расписание сгенерировано")
        QMessageBox.information(self, "Успех", "Расписание успешно сгенерировано!")
        self.filter_schedule()
        self.update_reports()
        self.create_backup()
    # --- Фильтрация расписания ---
    def filter_schedule(self):
        # Очистка модели
        self.schedule_model.clear()
        # Устанавливаем заголовки снова
        self.schedule_model.setHorizontalHeaderLabels(['Время', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'])
        
        week_text = self.week_var.currentText()
        week_num = int(week_text.split()[1]) if week_text and "Неделя" in week_text else 1
        group_name = self.group_filter_var.currentText()
        teacher_name = self.teacher_filter_var.currentText()
        classroom_name = self.classroom_filter_var.currentText()
        
        filtered_schedule = self.schedule.copy()
        if not filtered_schedule.empty and 'week' in filtered_schedule.columns:
            if week_num:
                filtered_schedule = filtered_schedule[filtered_schedule['week'] == week_num]
            if group_name:
                filtered_schedule = filtered_schedule[filtered_schedule['group_name'] == group_name]
            if teacher_name:
                filtered_schedule = filtered_schedule[filtered_schedule['teacher_name'] == teacher_name]
            if classroom_name:
                filtered_schedule = filtered_schedule[filtered_schedule['classroom_name'] == classroom_name]
            
            if not filtered_schedule.empty:
                days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
                times = sorted(filtered_schedule['time'].unique())
                
                # Добавляем временные интервалы в первую колонку
                for time_slot in times:
                    time_item = QStandardItem(time_slot)
                    self.schedule_model.appendRow([time_item])
                    
                    # Добавляем данные для каждого дня недели
                    for col, day in enumerate(days, 1):
                        lesson = filtered_schedule[
                            (filtered_schedule['time'] == time_slot) &
                            (filtered_schedule['day'] == day) &
                            (filtered_schedule['status'] == 'подтверждено')
                        ]
                        if not lesson.empty:
                            lesson_info = lesson.iloc[0]
                            text = f"{lesson_info['group_name']}\n{lesson_info['subject_name']}\n{lesson_info['teacher_name']}\n{lesson_info['classroom_name']}"
                            item = QStandardItem(text)
                            item.setTextAlignment(Qt.AlignCenter)
                            self.schedule_model.setItem(self.schedule_model.rowCount()-1, col, item)
            else:
                self.show_empty_schedule()
        else:
            self.show_empty_schedule()

        # Применяем сортировку
        self.schedule_proxy_model.sort(0, Qt.AscendingOrder)  # Сортировка по первому столбцу (времени)

    def show_empty_schedule(self):
        # Очистка модели
        self.schedule_model.clear()
        self.schedule_model.setHorizontalHeaderLabels(['Время', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'])
        
        times = [f"{8+i}:00-{8+i}:45" for i in range(self.settings['lessons_per_day'])]
        days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
        
        for time_slot in times:
            row_items = [QStandardItem(time_slot)]
            for _ in range(len(days)):
                row_items.append(QStandardItem(""))
            self.schedule_model.appendRow(row_items)
    # --- Ручное управление занятиями ---
    def add_lesson(self):
        if not self.groups or not self.subjects or not self.teachers or not self.classrooms:
            QMessageBox.warning(self, "Предупреждение", "Для добавления занятия необходимо, чтобы были созданы хотя бы одна группа, один предмет, один преподаватель и одна аудитория.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавить занятие")
        dialog.setModal(True)
        dialog.resize(500, 400)
        form_layout = QFormLayout(dialog)
        week_var = QComboBox()
        week_var.addItems([str(i) for i in range(1, self.settings['weeks'] + 1)])
        day_var = QComboBox()
        day_var.addItems(['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']])
        time_var = QComboBox()
        time_var.addItems([f"{8+i}:00-{8+i}:45" for i in range(self.settings['lessons_per_day'])])
        group_var = QComboBox()
        group_var.addItems([g['name'] for g in self.groups])
        subject_var = QComboBox()
        subject_var.addItems([s['name'] for s in self.subjects])
        teacher_var = QComboBox()
        teacher_var.addItems([t['name'] for t in self.teachers])
        classroom_var = QComboBox()
        classroom_var.addItems([c['name'] for c in self.classrooms])
        form_layout.addRow("Неделя:", week_var)
        form_layout.addRow("День недели:", day_var)
        form_layout.addRow("Время:", time_var)
        form_layout.addRow("Группа:", group_var)
        form_layout.addRow("Предмет:", subject_var)
        form_layout.addRow("Преподаватель:", teacher_var)
        form_layout.addRow("Аудитория:", classroom_var)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._save_lesson(
            week_var.currentText(), day_var.currentText(), time_var.currentText(),
            group_var.currentText(), subject_var.currentText(), teacher_var.currentText(),
            classroom_var.currentText(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _save_lesson(self, week, day, time, group_name, subject_name, teacher_name, classroom_name, dialog):
        selected_group = next((g for g in self.groups if g['name'] == group_name), None)
        selected_subject = next((s for s in self.subjects if s['name'] == subject_name), None)
        selected_teacher = next((t for t in self.teachers if t['name'] == teacher_name), None)
        selected_classroom = next((c for c in self.classrooms if c['name'] == classroom_name), None)
        if not all([selected_group, selected_subject, selected_teacher, selected_classroom]):
            QMessageBox.critical(self, "Ошибка", "Не удалось найти выбранные элементы в базе данных")
            return
        existing_lesson = self.schedule[
            (self.schedule['week'] == int(week)) &
            (self.schedule['day'] == day) &
            (self.schedule['time'] == time) &
            (self.schedule['group_id'] == selected_group['id']) &
            (self.schedule['status'] != 'свободно')
        ]
        if not existing_lesson.empty:
            if QMessageBox.question(self, "Подтверждение", "В выбранное время у этой группы уже есть занятие. Заменить его?") != QMessageBox.Yes:
                return
        target_row = self.schedule[
            (self.schedule['week'] == int(week)) &
            (self.schedule['day'] == day) &
            (self.schedule['time'] == time) &
            (self.schedule['group_id'] == selected_group['id'])
        ]
        if not target_row.empty:
            idx = target_row.index[0]
            self.schedule.loc[idx, 'subject_id'] = selected_subject['id']
            self.schedule.loc[idx, 'subject_name'] = selected_subject['name']
            self.schedule.loc[idx, 'teacher_id'] = selected_teacher['id']
            self.schedule.loc[idx, 'teacher_name'] = selected_teacher['name']
            self.schedule.loc[idx, 'classroom_id'] = selected_classroom['id']
            self.schedule.loc[idx, 'classroom_name'] = selected_classroom['name']
            self.schedule.loc[idx, 'status'] = 'подтверждено'
            self.filter_schedule()
            self.update_reports()
            self.create_backup()
            QMessageBox.information(self, "Успех", "Занятие успешно добавлено!")
            dialog.accept()
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось найти подходящий слот в расписании")
    def edit_lesson(self):
        """Редактирует выбранное занятие."""
        selected_indexes = self.schedule_view.selectionModel().selectedIndexes()
        if not selected_indexes:
            QMessageBox.information(self, "Информация", "Выберите занятие в таблице для редактирования")
            return
        
        # Получаем индекс выбранной ячейки
        selected_index = selected_indexes[0]
        row = selected_index.row()
        col = selected_index.column()
        
        # Первый столбец (0) - это время. Если выбран он, то удалять нечего.
        if col == 0:
            QMessageBox.warning(self, "Предупреждение", "Выберите ячейку с занятием (не время)")
            return
        
        # Получаем время из первого столбца той же строки
        time_index = self.schedule_proxy_model.index(row, 0)
        time_slot = self.schedule_proxy_model.data(time_index)
        if not time_slot:
            QMessageBox.warning(self, "Ошибка", "Не удалось определить время занятия")
            return
        
        # Получаем день недели из заголовка столбца
        day_header = self.schedule_proxy_model.headerData(col, Qt.Horizontal, Qt.DisplayRole)
        if not day_header:
            QMessageBox.warning(self, "Ошибка", "Не удалось определить день недели")
            return
        selected_day = day_header.text().strip()
        
        # Получаем текущий номер недели из фильтра
        week_text = self.week_var.currentText()
        try:
            selected_week = int(week_text.split()[1]) if "Неделя" in week_text else 1
        except (ValueError, IndexError):
            selected_week = 1
        
        # Фильтруем DataFrame, чтобы найти конкретное занятие
        target_lesson = self.schedule[
            (self.schedule['week'] == selected_week) &
            (self.schedule['day'] == selected_day) &
            (self.schedule['time'] == time_slot) &
            (self.schedule['status'] == 'подтверждено')
        ]
        if target_lesson.empty:
            QMessageBox.information(self, "Информация", "Выбранное занятие не найдено в расписании")
            return
        
        # Подтверждение редактирования
        lesson_info = target_lesson.iloc[0]
        confirm_text = (
            f"Вы уверены, что хотите отредактировать это занятие?\n"
            f"Предмет: {lesson_info['subject_name']}\n"
            f"Группа: {lesson_info['group_name']}\n"
            f"Преподаватель: {lesson_info['teacher_name']}\n"
            f"Аудитория: {lesson_info['classroom_name']}\n"
            f"День: {selected_day}\n"
            f"Время: {time_slot}\n"
            f"Неделя: {selected_week}"
        )
        if QMessageBox.question(self, "Подтверждение редактирования", confirm_text) != QMessageBox.Yes:
            return
        
        # Открываем диалог редактирования (в этом примере просто показываем сообщение)
        # Здесь должен быть вызов диалогового окна для редактирования данных
        QMessageBox.information(self, "Информация", "Функция редактирования занятия пока не реализована.")
        
    def delete_lesson(self):
        """Удаляет выбранное занятие из расписания."""
        selected_indexes = self.schedule_view.selectionModel().selectedIndexes()
        if not selected_indexes:
            QMessageBox.information(self, "Информация", "Выберите занятие в таблице для удаления")
            return

        # Получаем индекс выбранной ячейки
        selected_index = selected_indexes[0]
        row = selected_index.row()
        col = selected_index.column()

        # Первый столбец (0) - это время. Если выбран он, то удалять нечего.
        if col == 0:
            QMessageBox.warning(self, "Предупреждение", "Выберите ячейку с занятием (не время)")
            return

        # ПОЛУЧАЕМ СОДЕРЖИМОЕ ВЫБРАННОЙ ЯЧЕЙКИ
        cell_data = self.schedule_proxy_model.data(selected_index)
        if not cell_data or not cell_data.strip():
            QMessageBox.warning(self, "Предупреждение", "Выбранная ячейка пуста. Нет занятия для удаления.")
            return

        # Получаем время из первого столбца той же строки через прокси-модель
        time_index = self.schedule_proxy_model.index(row, 0)
        time_slot = self.schedule_proxy_model.data(time_index)
        if not time_slot:
            QMessageBox.warning(self, "Ошибка", "Не удалось определить время занятия")
            return

        # >>> ИСПРАВЛЕНО: Получаем заголовок дня недели из модели <<<
        day_header = self.schedule_proxy_model.headerData(col, Qt.Horizontal, Qt.DisplayRole)
        if not day_header:
            QMessageBox.warning(self, "Ошибка", "Не удалось определить день недели")
            return
        selected_day = day_header.strip()

        # Получаем текущий номер недели из фильтра
        week_text = self.week_var.currentText()
        try:
            selected_week = int(week_text.split()[1]) if "Неделя" in week_text else 1
        except (ValueError, IndexError):
            selected_week = 1

        # Фильтруем DataFrame, чтобы найти конкретное занятие
        target_lesson = self.schedule[
            (self.schedule['week'] == selected_week) &
            (self.schedule['day'] == selected_day) &
            (self.schedule['time'] == time_slot) &
            (self.schedule['status'] == 'подтверждено')
        ]
        if target_lesson.empty:
            QMessageBox.information(self, "Информация", "Выбранное занятие не найдено в расписании")
            return

        # Подтверждение удаления
        lesson_info = target_lesson.iloc[0]
        confirm_text = (
            f"Вы уверены, что хотите удалить это занятие?\n"
            f"Предмет: {lesson_info['subject_name']}\n"
            f"Группа: {lesson_info['group_name']}\n"
            f"Преподаватель: {lesson_info['teacher_name']}\n"
            f"Аудитория: {lesson_info['classroom_name']}\n"
            f"День: {selected_day}\n"
            f"Время: {time_slot}\n"
            f"Неделя: {selected_week}"
        )
        if QMessageBox.question(self, "Подтверждение удаления", confirm_text) != QMessageBox.Yes:
            return

        # Удаляем занятие: сбрасываем все данные и устанавливаем статус 'свободно'
        idx = target_lesson.index[0]
        self.schedule.loc[idx, ['subject_id', 'subject_name', 'teacher_id', 'teacher_name', 'classroom_id', 'classroom_name']] = [None, '', None, '', None, '']
        self.schedule.loc[idx, 'status'] = 'свободно'
        # Обновляем отображение таблицы и отчеты
        self.filter_schedule()
        self.update_reports()
        self.create_backup()
        QMessageBox.information(self, "Успех", "Занятие успешно удалено!")
    
    def substitute_lesson(self):
        """Заменяет выбранное занятие."""
        selected_indexes = self.schedule_view.selectionModel().selectedIndexes()
        if not selected_indexes:
            QMessageBox.information(self, "Информация", "Выберите занятие в таблице для замены")
            return

        # Получаем индекс выбранной ячейки
        selected_index = selected_indexes[0]
        row = selected_index.row()
        col = selected_index.column()

        # Первый столбец (0) - это время. Если выбран он, то заменять нечего.
        if col == 0:
            QMessageBox.warning(self, "Предупреждение", "Выберите ячейку с занятием (не время)")
            return

        # Получаем время из первого столбца той же строки через прокси-модель
        time_index = self.schedule_proxy_model.index(row, 0)
        time_slot = self.schedule_proxy_model.data(time_index)
        if not time_slot:
            QMessageBox.warning(self, "Ошибка", "Не удалось определить время занятия")
            return

        # Получаем день недели из заголовка столбца
        day_header = self.schedule_view.horizontalHeaderItem(col)
        if not day_header:
            QMessageBox.warning(self, "Ошибка", "Не удалось определить день недели")
            return
        selected_day = day_header.text().strip()

        # Получаем текущий номер недели из фильтра
        week_text = self.week_var.currentText()
        try:
            selected_week = int(week_text.split()[1]) if "Неделя" in week_text else 1
        except (ValueError, IndexError):
            selected_week = 1

        # Фильтруем DataFrame, чтобы найти конкретное занятие
        target_lesson = self.schedule[
            (self.schedule['week'] == selected_week) &
            (self.schedule['day'] == selected_day) &
            (self.schedule['time'] == time_slot) &
            (self.schedule['status'] == 'подтверждено')
        ]

        if target_lesson.empty:
            QMessageBox.information(self, "Информация", "Выбранное занятие не найдено в расписании")
            return

        # Получаем информацию о занятии
        lesson_info = target_lesson.iloc[0]

        # Создаем диалоговое окно для выбора нового преподавателя
        dialog = QDialog(self)
        dialog.setWindowTitle("Замена занятия")
        dialog.setModal(True)
        dialog.resize(400, 200)
        layout = QVBoxLayout(dialog)

        info_label = QLabel(f"Замена для:\n{lesson_info['subject_name']}\nГруппа: {lesson_info['group_name']}\nДень: {selected_day}\nВремя: {time_slot}")
        info_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(info_label)

        teacher_label = QLabel("Новый преподаватель:")
        layout.addWidget(teacher_label)

        teacher_combo = QComboBox()
        # Добавляем всех преподавателей, которые ведут этот предмет
        available_teachers = [t for t in self.teachers if lesson_info['subject_name'] in t['subjects']]
        teacher_names = [t['name'] for t in available_teachers]
        teacher_combo.addItems(teacher_names)
        layout.addWidget(teacher_combo)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        if dialog.exec_() == QDialog.Accepted:
            new_teacher_name = teacher_combo.currentText()
            new_teacher = next((t for t in available_teachers if t['name'] == new_teacher_name), None)

            if new_teacher:
                # Обновляем данные в DataFrame
                idx = target_lesson.index[0]
                self.schedule.loc[idx, 'teacher_id'] = new_teacher['id']
                self.schedule.loc[idx, 'teacher_name'] = new_teacher['name']

                # Добавляем запись в журнал замен (если он реализован)
                self.substitutions.append({
                    'date': datetime.now().isoformat(),
                    'week': selected_week,
                    'day': selected_day,
                    'time': time_slot,
                    'group': lesson_info['group_name'],
                    'subject': lesson_info['subject_name'],
                    'old_teacher': lesson_info['teacher_name'],
                    'new_teacher': new_teacher_name
                })

                # Обновляем отображение таблицы и отчеты
                self.filter_schedule()
                self.update_reports()
                self.create_backup()

                QMessageBox.information(self, "Успех", f"Занятие успешно заменено!\nНовый преподаватель: {new_teacher_name}")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось найти выбранного преподавателя")

        # В упрощенной версии мы не можем точно определить день и неделю из одной ячейки.
        # Поэтому показываем сообщение.
        # QMessageBox.information(self, "Информация", "Замена занятия пока не реализована в Qt-версии.")
        
    # --- Замены ---
    def substitute_lesson(self):
        selected_items = self.schedule_view.selectionModel().selectedIndexes()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите занятие в таблице для замены")
            return
        # Получаем данные из строки таблицы
        row = selected_items[0].row()
        time_slot = self.schedule_proxy_model.data(self.schedule_proxy_model.index(row, 0))
        # В упрощенной версии мы не можем точно определить день и неделю из одной ячейки.
        # Поэтому показываем сообщение.
        QMessageBox.information(self, "Информация", "Замена занятия пока не реализована в Qt-версии.")
    # --- Календарь ---
    def show_calendar(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Календарь расписания")
        dialog.setModal(True)
        dialog.resize(600, 500)
        layout = QVBoxLayout(dialog)
        calendar = QCalendarWidget()
        layout.addWidget(calendar)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)
        dialog.exec_()
    # --- Проверка конфликтов ---
    def check_conflicts(self):
        if self.schedule.empty:
            QMessageBox.information(self, "Информация", "Сначала сгенерируйте расписание")
            return
        teacher_conflicts = pd.DataFrame()
        classroom_conflicts = pd.DataFrame()
        group_conflicts = pd.DataFrame()
        if 'teacher_id' in self.schedule.columns:
            teacher_conflicts = self.schedule[
                (self.schedule['status'] == 'подтверждено') &
                (self.schedule.duplicated(['teacher_id', 'day', 'time', 'week'], keep=False))
            ]
        if 'classroom_id' in self.schedule.columns:
            classroom_conflicts = self.schedule[
                (self.schedule['status'] == 'подтверждено') &
                (self.schedule.duplicated(['classroom_id', 'day', 'time', 'week'], keep=False))
            ]
        if 'group_id' in self.schedule.columns:
            group_conflicts = self.schedule[
                (self.schedule['status'] == 'подтверждено') &
                (self.schedule.duplicated(['group_id', 'day', 'time', 'week'], keep=False))
            ]
        conflict_text = f"Конфликты преподавателей: {len(teacher_conflicts)}\n"
        conflict_text += f"Конфликты аудиторий: {len(classroom_conflicts)}\n"
        conflict_text += f"Конфликты групп: {len(group_conflicts)}\n"
        if not teacher_conflicts.empty:
            conflict_text += "Конфликты преподавателей:\n"
            for _, conflict in teacher_conflicts.head(10).iterrows():
                conflict_text += f"  {conflict['teacher_name']} - {conflict['day']} {conflict['time']} (неделя {conflict['week']})\n"
        if not classroom_conflicts.empty:
            conflict_text += "\nКонфликты аудиторий:\n"
            for _, conflict in classroom_conflicts.head(10).iterrows():
                conflict_text += f"  {conflict['classroom_name']} - {conflict['day']} {conflict['time']} (неделя {conflict['week']})\n"
        if not group_conflicts.empty:
            conflict_text += "\nКонфликты групп:\n"
            for _, conflict in group_conflicts.head(10).iterrows():
                conflict_text += f"  {conflict['group_name']} - {conflict['day']} {conflict['time']} (неделя {conflict['week']})\n"
        self.conflicts_text.setText(conflict_text)
        QMessageBox.information(self, "Проверка конфликтов", f"Найдено конфликтов:\nПреподавателей: {len(teacher_conflicts)}\nАудиторий: {len(classroom_conflicts)}\nГрупп: {len(group_conflicts)}")
    # --- Оптимизация ---
    def optimize_schedule(self):
        if self.schedule.empty:
            QMessageBox.information(self, "Информация", "Сначала сгенерируйте расписание")
            return
        self.progress.show()
        self.statusBar.showMessage("Оптимизация расписания...")
        conflicts = pd.DataFrame()
        if 'teacher_id' in self.schedule.columns:
            conflicts = self.schedule[
                (self.schedule['status'] == 'подтверждено') &
                (self.schedule.duplicated(['teacher_id', 'day', 'time', 'week'], keep=False))
            ]
        optimized_count = 0
        for idx in conflicts.index[:10]:
            group_id = self.schedule.loc[idx, 'group_id']
            week = self.schedule.loc[idx, 'week']
            free_slots = self.schedule[
                (self.schedule['group_id'] == group_id) &
                (self.schedule['week'] == week) &
                (self.schedule['status'] == 'свободно')
            ].index
            if len(free_slots) > 0:
                new_slot = random.choice(free_slots)
                self.schedule.loc[new_slot, 'subject_id'] = self.schedule.loc[idx, 'subject_id']
                self.schedule.loc[new_slot, 'subject_name'] = self.schedule.loc[idx, 'subject_name']
                self.schedule.loc[new_slot, 'teacher_id'] = self.schedule.loc[idx, 'teacher_id']
                self.schedule.loc[new_slot, 'teacher_name'] = self.schedule.loc[idx, 'teacher_name']
                self.schedule.loc[new_slot, 'classroom_id'] = self.schedule.loc[idx, 'classroom_id']
                self.schedule.loc[new_slot, 'classroom_name'] = self.schedule.loc[idx, 'classroom_name']
                self.schedule.loc[new_slot, 'status'] = 'подтверждено'
                self.schedule.loc[idx, ['subject_id', 'subject_name', 'teacher_id', 'teacher_name', 'classroom_id', 'classroom_name']] = [None, '', None, '', None, '']
                self.schedule.loc[idx, 'status'] = 'свободно'
                optimized_count += 1
        self.progress.hide()
        self.statusBar.showMessage("Оптимизация завершена")
        QMessageBox.information(self, "Оптимизация", f"Оптимизировано {optimized_count} занятий")
        self.filter_schedule()
        self.update_reports()
        self.create_backup()
    # --- Отчеты ---
    def update_reports(self):
        if self.schedule.empty:
            return
        # Очистка таблиц
        self.teacher_report_tree.setRowCount(0)
        self.group_report_tree.setRowCount(0)
        self.summary_text.clear()
        # Нагрузка преподавателей
        if 'teacher_name' in self.schedule.columns and not self.schedule[self.schedule['status'] == 'подтверждено'].empty:
            teacher_load = self.schedule[self.schedule['status'] == 'подтверждено'].groupby('teacher_name').size()
            for teacher, hours in teacher_load.items():
                teacher_data = self.schedule[
                    (self.schedule['teacher_name'] == teacher) &
                    (self.schedule['status'] == 'подтверждено')
                ]
                groups = ', '.join(teacher_data['group_name'].unique())
                subjects = ', '.join(teacher_data['subject_name'].unique())
                row = self.teacher_report_tree.rowCount()
                self.teacher_report_tree.insertRow(row)
                self.teacher_report_tree.setItem(row, 0, QTableWidgetItem(teacher))
                self.teacher_report_tree.setItem(row, 1, QTableWidgetItem(str(hours)))
                self.teacher_report_tree.setItem(row, 2, QTableWidgetItem(groups))
                self.teacher_report_tree.setItem(row, 3, QTableWidgetItem(subjects))
        # Нагрузка групп
        if 'group_name' in self.schedule.columns and not self.schedule[self.schedule['status'] == 'подтверждено'].empty:
            group_load = self.schedule[self.schedule['status'] == 'подтверждено'].groupby('group_name').size()
            for group, hours in group_load.items():
                group_data = self.schedule[
                    (self.schedule['group_name'] == group) &
                    (self.schedule['status'] == 'подтверждено')
                ]
                subjects = ', '.join(group_data['subject_name'].unique())
                teachers = ', '.join(group_data['teacher_name'].unique())
                row = self.group_report_tree.rowCount()
                self.group_report_tree.insertRow(row)
                self.group_report_tree.setItem(row, 0, QTableWidgetItem(group))
                self.group_report_tree.setItem(row, 1, QTableWidgetItem(str(hours)))
                self.group_report_tree.setItem(row, 2, QTableWidgetItem(subjects))
                self.group_report_tree.setItem(row, 3, QTableWidgetItem(teachers))
        # Сводный отчет
        summary_text = f"📊 Сводный отчет по расписанию\n"
        summary_text += f"Учреждение: {self.settings.get('school_name', 'Не указано')}\n"
        summary_text += f"Директор: {self.settings.get('director', 'Не указано')}\n"
        summary_text += f"Учебный год: {self.settings.get('academic_year', 'Не указано')}\n"
        summary_text += f"Дата начала: {self.settings.get('start_date', 'Не указано')}\n"
        summary_text += f"📅 Расписание на {self.settings.get('weeks', 0)} недель\n"
        summary_text += f"📅 Дней в неделю: {self.settings.get('days_per_week', 0)}\n"
        summary_text += f"📅 Занятий в день: {self.settings.get('lessons_per_day', 0)}\n"
        summary_text += f"👥 Групп: {len(self.groups)}\n"
        summary_text += f"👨‍🏫 Преподавателей: {len(self.teachers)}\n"
        summary_text += f"🏫 Аудиторий: {len(self.classrooms)}\n"
        summary_text += f"📚 Предметов: {len(self.subjects)}\n"
        summary_text += f"🎉 Праздничных дней: {len(self.holidays)}\n"
        if not self.schedule.empty:
            confirmed_lessons = len(self.schedule[self.schedule['status'] == 'подтверждено'])
            planned_lessons = len(self.schedule[self.schedule['status'] == 'запланировано'])
            free_slots = len(self.schedule[self.schedule['status'] == 'свободно'])
            summary_text += f"✅ Подтвержденных занятий: {confirmed_lessons}\n"
            summary_text += f"📝 Запланированных занятий: {planned_lessons}\n"
            summary_text += f"🕒 Свободных слотов: {free_slots}\n"
        self.summary_text.setText(summary_text)
    def show_reports(self):
        # Переключаемся на вкладку отчетов
        self.notebook.setCurrentIndex(5) # Индекс вкладки "Отчеты"
        self.update_reports()
    # --- Экспорт ---
    def export_to_excel(self):
        if self.schedule.empty:
            QMessageBox.information(self, "Информация", "Сначала сгенерируйте расписание")
            return
        filename, _ = QFileDialog.getSaveFileName(self, "Сохранить расписание как", "", "Excel files (*.xlsx);;All files (*.*)")
        if not filename:
            return
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                if not self.schedule.empty:
                    confirmed_schedule = self.schedule[self.schedule['status'] == 'подтверждено']
                    if not confirmed_schedule.empty:
                        confirmed_schedule.to_excel(writer, sheet_name='Расписание', index=False)
                    else:
                        self.schedule.to_excel(writer, sheet_name='Расписание', index=False)
                else:
                    pd.DataFrame(columns=['id', 'week', 'day', 'time', 'group_id', 'group_name',
                                        'subject_id', 'subject_name', 'teacher_id', 'teacher_name',
                                        'classroom_id', 'classroom_name', 'status']).to_excel(
                        writer, sheet_name='Расписание', index=False)
                pd.DataFrame(self.groups).to_excel(writer, sheet_name='Группы', index=False)
                pd.DataFrame(self.teachers).to_excel(writer, sheet_name='Преподаватели', index=False)
                pd.DataFrame(self.classrooms).to_excel(writer, sheet_name='Аудитории', index=False)
                pd.DataFrame(self.subjects).to_excel(writer, sheet_name='Предметы', index=False)
                pd.DataFrame(self.holidays).to_excel(writer, sheet_name='Праздники', index=False)
                if not self.schedule.empty and not self.schedule[self.schedule['status'] == 'подтверждено'].empty:
                    teacher_report = self.schedule[self.schedule['status'] == 'подтверждено'].groupby('teacher_name').agg({
                        'group_name': lambda x: ', '.join(x.unique()),
                        'subject_name': lambda x: ', '.join(x.unique())
                    }).reset_index()
                    teacher_report.columns = ['Преподаватель', 'Группы', 'Предметы']
                    teacher_report['Часы'] = self.schedule[self.schedule['status'] == 'подтверждено'].groupby('teacher_name').size().values
                    teacher_report.to_excel(writer, sheet_name='Нагрузка преподавателей', index=False)
                    group_report = self.schedule[self.schedule['status'] == 'подтверждено'].groupby('group_name').agg({
                        'teacher_name': lambda x: ', '.join(x.unique()),
                        'subject_name': lambda x: ', '.join(x.unique())
                    }).reset_index()
                    group_report.columns = ['Группа', 'Преподаватели', 'Предметы']
                    group_report['Часы'] = self.schedule[self.schedule['status'] == 'подтверждено'].groupby('group_name').size().values
                    group_report.to_excel(writer, sheet_name='Нагрузка групп', index=False)
            QMessageBox.information(self, "Экспорт", f"Расписание успешно экспортировано в {filename}")
            self.create_backup()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка экспорта: {str(e)}")
    def export_to_html(self):
        if self.schedule.empty:
            QMessageBox.information(self, "Информация", "Сначала сгенерируйте расписание")
            return
        filename, _ = QFileDialog.getSaveFileName(self, "Сохранить расписание как HTML", "", "HTML files (*.html);;All files (*.*)")
        if not filename:
            return
        try:
            confirmed_schedule = self.schedule[self.schedule['status'] == 'подтверждено']
            if confirmed_schedule.empty:
                QMessageBox.warning(self, "Предупреждение", "Нет подтвержденных занятий для экспорта")
                return
            unique_groups = confirmed_schedule['group_name'].unique()
            unique_groups.sort()
            html_content = f"""
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Расписание занятий - {self.settings.get('school_name', 'Образовательное учреждение')}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f7fa;
        }}
        .header {{
            text-align: center;
            margin-bottom: 30px;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .header h1 {{
            margin: 0;
            font-size: 2.5em;
        }}
        .header p {{
            margin: 10px 0 0;
            font-size: 1.2em;
        }}
        .info {{
            background-color: white;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }}
        .schedule-container {{
            display: flex;
            flex-wrap: wrap;
            gap: 30px;
            justify-content: center;
        }}
        .schedule-table {{
            background-color: white;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            overflow: hidden;
            flex: 1 1 450px;
            min-width: 300px;
        }}
        .schedule-title {{
            background-color: #2c3e50;
            color: white;
            padding: 15px;
            text-align: center;
            font-weight: bold;
            font-size: 1.3em;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th, td {{
            padding: 12px;
            text-align: center;
            border: 1px solid #e0e0e0;
        }}
        th {{
            background-color: #34495e;
            color: white;
            font-weight: bold;
        }}
        tr:nth-child(even) {{
            background-color: #f8f9fa;
        }}
        tr:hover {{
            background-color: #e3f2fd;
        }}
        .lesson-info {{
            font-size: 0.9em;
        }}
        .no-lesson {{
            color: #95a5a6;
            font-style: italic;
        }}
        .footer {{
            text-align: center;
            margin-top: 40px;
            padding: 20px;
            color: #7f8c8d;
            font-size: 0.9em;
            border-top: 1px solid #ecf0f1;
        }}
        @media (max-width: 768px) {{
            .schedule-container {{
                flex-direction: column;
            }}
            body {{
                padding: 10px;
            }}
            .header h1 {{
                font-size: 2em;
            }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📅 Расписание занятий</h1>
        <p>{self.settings.get('school_name', 'Образовательное учреждение')}</p>
        <p>Учебный год: {self.settings.get('academic_year', 'Не указан')}</p>
    </div>
    <div class="info">
        <p><strong>Всего групп:</strong> {len(unique_groups)}</p>
        <p><strong>Период расписания:</strong> {self.settings.get('weeks', 0)} недель</p>
        <p><strong>Дней в неделю:</strong> {self.settings.get('days_per_week', 0)}</p>
        <p><strong>Занятий в день:</strong> {self.settings.get('lessons_per_day', 0)}</p>
    </div>
    <div class="schedule-container">
"""
            for group_name in unique_groups:
                group_schedule = confirmed_schedule[confirmed_schedule['group_name'] == group_name]
                days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
                times = sorted(group_schedule['time'].unique())
                html_content += f"""
        <div class="schedule-table">
            <div class="schedule-title">Группа: {group_name}</div>
            <table>
                <thead>
                    <tr>
                        <th>Время</th>
                        {''.join(f'<th>{day}</th>' for day in days)}
                    </tr>
                </thead>
                <tbody>
"""
                for time_slot in times:
                    html_content += f"                    <tr>\n                        <td><strong>{time_slot}</strong></td>\n"
                    for day in days:
                        lesson = group_schedule[
                            (group_schedule['time'] == time_slot) &
                            (group_schedule['day'] == day)
                        ]
                        if not lesson.empty:
                            lesson_info = lesson.iloc[0]
                            lesson_html = f"{lesson_info['subject_name']}<br><span class='lesson-info'>{lesson_info['teacher_name']}<br>{lesson_info['classroom_name']}</span>"
                        else:
                            lesson_html = "<span class='no-lesson'>—</span>"
                        html_content += f"                        <td>{lesson_html}</td>\n"
                    html_content += "                    </tr>\n"
                html_content += """                </tbody>
            </table>
        </div>
"""
            html_content += f"""
    </div>
    <div class="footer">
        <p>Расписание сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}</p>
        <p>Директор: {self.settings.get('director', 'Не указан')}</p>
    </div>
</body>
</html>
"""
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(html_content)
            QMessageBox.information(self, "Экспорт", f"Расписание успешно экспортировано в HTML-файл:\n{filename}")
            self.create_backup()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка экспорта в HTML: {str(e)}")
    # --- Праздники ---
    def add_holiday(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавить праздник")
        dialog.setModal(True)
        dialog.resize(400, 200)
        form_layout = QFormLayout(dialog)
        date_entry = QDateEdit()
        date_entry.setCalendarPopup(True)
        date_entry.setDate(datetime.now().date())
        name_entry = QLineEdit()
        type_combo = QComboBox()
        type_combo.addItems(["Государственный", "Учебный", "Каникулы"])
        form_layout.addRow("Дата:", date_entry)
        form_layout.addRow("Название:", name_entry)
        form_layout.addRow("Тип:", type_combo)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._save_holiday(date_entry.date().toString("yyyy-MM-dd"), name_entry.text(), type_combo.currentText(), dialog))
        button_box.rejected.connect(dialog.reject)
        form_layout.addRow(button_box)
        dialog.exec_()
    def _save_holiday(self, date_str, name, holiday_type, dialog):
        if not date_str or not name:
            QMessageBox.warning(self, "Предупреждение", "Заполните все поля")
            return
        try:
            datetime.strptime(date_str, '%Y-%m-%d')
            self.holidays.append({
                'date': date_str,
                'name': name,
                'type': holiday_type
            })
            self.load_holidays_data()
            self.create_backup()
            dialog.accept()
        except ValueError:
            QMessageBox.critical(self, "Ошибка", "Неверный формат даты. Используйте ГГГГ-ММ-ДД")
    def delete_holiday(self):
        selected_items = self.holidays_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите праздник для удаления")
            return
        if QMessageBox.question(self, "Подтверждение", "Удалить выбранный праздник?") == QMessageBox.Yes:
            row = selected_items[0].row()
            holiday_date = self.holidays_tree.item(row, 0).text()
            holiday_name = self.holidays_tree.item(row, 1).text()
            self.holidays = [h for h in self.holidays if not (h['date'] == holiday_date and h['name'] == holiday_name)]
            self.load_holidays_data()
            self.create_backup()
    # --- Архив расписаний ---
    def save_current_schedule(self):
        if self.schedule.empty:
            QMessageBox.warning(self, "Предупреждение", "Нет расписания для сохранения. Сначала сгенерируйте его.")
            return
        school_name = self.settings.get('school_name', 'Расписание').replace(" ", "_")
        academic_year = self.settings.get('academic_year', 'Год_не_указан').replace("/", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{school_name}_{academic_year}_{timestamp}.json"
        filepath = os.path.join(self.archive_dir, filename)
        try:
            schedule_dict = self.schedule.to_dict(orient='records') if not self.schedule.empty else []
            archive_data = {
                'settings': self.settings,
                'groups': self.groups,
                'teachers': self.teachers,
                'classrooms': self.classrooms,
                'subjects': self.subjects,
                'schedule': schedule_dict,
                'holidays': self.holidays,
                'substitutions': self.substitutions,
                'saved_at': datetime.now().isoformat()
            }
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(archive_data, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "Успех", f"Расписание успешно сохранено в архив!\nФайл: {filename}")
            self.load_archive_list()
            self.create_backup()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка сохранения расписания в архив: {str(e)}")
    def load_archive_list(self):
        self.archive_tree.setRowCount(0)
        try:
            archive_files = [f for f in os.listdir(self.archive_dir) if f.endswith('.json')]
            archive_files.sort(key=lambda x: os.path.getmtime(os.path.join(self.archive_dir, x)), reverse=True)
            for filename in archive_files:
                filepath = os.path.join(self.archive_dir, filename)
                stat = os.stat(filepath)
                creation_time = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    groups_count = len(data.get('groups', []))
                    teachers_count = len(data.get('teachers', []))
                    classrooms_count = len(data.get('classrooms', []))
                    subjects_count = len(data.get('subjects', []))
                    schedule = data.get('schedule', [])
                    lessons_count = len([s for s in schedule if s.get('status') == 'подтверждено'])
                    row = self.archive_tree.rowCount()
                    self.archive_tree.insertRow(row)
                    self.archive_tree.setItem(row, 0, QTableWidgetItem(filename))
                    self.archive_tree.setItem(row, 1, QTableWidgetItem(creation_time))
                    self.archive_tree.setItem(row, 2, QTableWidgetItem(str(groups_count)))
                    self.archive_tree.setItem(row, 3, QTableWidgetItem(str(teachers_count)))
                    self.archive_tree.setItem(row, 4, QTableWidgetItem(str(classrooms_count)))
                    self.archive_tree.setItem(row, 5, QTableWidgetItem(str(subjects_count)))
                    self.archive_tree.setItem(row, 6, QTableWidgetItem(str(lessons_count)))
                except Exception as e:
                    row = self.archive_tree.rowCount()
                    self.archive_tree.insertRow(row)
                    self.archive_tree.setItem(row, 0, QTableWidgetItem(filename))
                    self.archive_tree.setItem(row, 1, QTableWidgetItem(creation_time))
                    self.archive_tree.setItem(row, 2, QTableWidgetItem("Ошибка"))
                    self.archive_tree.setItem(row, 3, QTableWidgetItem("Ошибка"))
                    self.archive_tree.setItem(row, 4, QTableWidgetItem("Ошибка"))
                    self.archive_tree.setItem(row, 5, QTableWidgetItem("Ошибка"))
                    self.archive_tree.setItem(row, 6, QTableWidgetItem("Ошибка"))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка загрузки списка архива: {e}")
    def load_archived_schedule(self):
        selected_items = self.archive_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите расписание для загрузки")
            return
        row = selected_items[0].row()
        filename = self.archive_tree.item(row, 0).text()
        filepath = os.path.join(self.archive_dir, filename)
        if QMessageBox.question(self, "Подтверждение", f"Вы уверены, что хотите загрузить расписание из файла {filename}?\nТекущие данные будут заменены.") == QMessageBox.Yes:
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.settings = data.get('settings', self.settings)
                self.groups = data.get('groups', [])
                self.teachers = data.get('teachers', [])
                self.classrooms = data.get('classrooms', [])
                self.subjects = data.get('subjects', [])
                self.holidays = data.get('holidays', [])
                self.substitutions = data.get('substitutions', [])
                schedule_data = data.get('schedule', [])
                if schedule_data:
                    self.schedule = pd.DataFrame(schedule_data)
                else:
                    self.schedule = pd.DataFrame()
                self.load_groups_data()
                self.load_teachers_data()
                self.load_classrooms_data()
                self.load_subjects_data()
                self.load_holidays_data()
                self.filter_schedule()
                self.update_reports()
                QMessageBox.information(self, "Успех", f"Расписание успешно загружено из {filename}")
                self.create_backup()
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка загрузки расписания: {str(e)}")
    def delete_archived_schedule(self):
        selected_items = self.archive_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите расписание для удаления")
            return
        row = selected_items[0].row()
        filename = self.archive_tree.item(row, 0).text()
        filepath = os.path.join(self.archive_dir, filename)
        if QMessageBox.question(self, "Подтверждение", f"Вы уверены, что хотите удалить файл {filename}?") == QMessageBox.Yes:
            try:
                os.remove(filepath)
                self.load_archive_list()
                QMessageBox.information(self, "Успех", f"Расписание {filename} успешно удалено из архива")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка удаления: {str(e)}")
    def export_archived_schedule(self):
        selected_items = self.archive_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите расписание для экспорта")
            return
        row = selected_items[0].row()
        filename = self.archive_tree.item(row, 0).text()
        filepath = os.path.join(self.archive_dir, filename)
        save_path, _ = QFileDialog.getSaveFileName(self, f"Экспорт расписания {filename}", "", "Excel files (*.xlsx);;All files (*.*)")
        if not save_path:
            return
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            schedule_data = data.get('schedule', [])
            if not schedule_data:
                QMessageBox.warning(self, "Предупреждение", "В выбранном файле нет данных расписания")
                return
            archive_schedule = pd.DataFrame(schedule_data)
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                confirmed_schedule = archive_schedule[archive_schedule['status'] == 'подтверждено'] if not archive_schedule.empty else archive_schedule
                if not confirmed_schedule.empty:
                    confirmed_schedule.to_excel(writer, sheet_name='Расписание', index=False)
                else:
                    archive_schedule.to_excel(writer, sheet_name='Расписание', index=False)
                pd.DataFrame(data.get('groups', [])).to_excel(writer, sheet_name='Группы', index=False)
                pd.DataFrame(data.get('teachers', [])).to_excel(writer, sheet_name='Преподаватели', index=False)
                pd.DataFrame(data.get('classrooms', [])).to_excel(writer, sheet_name='Аудитории', index=False)
                pd.DataFrame(data.get('subjects', [])).to_excel(writer, sheet_name='Предметы', index=False)
                pd.DataFrame(data.get('holidays', [])).to_excel(writer, sheet_name='Праздники', index=False)
                if not archive_schedule.empty and not archive_schedule[archive_schedule['status'] == 'подтверждено'].empty:
                    teacher_report = archive_schedule[archive_schedule['status'] == 'подтверждено'].groupby('teacher_name').agg({
                        'group_name': lambda x: ', '.join(x.unique()),
                        'subject_name': lambda x: ', '.join(x.unique())
                    }).reset_index()
                    teacher_report.columns = ['Преподаватель', 'Группы', 'Предметы']
                    teacher_report['Часы'] = archive_schedule[archive_schedule['status'] == 'подтверждено'].groupby('teacher_name').size().values
                    teacher_report.to_excel(writer, sheet_name='Нагрузка преподавателей', index=False)
                    group_report = archive_schedule[archive_schedule['status'] == 'подтверждено'].groupby('group_name').agg({
                        'teacher_name': lambda x: ', '.join(x.unique()),
                        'subject_name': lambda x: ', '.join(x.unique())
                    }).reset_index()
                    group_report.columns = ['Группа', 'Преподаватели', 'Предметы']
                    group_report['Часы'] = archive_schedule[archive_schedule['status'] == 'подтверждено'].groupby('group_name').size().values
                    group_report.to_excel(writer, sheet_name='Нагрузка групп', index=False)
            QMessageBox.information(self, "Экспорт", f"Расписание успешно экспортировано в {save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка экспорта: {str(e)}")
    # --- Настройки ---
    def open_settings(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Настройки приложения")
        dialog.setModal(True)
        dialog.resize(550, 700)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        # Основные настройки
        basic_frame = QGroupBox("Основные настройки")
        basic_layout = QFormLayout(basic_frame)
        school_name_var = QLineEdit(self.settings.get('school_name', ''))
        director_var = QLineEdit(self.settings.get('director', ''))
        academic_year_var = QLineEdit(self.settings.get('academic_year', ''))
        start_date_var = QLineEdit(self.settings.get('start_date', datetime.now().date().isoformat()))
        basic_layout.addRow("Название учреждения:", school_name_var)
        basic_layout.addRow("Директор:", director_var)
        basic_layout.addRow("Учебный год:", academic_year_var)
        basic_layout.addRow("Дата начала года:", start_date_var)
        scroll_layout.addWidget(basic_frame)
        # Параметры расписания
        schedule_frame = QGroupBox("Параметры расписания")
        schedule_layout = QFormLayout(schedule_frame)
        days_per_week_var = QSpinBox()
        days_per_week_var.setRange(1, 7)
        days_per_week_var.setValue(self.settings.get('days_per_week', 5))
        lessons_per_day_var = QSpinBox()
        lessons_per_day_var.setRange(1, 12)
        lessons_per_day_var.setValue(self.settings.get('lessons_per_day', 6))
        weeks_var = QSpinBox()
        weeks_var.setRange(1, 52)
        weeks_var.setValue(self.settings.get('weeks', 2))
        schedule_layout.addRow("Дней в неделю:", days_per_week_var)
        schedule_layout.addRow("Занятий в день:", lessons_per_day_var)
        schedule_layout.addRow("Недель:", weeks_var)
        scroll_layout.addWidget(schedule_frame)
        # Расписание звонков
        bell_frame = QGroupBox("Расписание звонков")
        bell_layout = QFormLayout(bell_frame)
        bell_schedule_var = QLineEdit(self.settings.get('bell_schedule', '8:00-8:45,8:55-9:40,9:50-10:35,10:45-11:30,11:40-12:25,12:35-13:20'))
        bell_schedule_var.setReadOnly(True)
        bell_layout.addRow("Текущее расписание:", bell_schedule_var)
        open_editor_btn = QPushButton("⚙️ Открыть редактор...")
        bell_layout.addRow(open_editor_btn)
        bell_layout.addRow(QLabel("Редактор расписания звонков", font=QFont('Segoe UI', 9, QFont.StyleItalic)))
        scroll_layout.addWidget(bell_frame)
        # Настройки авто-бэкапа
        backup_frame = QGroupBox("Настройки авто-бэкапа")
        backup_layout = QFormLayout(backup_frame)
        auto_backup_var = QCheckBox()
        auto_backup_var.setChecked(self.settings.get('auto_backup', True))
        backup_interval_var = QSpinBox()
        backup_interval_var.setRange(1, 1440)
        backup_interval_var.setValue(self.settings.get('backup_interval', 30))
        max_backups_var = QSpinBox()
        max_backups_var.setRange(1, 100)
        max_backups_var.setValue(self.settings.get('max_backups', 10))
        backup_layout.addRow("Авто-бэкап:", auto_backup_var)
        backup_layout.addRow("Интервал бэкапа (мин):", backup_interval_var)
        backup_layout.addRow("Макс. бэкапов:", max_backups_var)
        scroll_layout.addWidget(backup_frame)
        scroll_area.setWidget(scroll_content)
        main_layout = QVBoxLayout(dialog)
        main_layout.addWidget(scroll_area)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self._save_settings(
            school_name_var.text(), director_var.text(), academic_year_var.text(), start_date_var.text(),
            days_per_week_var.value(), lessons_per_day_var.value(), weeks_var.value(),
            auto_backup_var.isChecked(), backup_interval_var.value(), max_backups_var.value(),
            bell_schedule_var.text(), dialog))
        button_box.rejected.connect(dialog.reject)
        main_layout.addWidget(button_box)
        # Подключаем кнопку редактора звонков
        def open_editor_wrapper():
            current_schedule = bell_schedule_var.text()
            editor = BellScheduleEditor(dialog, current_schedule)
            if editor.exec_() == QDialog.Accepted and editor.result is not None:
                bell_schedule_var.setText(editor.result)
        open_editor_btn.clicked.connect(open_editor_wrapper)
        dialog.exec_()
    def _save_settings(self, school_name, director, academic_year, start_date,
                      days_per_week, lessons_per_day, weeks,
                      auto_backup, backup_interval, max_backups,
                      bell_schedule, dialog):
        self.settings['school_name'] = school_name
        self.settings['director'] = director
        self.settings['academic_year'] = academic_year
        self.settings['start_date'] = start_date
        self.settings['days_per_week'] = days_per_week
        self.settings['lessons_per_day'] = lessons_per_day
        self.settings['weeks'] = weeks
        self.settings['auto_backup'] = auto_backup
        self.settings['backup_interval'] = backup_interval
        self.settings['max_backups'] = max_backups
        self.settings['bell_schedule'] = bell_schedule
        self.days_var.setValue(days_per_week)
        self.lessons_var.setValue(lessons_per_day)
        self.weeks_var.setValue(weeks)
        self.restart_auto_backup()
        self.update_backup_indicator()
        dialog.accept()
    # --- Бэкапы ---
    def open_backup_manager(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Менеджер бэкапов")
        dialog.setModal(True)
        dialog.resize(600, 400)
        layout = QVBoxLayout(dialog)
        label = QLabel("🛡️ Менеджер бэкапов")
        label.setFont(QFont('Arial', 14, QFont.Bold))
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)
        self.backup_tree = QTableWidget(0, 3)
        self.backup_tree.setHorizontalHeaderLabels(['Имя файла', 'Дата создания', 'Размер'])
        self.backup_tree.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.backup_tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.backup_tree, 1)
        # Контекстное меню
        self.backup_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.backup_tree.customContextMenuRequested.connect(self.show_backup_context_menu)
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        refresh_btn = QPushButton("🔄 Обновить")
        refresh_btn.clicked.connect(self.load_backup_list)
        create_btn = QPushButton("💾 Создать бэкап")
        create_btn.clicked.connect(self.create_backup)
        restore_btn = QPushButton("📂 Восстановить")
        restore_btn.clicked.connect(self.restore_backup)
        delete_btn = QPushButton("🗑️ Удалить")
        delete_btn.clicked.connect(self.delete_backup)
        button_layout.addWidget(refresh_btn)
        button_layout.addWidget(create_btn)
        button_layout.addWidget(restore_btn)
        button_layout.addWidget(delete_btn)
        button_layout.addStretch()
        layout.addWidget(button_frame)
        self.load_backup_list()
        dialog.exec_()
    def show_backup_context_menu(self, position):
        menu = QMenu()
        restore_action = QAction("📂 Восстановить", self)
        delete_action = QAction("🗑️ Удалить", self)
        refresh_action = QAction("🔄 Обновить список", self)
        menu.addAction(restore_action)
        menu.addAction(delete_action)
        menu.addSeparator()
        menu.addAction(refresh_action)
        action = menu.exec_(self.backup_tree.mapToGlobal(position))
        if action == restore_action:
            self.restore_backup()
        elif action == delete_action:
            self.delete_backup()
        elif action == refresh_action:
            self.load_backup_list()
    def load_backup_list(self):
        self.backup_tree.setRowCount(0)
        try:
            backup_files = [f for f in os.listdir(self.backup_dir) if f.endswith('.zip')]
            backup_files.sort(reverse=True)
            for filename in backup_files:
                filepath = os.path.join(self.backup_dir, filename)
                stat = os.stat(filepath)
                creation_time = datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
                size = stat.st_size
                if size < 1024:
                    size_str = f"{size} B"
                elif size < 1024**2:
                    size_str = f"{size/1024:.1f} KB"
                else:
                    size_str = f"{size/(1024**2):.1f} MB"
                row = self.backup_tree.rowCount()
                self.backup_tree.insertRow(row)
                self.backup_tree.setItem(row, 0, QTableWidgetItem(filename))
                self.backup_tree.setItem(row, 1, QTableWidgetItem(creation_time))
                self.backup_tree.setItem(row, 2, QTableWidgetItem(size_str))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка загрузки списка бэкапов: {e}")
    def restore_backup(self):
        selected_items = self.backup_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите бэкап для восстановления")
            return
        row = selected_items[0].row()
        filename = self.backup_tree.item(row, 0).text()
        filepath = os.path.join(self.backup_dir, filename)
        if QMessageBox.question(self, "Подтверждение", f"Вы уверены, что хотите восстановить данные из {filename}?\nТекущие данные будут потеряны.") == QMessageBox.Yes:
            try:
                with zipfile.ZipFile(filepath, 'r') as zip_ref:
                    zip_ref.extractall('.')
                QMessageBox.information(self, "Успех", f"Данные успешно восстановлены из {filename}")
                self.load_data()
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка восстановления: {e}")
    def delete_backup(self):
        selected_items = self.backup_tree.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Информация", "Выберите бэкап для удаления")
            return
        row = selected_items[0].row()
        filename = self.backup_tree.item(row, 0).text()
        filepath = os.path.join(self.backup_dir, filename)
        if QMessageBox.question(self, "Подтверждение", f"Вы уверены, что хотите удалить {filename}?") == QMessageBox.Yes:
            try:
                os.remove(filepath)
                self.load_backup_list()
                QMessageBox.information(self, "Успех", f"Бэкап {filename} успешно удален")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка удаления: {e}")
    def create_backup(self):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"backup_{timestamp}.zip"
            backup_filepath = os.path.join(self.backup_dir, backup_filename)
            temp_data_file = "temp_data.json"
            data = {
                'settings': self.settings,
                'groups': self.groups,
                'teachers': self.teachers,
                'classrooms': self.classrooms,
                'subjects': self.subjects,
                'schedule': self.schedule.to_dict() if not self.schedule.empty else {},
                'substitutions': self.substitutions,
                'holidays': self.holidays
            }
            with open(temp_data_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            with zipfile.ZipFile(backup_filepath, 'w') as zipf:
                zipf.write(temp_data_file)
            os.remove(temp_data_file)
            self.cleanup_old_backups()
            self.last_backup_time = datetime.now()
            self.update_backup_indicator()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка создания бэкапа: {str(e)}")
    def cleanup_old_backups(self):
        try:
            backup_files = [f for f in os.listdir(self.backup_dir) if f.endswith('.zip')]
            backup_files.sort(key=lambda x: os.path.getctime(os.path.join(self.backup_dir, x)))
            while len(backup_files) > self.settings.get('max_backups', 10):
                oldest_file = backup_files.pop(0)
                os.remove(os.path.join(self.backup_dir, oldest_file))
        except Exception as e:
            pass
    def start_auto_backup(self):
        if self.settings.get('auto_backup', True):
            interval = self.settings.get('backup_interval', 30) * 60 * 1000
            self.backup_timer = QTimer()
            self.backup_timer.timeout.connect(self.auto_backup)
            self.backup_timer.start(interval)
    def auto_backup(self):
        self.create_backup()
    def restart_auto_backup(self):
        if self.backup_timer:
            self.backup_timer.stop()
        self.start_auto_backup()
    def update_backup_indicator(self):
        if self.settings.get('auto_backup', True):
            self.backup_status_label.setText("Авто-бэкап: ВКЛ")
            self.backup_status_label.setStyleSheet(f"color: {COLORS['success']};")
            if self.last_backup_time:
                next_backup = self.last_backup_time + timedelta(minutes=self.settings.get('backup_interval', 30))
                self.next_backup_time = next_backup
                self.backup_info_label.setText(f"Следующий: {next_backup.strftime('%H:%M')}")
            else:
                self.backup_info_label.setText("Следующий: --:--")
        else:
            self.backup_status_label.setText("Авто-бэкап: ВЫКЛ")
            self.backup_status_label.setStyleSheet(f"color: {COLORS['danger']};")
            self.backup_info_label.setText("")
    # --- Прочее ---
    def check_and_update_experience(self):
        current_year = datetime.now().year
        last_update_year = self.settings.get('last_academic_year_update', current_year)
        if current_year > last_update_year:
            for teacher in self.teachers:
                teacher['experience'] = teacher.get('experience', 0) + (current_year - last_update_year)
            self.settings['last_academic_year_update'] = current_year
            self.load_teachers_data()
            self.create_backup()
    def find_free_slot(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Найти свободное время")
        dialog.setModal(True)
        dialog.resize(400, 250)
        layout = QVBoxLayout(dialog)
        search_type_var = QComboBox()
        search_type_var.addItems(["Группа", "Преподаватель", "Аудитория"])
        element_var = QComboBox()
        layout.addWidget(QLabel("Выберите тип:"))
        layout.addWidget(search_type_var)
        layout.addWidget(QLabel("Выберите элемент:"))
        layout.addWidget(element_var)
        def update_combo():
            search_type = search_type_var.currentText()
            if search_type == "Группа":
                values = [g['name'] for g in self.groups]
            elif search_type == "Преподаватель":
                values = [t['name'] for t in self.teachers]
            else:
                values = [c['name'] for c in self.classrooms]
            element_var.clear()
            element_var.addItems(values)
        search_type_var.currentIndexChanged.connect(update_combo)
        update_combo()
        def search_slot():
            search_type = search_type_var.currentText()
            element_name = element_var.currentText()
            if not element_name:
                QMessageBox.warning(self, "Предупреждение", "Выберите элемент")
                return
            free_slots = self.schedule[self.schedule['status'] == 'свободно']
            if search_type == "Группа":
                free_slots = free_slots[free_slots['group_name'] == element_name]
            elif search_type == "Преподаватель":
                free_slots = free_slots[free_slots['teacher_name'] == '']
            else:
                free_slots = free_slots[free_slots['classroom_name'] == '']
            if not free_slots.empty:
                first_slot = free_slots.iloc[0]
                QMessageBox.information(self, "Свободный слот", f"Ближайший свободный слот:\nНеделя: {first_slot['week']}\nДень: {first_slot['day']}\nВремя: {first_slot['time']}\nГруппа: {first_slot['group_name']}")
            else:
                QMessageBox.information(self, "Информация", "Свободных слотов не найдено")
        search_btn = QPushButton("Найти")
        search_btn.clicked.connect(search_slot)
        layout.addWidget(search_btn)
        dialog.exec_()
    def show_about(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("О программе")
        dialog.setModal(True)
        dialog.setFixedSize(605, 650)
        dialog.setWindowFlags(dialog.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        # Логотип
        logo_label = QLabel("🎓")
        logo_label.setFont(QFont('Arial', 48, QFont.Bold))
        logo_label.setAlignment(Qt.AlignCenter)
        scroll_layout.addWidget(logo_label)
        # Название
        title_label = QLabel("Система автоматического составления расписания")
        title_label.setFont(QFont('Arial', 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        scroll_layout.addWidget(title_label)
        # Версия
        version_label = QLabel("Версия 2.0")
        version_label.setFont(QFont('Arial', 11, QFont.Bold))
        version_label.setAlignment(Qt.AlignCenter)
        scroll_layout.addWidget(version_label)
        # Описание
        description_text = (
            "Это профессиональное приложение предназначено для автоматизации и управления\n"
            "процессом составления учебного расписания в образовательных учреждениях.\n"
            "Программа объединяет в себе мощный функционал автоматической генерации и гибкие\n"
            "инструменты ручного управления, что делает ее незаменимым помощником для\n"
            "администраторов и завучей."
        )
        desc_label = QLabel(description_text)
        desc_label.setWordWrap(True)
        desc_label.setAlignment(Qt.AlignLeft)
        scroll_layout.addWidget(desc_label)
        # Основные возможности
        features_title = QLabel("🔑 Основные возможности:")
        features_title.setFont(QFont('Arial', 12, QFont.Bold))
        features_title.setAlignment(Qt.AlignLeft)
        scroll_layout.addWidget(features_title)
        features_text = (
            "• 🚀 Автоматическая генерация расписания на основе заданных параметров\n"
            "• ✍️ Полный ручной контроль: добавление, редактирование и удаление занятий\n"
            "• 🔄 Управление заменами преподавателей с ведением журнала\n"
            "• 📅 Визуальный календарь для удобного просмотра расписания\n"
            "• 📊 Комплексная система отчетов по нагрузке преподавателей и групп\n"
            "• 🧩 Проверка и автоматическая оптимизация для устранения конфликтов\n"
            "• 💾 Архивирование и загрузка ранее сгенерированных расписаний\n"
            "• 🌐 Экспорт расписания в HTML для публикации на сайте учреждения\n"
            "• 📤 Экспорт данных и отчетов в формат Excel\n"
            "• 🛡️Автоматическое и ручное создание резервных копий данных\n"
            "• 🎓 Управление базами данных: группы, преподаватели, аудитории, предметы, праздники"
        )
        features_label = QLabel(features_text)
        features_label.setWordWrap(True)
        features_label.setAlignment(Qt.AlignLeft)
        scroll_layout.addWidget(features_label)
        # Информация о разработчике
        dev_frame = QFrame()
        dev_layout = QVBoxLayout(dev_frame)
        dev_layout.addWidget(QLabel("👨‍💻 Разработчик:", font=QFont('Arial', 10, QFont.Bold)))
        dev_layout.addWidget(QLabel("Команда разработчиков Mikhail Lukomskiy"))
        dev_layout.addWidget(QLabel("📧 Контактная почта: support@lukomsky.ru"))
        dev_layout.addWidget(QLabel("🌐 Официальный сайт: www.lukomsky.ru"))
        dev_layout.addWidget(QLabel("📅 Год выпуска: 2025"))
        scroll_layout.addWidget(dev_frame)
        # Информация об учреждении
        school_frame = QFrame()
        school_layout = QVBoxLayout(school_frame)
        school_layout.addWidget(QLabel("🏫 Текущее учреждение:", font=QFont('Arial', 10, QFont.Bold)))
        school_layout.addWidget(QLabel(f"{self.settings.get('school_name', 'Не указано')}"))
        school_layout.addWidget(QLabel(f"Директор: {self.settings.get('director', 'Не указан')}"))
        school_layout.addWidget(QLabel(f"Учебный год: {self.settings.get('academic_year', 'Не указан')}"))
        scroll_layout.addWidget(school_frame)
        scroll_area.setWidget(scroll_content)
        main_layout = QVBoxLayout(dialog)
        main_layout.addWidget(scroll_area)
        close_button = QPushButton("Закрыть")
        close_button.clicked.connect(dialog.accept)
        main_layout.addWidget(close_button)
        dialog.exec_()
    def save_data(self):
        data = {
            'settings': self.settings,
            'groups': self.groups,
            'teachers': self.teachers,
            'classrooms': self.classrooms,
            'subjects': self.subjects,
            'substitutions': self.substitutions,
            'holidays': self.holidays
        }
        filename, _ = QFileDialog.getSaveFileName(self, "Сохранить данные как", "", "JSON files (*.json);;All files (*.*)")
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                QMessageBox.information(self, "Сохранение", "Данные успешно сохранены!")
                self.create_backup()
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка сохранения: {str(e)}")
    def load_data(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Загрузить данные", "", "JSON files (*.json);;All files (*.*)")
        if filename and os.path.exists(filename):
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.settings = data.get('settings', self.settings)
                self.groups = data.get('groups', [])
                self.teachers = data.get('teachers', [])
                self.classrooms = data.get('classrooms', [])
                self.subjects = data.get('subjects', [])
                self.substitutions = data.get('substitutions', [])
                self.holidays = data.get('holidays', [])
                self.load_groups_data()
                self.load_teachers_data()
                self.load_classrooms_data()
                self.load_subjects_data()
                self.load_holidays_data()
                QMessageBox.information(self, "Загрузка", "Данные успешно загружены")
                self.create_backup()
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка загрузки: {str(e)}")
    def open_substitutions(self):
        # В упрощенной версии журнал замен не реализован.
        QMessageBox.information(self, "Информация", "Журнал замен пока не реализован в Qt-версии.")
    # --- Запуск приложения ---
    def run(self):
        self.show()
# Запуск приложения
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScheduleApp()
    window.run()
    sys.exit(app.exec_())
