import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import threading
import json
import os
import calendar
import shutil
import zipfile
from pathlib import Path

class BellScheduleEditor:
    """Диалоговое окно для редактирования расписания звонков"""
    def __init__(self, parent, current_schedule_str):
        self.parent = parent
        self.result = None  # Будет содержать итоговую строку расписания
        self.slots = self._parse_schedule_string(current_schedule_str)
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Редактор расписания звонков")
        self.dialog.geometry("500x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(False, False)
        
        self.create_widgets()
        
    def _parse_schedule_string(self, schedule_str):
        """Преобразует строку расписания в список кортежей (начало, конец)"""
        if not schedule_str:
            return []
        try:
            slots = []
            for slot_str in schedule_str.split(','):
                start, end = slot_str.strip().split('-')
                slots.append((start.strip(), end.strip()))
            return slots
        except Exception:
            # В случае ошибки возвращаем стандартное расписание
            return [("8:00", "8:45"), ("8:55", "9:40"), ("9:50", "10:35"), 
                    ("10:45", "11:30"), ("11:40", "12:25"), ("12:35", "13:20")]
    
    def _format_schedule_string(self):
        """Преобразует список слотов в строку для сохранения"""
        return ','.join([f"{start}-{end}" for start, end in self.slots])
    
    def create_widgets(self):
        # Основной фрейм
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        ttk.Label(main_frame, text="Уроки:", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        # Фрейм для списка слотов с прокруткой
        slots_frame = ttk.Frame(main_frame)
        slots_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Создаем Treeview для отображения слотов
        columns = ('№', 'Начало', 'Конец')
        self.tree = ttk.Treeview(slots_frame, columns=columns, show='headings', height=10)
        self.tree.heading('№', text='№')
        self.tree.heading('Начало', text='Начало')
        self.tree.heading('Конец', text='Конец')
        self.tree.column('№', width=40, anchor='center')
        self.tree.column('Начало', width=100, anchor='center')
        self.tree.column('Конец', width=100, anchor='center')
        
        # Прокрутка
        scrollbar = ttk.Scrollbar(slots_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Кнопки управления
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="➕ Добавить", command=self.add_slot).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="✏️ Редактировать", command=self.edit_slot).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="🗑️ Удалить", command=self.delete_slot).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="⬆️ Вверх", command=self.move_up).pack(side=tk.LEFT, padx=(20, 5))
        ttk.Button(btn_frame, text="⬇️ Вниз", command=self.move_down).pack(side=tk.LEFT)
        
        # Кнопки ОК/Отмена
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(action_frame, text="Отмена", command=self.dialog.destroy).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(action_frame, text="ОК", command=self.save_and_close).pack(side=tk.RIGHT)
        
        # Загружаем данные в таблицу
        self.load_slots_to_tree()
    
    def load_slots_to_tree(self):
        """Загружает список слотов в Treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for i, (start, end) in enumerate(self.slots, 1):
            self.tree.insert('', tk.END, values=(i, start, end))
    
    def add_slot(self):
        """Добавить новый слот"""
        self._open_slot_editor("Добавить слот")
    
    def edit_slot(self):
        """Редактировать выбранный слот"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите слот для редактирования")
            return
        
        item = self.tree.item(selected[0])
        values = item['values']
        slot_index = int(values[0]) - 1  # Номер слота (начинается с 1)
        
        self._open_slot_editor("Редактировать слот", slot_index)
    
    def _open_slot_editor(self, title, slot_index=None):
        """Открывает диалоговое окно для редактирования одного слота"""
        editor = tk.Toplevel(self.dialog)
        editor.title(title)
        editor.geometry("300x150")
        editor.transient(self.dialog)
        editor.grab_set()
        editor.resizable(False, False)
        
        ttk.Label(editor, text="Начало (ЧЧ:ММ):").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        start_var = tk.StringVar()
        start_entry = ttk.Entry(editor, textvariable=start_var, width=10)
        start_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(editor, text="Конец (ЧЧ:ММ):").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        end_var = tk.StringVar()
        end_entry = ttk.Entry(editor, textvariable=end_var, width=10)
        end_entry.grid(row=1, column=1, padx=10, pady=10)
        
        if slot_index is not None:
            # Заполняем поля текущими значениями
            start_var.set(self.slots[slot_index][0])
            end_var.set(self.slots[slot_index][1])
        
        def save_slot():
            start_time = start_var.get().strip()
            end_time = end_var.get().strip()
            
            if not self._validate_time_format(start_time) or not self._validate_time_format(end_time):
                messagebox.showerror("Ошибка", "Неверный формат времени. Используйте ЧЧ:ММ (например, 08:30)")
                return
            
            if slot_index is None:
                # Добавление нового слота
                self.slots.append((start_time, end_time))
            else:
                # Редактирование существующего слота
                self.slots[slot_index] = (start_time, end_time)
            
            self.load_slots_to_tree()
            editor.destroy()
        
        ttk.Button(editor, text="Сохранить", command=save_slot).grid(row=2, column=0, columnspan=2, pady=10)
    
    def _validate_time_format(self, time_str):
        """Проверяет, соответствует ли строка формату ЧЧ:ММ"""
        import re
        pattern = r'^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$'
        return re.match(pattern, time_str) is not None
    
    def delete_slot(self):
        """Удалить выбранный слот"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите слот для удаления")
            return
        
        if messagebox.askyesno("Подтверждение", "Удалить выбранный слот?"):
            item = self.tree.item(selected[0])
            values = item['values']
            slot_index = int(values[0]) - 1
            del self.slots[slot_index]
            self.load_slots_to_tree()
    
    def move_up(self):
        """Переместить выбранный слот вверх"""
        selected = self.tree.selection()
        if not selected:
            return
        
        item = self.tree.item(selected[0])
        values = item['values']
        slot_index = int(values[0]) - 1
        
        if slot_index > 0:
            # Меняем местами с предыдущим слотом
            self.slots[slot_index], self.slots[slot_index - 1] = self.slots[slot_index - 1], self.slots[slot_index]
            self.load_slots_to_tree()
            # Выбираем тот же слот после обновления
            new_selection_id = self.tree.get_children()[slot_index - 1]
            self.tree.selection_set(new_selection_id)
    
    def move_down(self):
        """Переместить выбранный слот вниз"""
        selected = self.tree.selection()
        if not selected:
            return
        
        item = self.tree.item(selected[0])
        values = item['values']
        slot_index = int(values[0]) - 1
        
        if slot_index < len(self.slots) - 1:
            # Меняем местами со следующим слотом
            self.slots[slot_index], self.slots[slot_index + 1] = self.slots[slot_index + 1], self.slots[slot_index]
            self.load_slots_to_tree()
            # Выбираем тот же слот после обновления
            new_selection_id = self.tree.get_children()[slot_index + 1]
            self.tree.selection_set(new_selection_id)
    
    def save_and_close(self):
        """Сохраняет результат и закрывает диалог"""
        self.result = self._format_schedule_string()
        self.dialog.destroy()

class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎓 Система автоматического составления расписания")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 800)
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
            'backup_interval': 30,  # минуты
            'max_backups': 10,
            'last_academic_year_update': datetime.now().year  # Год последнего обновления стажа
        }
        self.groups = []
        self.teachers = []
        self.classrooms = []
        self.subjects = []
        self.schedule = pd.DataFrame()
        self.substitutions = []  # Журнал замен
        self.holidays = []  # Праздничные дни
        self.backup_timer = None
        self.last_backup_time = None
        self.next_backup_time = None
        # Инициализация переменных для комбобоксов
        self.week_var = tk.StringVar()
        self.group_filter_var = tk.StringVar()
        self.teacher_filter_var = tk.StringVar()
        self.classroom_filter_var = tk.StringVar()
        # Создание директории для бэкапов
        self.backup_dir = "backups"
        if not os.path.exists(self.backup_dir):
            os.makedirs(self.backup_dir)
        self.create_widgets()
        self.load_data()
        self.start_auto_backup()
        self.check_and_update_experience()  # Проверка и обновление стажа при запуске

    def create_widgets(self):
        # Создание стилей
        style = ttk.Style()
        style.theme_use('clam')  # Используем тему 'clam' для большей кастомизации

        # --- НОВАЯ ЦВЕТОВАЯ ПАЛИТРА ---
        colors = {
            'primary': '#4a6fa5',      # Основной синий
            'secondary': '#166088',    # Темно-синий
            'accent': '#47b8e0',       # Акцентный голубой
            'success': '#3cba92',      # Зеленый (успех)
            'warning': '#f4a261',      # Оранжевый (предупреждение)
            'danger': '#e71d36',       # Красный (ошибка)
            'light': '#f8f9fa',        # Светлый фон
            'dark': '#343a40',         # Темный текст
            'border': '#dee2e6',       # Цвет границ
            'hover': '#e9ecef'         # Цвет при наведении
        }

        # Настройка стилей
        style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'), foreground=colors['secondary'])
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'), foreground=colors['dark'])
        style.configure('TButton', font=('Segoe UI', 10), padding=6)
        style.configure('TNotebook.Tab', padding=[12, 8], font=('Segoe UI', 10))
        style.configure('Backup.TLabel', font=('Segoe UI', 10, 'bold'))
        style.configure('BackupActive.TLabel', font=('Segoe UI', 10, 'bold'), foreground=colors['success'])
        style.configure('BackupInactive.TLabel', font=('Segoe UI', 10, 'bold'), foreground=colors['danger'])

        # Стиль для кнопок
        style.configure('Accent.TButton', background=colors['accent'], foreground='white', font=('Segoe UI', 10, 'bold'))
        style.map('Accent.TButton', background=[('active', '#3aa1c9')]) # Цвет при наведении

        # Стиль для Treeview (таблиц)
        style.configure('Treeview', rowheight=80, font=('Segoe UI', 10), background='white', fieldbackground='white') # Увеличена высота строки
        style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'), background=colors['primary'], foreground='white')
        style.map('Treeview', background=[('selected', colors['accent'])], foreground=[('selected', 'white')])
        # --- КОНЕЦ НОВЫХ СТИЛЕЙ ---
        # Главный фрейм
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        # Заголовок
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        title_label = ttk.Label(title_frame, text="🎓 Система автоматического составления расписания", 
                               style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        # Кнопки управления
        buttons_frame = ttk.Frame(title_frame)
        buttons_frame.pack(side=tk.RIGHT)
        ttk.Button(buttons_frame, text="💾 Сохранить", 
                  command=self.save_data).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="📂 Загрузить", 
                  command=self.load_data).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="⚙️ Настройки", 
                  command=self.open_settings).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="🛡️ Бэкап", 
                  command=self.open_backup_manager).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="❓ О программе", 
                  command=self.show_about).pack(side=tk.LEFT)
        # Индикатор авто-бэкапа
        self.backup_indicator_frame = ttk.Frame(title_frame)
        self.backup_indicator_frame.pack(side=tk.RIGHT, padx=(20, 0))
        self.backup_status_label = ttk.Label(self.backup_indicator_frame, 
                                           text="Авто-бэкап: ВКЛ", 
                                           style='BackupActive.TLabel')
        self.backup_status_label.pack(side=tk.TOP)
        self.backup_info_label = ttk.Label(self.backup_indicator_frame, 
                                         text="Следующий: --:--", 
                                         style='Backup.TLabel')
        self.backup_info_label.pack(side=tk.TOP)
        # Верхняя панель с настройками
        settings_frame = ttk.LabelFrame(main_frame, text="⚙️ Настройки расписания", padding=10)
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        # Параметры расписания
        params_frame = ttk.Frame(settings_frame)
        params_frame.pack(fill=tk.X)
        ttk.Label(params_frame, text="Дней в неделю:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.days_var = tk.StringVar(value=str(self.settings['days_per_week']))
        days_spin = ttk.Spinbox(params_frame, from_=1, to=7, width=5, textvariable=self.days_var)
        days_spin.grid(row=0, column=1, padx=(0, 20))
        ttk.Label(params_frame, text="Занятий в день:").grid(row=0, column=2, sticky=tk.W, padx=(0, 10))
        self.lessons_var = tk.StringVar(value=str(self.settings['lessons_per_day']))
        lessons_spin = ttk.Spinbox(params_frame, from_=1, to=12, width=5, textvariable=self.lessons_var)
        lessons_spin.grid(row=0, column=3, padx=(0, 20))
        ttk.Label(params_frame, text="Недель:").grid(row=0, column=4, sticky=tk.W, padx=(0, 10))
        self.weeks_var = tk.StringVar(value=str(self.settings['weeks']))
        weeks_spin = ttk.Spinbox(params_frame, from_=1, to=12, width=5, textvariable=self.weeks_var)
        weeks_spin.grid(row=0, column=5, padx=(0, 20))
        # Кнопки управления
        buttons_frame2 = ttk.Frame(settings_frame)
        buttons_frame2.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(buttons_frame2, text="🚀 Сгенерировать расписание", 
                  command=self.generate_schedule_thread).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame2, text="🔍 Проверить конфликт", 
                  command=self.check_conflicts).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame2, text="⚡ Оптимизировать", 
                  command=self.optimize_schedule).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame2, text="📊 Отчеты", 
                  command=self.show_reports).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame2, text="📤 Экспорт в Excel", 
                  command=self.export_to_excel).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame2, text="🌐 Экспорт в HTML", 
                  command=self.export_to_html).pack(side=tk.LEFT, padx=(0, 10)) # Новая кнопка
        ttk.Button(buttons_frame2, text="🔄 Журнал замен", 
                  command=self.open_substitutions).pack(side=tk.LEFT)
        # Прогресс-бар
        self.progress = ttk.Progressbar(settings_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=(10, 0))
        # Основная область с вкладками
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        # Вкладка Группы
        self.groups_frame = ttk.Frame(notebook)
        notebook.add(self.groups_frame, text="👥 Группы")
        self.create_groups_tab()
        # Вкладка Преподаватели
        self.teachers_frame = ttk.Frame(notebook)
        notebook.add(self.teachers_frame, text="👨‍🏫 Преподаватели")
        self.create_teachers_tab()
        # Вкладка Аудитории
        self.classrooms_frame = ttk.Frame(notebook)
        notebook.add(self.classrooms_frame, text="🏫 Аудитории")
        self.create_classrooms_tab()
        # Вкладка Предметы
        self.subjects_frame = ttk.Frame(notebook)
        notebook.add(self.subjects_frame, text="📚 Предметы")
        self.create_subjects_tab()
        # Вкладка Расписание
        self.schedule_frame = ttk.Frame(notebook)
        notebook.add(self.schedule_frame, text="📅 Расписание")
        self.create_schedule_tab()
        # Вкладка Отчеты
        self.reports_frame = ttk.Frame(notebook)
        notebook.add(self.reports_frame, text="📈 Отчеты")
        self.create_reports_tab()
        # Вкладка Праздники
        self.holidays_frame = ttk.Frame(notebook)
        notebook.add(self.holidays_frame, text="🎉 Праздники")
        self.create_holidays_tab()
        # --- НОВАЯ ВКЛАДКА: Архив расписаний ---
        self.archive_frame = ttk.Frame(notebook)
        notebook.add(self.archive_frame, text="💾 Архив расписаний")
        self.create_archive_tab()
        # Статусная строка
        self.status_var = tk.StringVar(value="Готов к работе")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(10, 0))
        # Обновляем индикатор бэкапа
        self.update_backup_indicator()

    def save_current_schedule(self):
        """Сохранить текущее расписание в архив"""
        if self.schedule.empty:
            messagebox.showwarning("Предупреждение", "Нет расписания для сохранения. Сначала сгенерируйте его.")
            return

        # Создаем имя файла
        school_name = self.settings.get('school_name', 'Расписание').replace(" ", "_")
        academic_year = self.settings.get('academic_year', 'Год_не_указан').replace("/", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{school_name}_{academic_year}_{timestamp}.json"
        filepath = os.path.join(self.archive_dir, filename)

        try:
            # Подготавливаем данные для сохранения
            # Преобразуем DataFrame в словарь для сериализации
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

            messagebox.showinfo("Успех", f"Расписание успешно сохранено в архив!\nФайл: {filename}")
            self.load_archive_list() # Обновляем список
            self.create_backup() # Создаем бэкап после сохранения в архив

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка сохранения расписания в архив: {str(e)}")

    def load_archive_list(self):
        """Загрузить список архивных расписаний"""
        # Очистка таблицы
        for item in self.archive_tree.get_children():
            self.archive_tree.delete(item)

        try:
            # Получаем список файлов
            archive_files = [f for f in os.listdir(self.archive_dir) if f.endswith('.json')]
            archive_files.sort(key=lambda x: os.path.getmtime(os.path.join(self.archive_dir, x)), reverse=True) # Сортировка по дате (новые первые)

            for filename in archive_files:
                filepath = os.path.join(self.archive_dir, filename)
                stat = os.stat(filepath)
                creation_time = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')

                # Читаем файл, чтобы получить статистику
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    groups_count = len(data.get('groups', []))
                    teachers_count = len(data.get('teachers', []))
                    classrooms_count = len(data.get('classrooms', []))
                    subjects_count = len(data.get('subjects', []))
                    # Считаем только подтвержденные занятия
                    schedule = data.get('schedule', [])
                    lessons_count = len([s for s in schedule if s.get('status') == 'подтверждено'])

                    self.archive_tree.insert('', tk.END, values=(
                        filename,
                        creation_time,
                        groups_count,
                        teachers_count,
                        classrooms_count,
                        subjects_count,
                        lessons_count
                    ))
                except Exception as e:
                    # Если файл поврежден, показываем только имя и дату
                    self.archive_tree.insert('', tk.END, values=(
                        filename,
                        creation_time,
                        "Ошибка",
                        "Ошибка",
                        "Ошибка",
                        "Ошибка",
                        "Ошибка"
                    ))

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка загрузки списка архива: {e}")

    def load_archived_schedule(self):
        """Загрузить выбранное расписание из архива"""
        selected = self.archive_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите расписание для загрузки")
            return

        item = self.archive_tree.item(selected[0])
        filename = item['values'][0]
        filepath = os.path.join(self.archive_dir, filename)

        if messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите загрузить расписание из файла {filename}?\nТекущие данные будут заменены."):
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                # Загружаем данные
                self.settings = data.get('settings', self.settings)
                self.groups = data.get('groups', [])
                self.teachers = data.get('teachers', [])
                self.classrooms = data.get('classrooms', [])
                self.subjects = data.get('subjects', [])
                self.holidays = data.get('holidays', [])
                self.substitutions = data.get('substitutions', [])

                # Загружаем расписание
                schedule_data = data.get('schedule', [])
                if schedule_data:
                    self.schedule = pd.DataFrame(schedule_data)
                else:
                    self.schedule = pd.DataFrame()

                # Обновляем интерфейс
                self.load_groups_data()
                self.load_teachers_data()
                self.load_classrooms_data()
                self.load_subjects_data()
                self.load_holidays_data()
                self.filter_schedule()
                self.update_reports()

                messagebox.showinfo("Успех", f"Расписание успешно загружено из {filename}")
                self.create_backup() # Создаем бэкап после загрузки

            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка загрузки расписания: {str(e)}")

    def delete_archived_schedule(self):
        """Удалить выбранное расписание из архива"""
        selected = self.archive_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите расписание для удаления")
            return

        item = self.archive_tree.item(selected[0])
        filename = item['values'][0]
        filepath = os.path.join(self.archive_dir, filename)

        if messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите удалить файл {filename}?"):
            try:
                os.remove(filepath)
                self.load_archive_list() # Обновляем список
                messagebox.showinfo("Успех", f"Расписание {filename} успешно удалено из архива")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка удаления: {str(e)}")
                
    def export_archived_schedule(self):
        """Экспорт выбранного архивного расписания в Excel"""
        selected = self.archive_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите расписание для экспорта")
            return

        item = self.archive_tree.item(selected[0])
        filename = item['values'][0]
        filepath = os.path.join(self.archive_dir, filename)

        # Выбираем место для сохранения Excel-файла
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title=f"Экспорт расписания {filename}"
        )

        if not save_path:
            return

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # Создаем DataFrame из архивного расписания
            schedule_data = data.get('schedule', [])
            if not schedule_data:
                messagebox.showwarning("Предупреждение", "В выбранном файле нет данных расписания")
                return

            archive_schedule = pd.DataFrame(schedule_data)

            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                # Основное расписание
                confirmed_schedule = archive_schedule[archive_schedule['status'] == 'подтверждено'] if not archive_schedule.empty else archive_schedule
                if not confirmed_schedule.empty:
                    confirmed_schedule.to_excel(writer, sheet_name='Расписание', index=False)
                else:
                    archive_schedule.to_excel(writer, sheet_name='Расписание', index=False)

                # Исходные данные
                pd.DataFrame(data.get('groups', [])).to_excel(writer, sheet_name='Группы', index=False)
                pd.DataFrame(data.get('teachers', [])).to_excel(writer, sheet_name='Преподаватели', index=False)
                pd.DataFrame(data.get('classrooms', [])).to_excel(writer, sheet_name='Аудитории', index=False)
                pd.DataFrame(data.get('subjects', [])).to_excel(writer, sheet_name='Предметы', index=False)
                pd.DataFrame(data.get('holidays', [])).to_excel(writer, sheet_name='Праздники', index=False)

                # Отчеты (если есть подтвержденные занятия)
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

            messagebox.showinfo("Экспорт", f"Расписание успешно экспортировано в {save_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка экспорта: {str(e)}")                

    def create_groups_tab(self):
        # Фрейм для кнопок
        btn_frame = ttk.Frame(self.groups_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btn_frame, text="➕ Добавить группу", command=self.add_group).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="✏️ Редактировать", command=self.edit_group).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="🗑️ Удалить группу", command=self.delete_group).pack(side=tk.LEFT)
        # Таблица групп
        columns = ('ID', 'Название', 'Тип', 'Студенты', 'Курс', 'Специальность')
        self.groups_tree = ttk.Treeview(self.groups_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.groups_tree.heading(col, text=col)
            self.groups_tree.column(col, width=120)
        self.groups_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        # Прокрутка
        scrollbar = ttk.Scrollbar(self.groups_frame, orient=tk.VERTICAL, command=self.groups_tree.yview)
        self.groups_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def create_teachers_tab(self):
        # Фрейм для кнопок
        btn_frame = ttk.Frame(self.teachers_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btn_frame, text="➕ Добавить преподавателя", command=self.add_teacher).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="✏️ Редактировать", command=self.edit_teacher).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="🗑️ Удалить преподавателя", command=self.delete_teacher).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="🔄 Обновить стаж", command=self.update_all_experience).pack(side=tk.LEFT, padx=(10, 0))
        # Таблица преподавателей
        columns = ('ID', 'ФИО', 'Предметы', 'Макс. часов', 'Квалификация', 'Стаж')
        self.teachers_tree = ttk.Treeview(self.teachers_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.teachers_tree.heading(col, text=col)
            self.teachers_tree.column(col, width=150)
        self.teachers_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        # Прокрутка
        scrollbar = ttk.Scrollbar(self.teachers_frame, orient=tk.VERTICAL, command=self.teachers_tree.yview)
        self.teachers_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def create_classrooms_tab(self):
        # Фрейм для кнопок
        btn_frame = ttk.Frame(self.classrooms_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btn_frame, text="➕ Добавить аудиторию", command=self.add_classroom).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="✏️ Редактировать", command=self.edit_classroom).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="🗑️ Удалить аудиторию", command=self.delete_classroom).pack(side=tk.LEFT)
        # Таблица аудиторий
        columns = ('ID', 'Номер', 'Вместимость', 'Тип', 'Оборудование', 'Расположение')
        self.classrooms_tree = ttk.Treeview(self.classrooms_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.classrooms_tree.heading(col, text=col)
            self.classrooms_tree.column(col, width=150)
        self.classrooms_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        # Прокрутка
        scrollbar = ttk.Scrollbar(self.classrooms_frame, orient=tk.VERTICAL, command=self.classrooms_tree.yview)
        self.classrooms_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def create_subjects_tab(self):
        # Фрейм для кнопок
        btn_frame = ttk.Frame(self.subjects_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btn_frame, text="➕ Добавить предмет", command=self.add_subject).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="✏️ Редактировать", command=self.edit_subject).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="🗑️ Удалить предмет", command=self.delete_subject).pack(side=tk.LEFT)
        # Таблица предметов
        columns = ('ID', 'Название', 'Тип группы', 'Часов/неделю', 'Форма контроля', 'Кафедра')
        self.subjects_tree = ttk.Treeview(self.subjects_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.subjects_tree.heading(col, text=col)
            self.subjects_tree.column(col, width=150)
        self.subjects_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        # Прокрутка
        scrollbar = ttk.Scrollbar(self.subjects_frame, orient=tk.VERTICAL, command=self.subjects_tree.yview)
        self.subjects_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def create_schedule_tab(self):
        # Фрейм для фильтров
        filter_frame = ttk.Frame(self.schedule_frame)
        filter_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Группировка фильтров в логические блоки
        ttk.Label(filter_frame, text="Фильтры:", font=('Segoe UI', 11, 'bold')).pack(side=tk.LEFT, padx=(0, 20))

        # Неделя
        week_frame = ttk.Frame(filter_frame)
        week_frame.pack(side=tk.LEFT, padx=(0, 15))
        ttk.Label(week_frame, text="Неделя:").pack(side=tk.LEFT)
        week_combo = ttk.Combobox(week_frame, textvariable=self.week_var, width=8, state="readonly")
        week_combo['values'] = [f"Неделя {i}" for i in range(1, 13)]
        week_combo.current(0)
        week_combo.pack(side=tk.LEFT, padx=(5, 0))
        week_combo.bind('<<ComboboxSelected>>', self.filter_schedule)

        # Группа
        group_frame = ttk.Frame(filter_frame)
        group_frame.pack(side=tk.LEFT, padx=(0, 15))
        ttk.Label(group_frame, text="Группа:").pack(side=tk.LEFT)
        group_combo = ttk.Combobox(group_frame, textvariable=self.group_filter_var, width=12, state="readonly")
        group_combo.pack(side=tk.LEFT, padx=(5, 0))
        group_combo.bind('<<ComboboxSelected>>', self.filter_schedule)

        # Преподаватель
        teacher_frame = ttk.Frame(filter_frame)
        teacher_frame.pack(side=tk.LEFT, padx=(0, 15))
        ttk.Label(teacher_frame, text="Преподаватель:").pack(side=tk.LEFT)
        teacher_combo = ttk.Combobox(teacher_frame, textvariable=self.teacher_filter_var, width=12, state="readonly")
        teacher_combo.pack(side=tk.LEFT, padx=(5, 0))
        teacher_combo.bind('<<ComboboxSelected>>', self.filter_schedule)

        # Аудитория
        classroom_frame = ttk.Frame(filter_frame)
        classroom_frame.pack(side=tk.LEFT, padx=(0, 15))
        ttk.Label(classroom_frame, text="Аудитория:").pack(side=tk.LEFT)
        classroom_combo = ttk.Combobox(classroom_frame, textvariable=self.classroom_filter_var, width=12, state="readonly")
        classroom_combo.pack(side=tk.LEFT, padx=(5, 0))
        classroom_combo.bind('<<ComboboxSelected>>', self.filter_schedule)

        # Кнопка обновления
        ttk.Button(filter_frame, text="🔄 Обновить", command=self.filter_schedule).pack(side=tk.LEFT, padx=(10, 0))

        # Таблица расписания
        columns = ('Время', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье')
        self.schedule_tree = ttk.Treeview(self.schedule_frame, columns=columns, show='headings', height=20)

        # Настройка заголовков и ширины столбцов
        self.schedule_tree.column('Время', width=100, anchor='center')
        for day in columns[1:]:
            self.schedule_tree.column(day, width=220, anchor='center') # Увеличена ширина для лучшей читаемости

        for col in columns:
            self.schedule_tree.heading(col, text=col)

        self.schedule_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=(0, 10), pady=(0, 10))

        # Прокрутка
        scrollbar_y = ttk.Scrollbar(self.schedule_frame, orient=tk.VERTICAL, command=self.schedule_tree.yview)
        scrollbar_x = ttk.Scrollbar(self.schedule_frame, orient=tk.HORIZONTAL, command=self.schedule_tree.xview)
        self.schedule_tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Кнопки для расписания (используем акцентный стиль)
        schedule_buttons_frame = ttk.Frame(self.schedule_frame)
        schedule_buttons_frame.pack(fill=tk.X, pady=(10, 0))

        buttons = [
            ("➕ Добавить занятие", self.add_lesson),
            ("✏️ Редактировать", self.edit_lesson),
            ("🗑️ Удалить занятие", self.delete_lesson),
            ("🔄 Заменить занятие", self.substitute_lesson),
            ("📅 Календарь", self.show_calendar),
            ("🌐 Экспорт в HTML", self.export_to_html),
            ("⏱️ Найти свободное время", self.find_free_slot)  # Новая кнопка
        ]

        for text, command in buttons:
            btn = ttk.Button(schedule_buttons_frame, text=text, command=command)
            btn.pack(side=tk.LEFT, padx=(0, 5))

    def find_free_slot(self):
        """Найти ближайший свободный слот для выбранной группы, преподавателя или аудитории"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Найти свободное время")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Выберите тип:").pack(pady=5)
        search_type_var = tk.StringVar(value="group")
        ttk.Radiobutton(dialog, text="Группа", variable=search_type_var, value="group").pack()
        ttk.Radiobutton(dialog, text="Преподаватель", variable=search_type_var, value="teacher").pack()
        ttk.Radiobutton(dialog, text="Аудитория", variable=search_type_var, value="classroom").pack()

        ttk.Label(dialog, text="Выберите элемент:").pack(pady=5)
        element_var = tk.StringVar()
        element_combo = ttk.Combobox(dialog, textvariable=element_var, state="readonly")
        element_combo.pack()

        def update_combo():
            search_type = search_type_var.get()
            if search_type == "group":
                values = [g['name'] for g in self.groups]
            elif search_type == "teacher":
                values = [t['name'] for t in self.teachers]
            else:
                values = [c['name'] for c in self.classrooms]
            element_combo['values'] = values
            if values:
                element_var.set(values[0])

        search_type_var.trace('w', lambda *args: update_combo())
        update_combo()

        def search_slot():
            search_type = search_type_var.get()
            element_name = element_var.get()
            if not element_name:
                messagebox.showwarning("Предупреждение", "Выберите элемент")
                return

            # Фильтруем свободные слоты
            free_slots = self.schedule[self.schedule['status'] == 'свободно']
            if search_type == "group":
                free_slots = free_slots[free_slots['group_name'] == element_name]
            elif search_type == "teacher":
                # Для преподавателя ищем слоты, где он может преподавать
                free_slots = free_slots[free_slots['teacher_name'] == ''] # Упрощенно
            else:
                free_slots = free_slots[free_slots['classroom_name'] == '']

            if not free_slots.empty:
                first_slot = free_slots.iloc[0]
                messagebox.showinfo("Свободный слот", 
                    f"Ближайший свободный слот:\n"
                    f"Неделя: {first_slot['week']}\n"
                    f"День: {first_slot['day']}\n"
                    f"Время: {first_slot['time']}\n"
                    f"Группа: {first_slot['group_name']}")
            else:
                messagebox.showinfo("Информация", "Свободных слотов не найдено")
    
        ttk.Button(dialog, text="Найти", command=search_slot).pack(pady=20)

    def create_reports_tab(self):
        # Фреймы для отчетов
        reports_notebook = ttk.Notebook(self.reports_frame)
        reports_notebook.pack(fill=tk.BOTH, expand=True)
        # Отчет по нагрузке преподавателей
        teacher_report_frame = ttk.Frame(reports_notebook)
        reports_notebook.add(teacher_report_frame, text="👨‍🏫 Нагрузка преподавателей")
        columns = ('Преподаватель', 'Часы', 'Группы', 'Предметы')
        self.teacher_report_tree = ttk.Treeview(teacher_report_frame, columns=columns, show='headings')
        for col in columns:
            self.teacher_report_tree.heading(col, text=col)
            self.teacher_report_tree.column(col, width=200)
        self.teacher_report_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=10, pady=10)
        scrollbar = ttk.Scrollbar(teacher_report_frame, orient=tk.VERTICAL, command=self.teacher_report_tree.yview)
        self.teacher_report_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # Отчет по нагрузке групп
        group_report_frame = ttk.Frame(reports_notebook)
        reports_notebook.add(group_report_frame, text="👥 Нагрузка групп")
        columns = ('Группа', 'Часы', 'Предметы', 'Преподаватели')
        self.group_report_tree = ttk.Treeview(group_report_frame, columns=columns, show='headings')
        for col in columns:
            self.group_report_tree.heading(col, text=col)
            self.group_report_tree.column(col, width=200)
        self.group_report_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=10, pady=10)
        scrollbar = ttk.Scrollbar(group_report_frame, orient=tk.VERTICAL, command=self.group_report_tree.yview)
        self.group_report_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # Отчет по конфликтам
        conflicts_frame = ttk.Frame(reports_notebook)
        reports_notebook.add(conflicts_frame, text="⚠️ Конфликты")
        self.conflicts_text = tk.Text(conflicts_frame, wrap=tk.WORD)
        self.conflicts_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=10, pady=10)
        scrollbar = ttk.Scrollbar(conflicts_frame, orient=tk.VERTICAL, command=self.conflicts_text.yview)
        self.conflicts_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # Текстовый отчет
        summary_frame = ttk.Frame(reports_notebook)
        reports_notebook.add(summary_frame, text="📋 Сводка")
        self.summary_text = tk.Text(summary_frame, wrap=tk.WORD)
        self.summary_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=10, pady=10)
        scrollbar = ttk.Scrollbar(summary_frame, orient=tk.VERTICAL, command=self.summary_text.yview)
        self.summary_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # Теперь можно загрузить данные, так как все фреймы созданы
        self.load_groups_data()
        self.load_teachers_data()
        self.load_classrooms_data()
        self.load_subjects_data()

    def add_holiday(self):
        """Добавить праздник"""
        # Создание диалогового окна для добавления праздника
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить праздник")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        ttk.Label(dialog, text="Дата (ГГГГ-ММ-ДД):").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        date_entry = ttk.Entry(dialog, width=20)
        date_entry.grid(row=0, column=1, padx=10, pady=5)
        date_entry.insert(0, datetime.now().date().isoformat()) # Предзаполнение текущей датой
        ttk.Label(dialog, text="Название:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=1, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Тип:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        type_var = tk.StringVar(value="Государственный")
        type_combo = ttk.Combobox(dialog, textvariable=type_var, 
                                 values=["Государственный", "Учебный", "Каникулы"], state="readonly")
        type_combo.grid(row=2, column=1, padx=10, pady=5)
        def save_holiday():
            if date_entry.get() and name_entry.get():
                # Проверка формата даты (упрощенная)
                try:
                    datetime.strptime(date_entry.get(), '%Y-%m-%d')
                    self.holidays.append({
                        'date': date_entry.get(),
                        'name': name_entry.get(),
                        'type': type_var.get()
                    })
                    self.load_holidays_data()
                    dialog.destroy()
                    self.create_backup()
                except ValueError:
                    messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ГГГГ-ММ-ДД")
            else:
                messagebox.showwarning("Предупреждение", "Заполните все поля")
        ttk.Button(dialog, text="Сохранить", command=save_holiday).grid(row=3, column=0, columnspan=2, pady=20)

    def delete_holiday(self):
        """Удалить праздник"""
        selected = self.holidays_tree.selection()
        if selected:
            if messagebox.askyesno("Подтверждение", "Удалить выбранный праздник?"):
                item = self.holidays_tree.item(selected[0])
                holiday_date = item['values'][0]
                holiday_name = item['values'][1]
                # Находим и удаляем праздник по дате и названию (предполагаем, что они уникальны вместе)
                self.holidays = [h for h in self.holidays if not (h['date'] == holiday_date and h['name'] == holiday_name)]
                self.load_holidays_data()
                self.create_backup()
        else:
            messagebox.showinfo("Информация", "Выберите праздник для удаления")

    def create_holidays_tab(self):
        # Фрейм для кнопок
        btn_frame = ttk.Frame(self.holidays_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btn_frame, text="➕ Добавить праздник", command=self.add_holiday).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="🗑️ Удалить праздник", command=self.delete_holiday).pack(side=tk.LEFT)
        # Таблица праздников
        columns = ('Дата', 'Название', 'Тип')
        self.holidays_tree = ttk.Treeview(self.holidays_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.holidays_tree.heading(col, text=col)
            self.holidays_tree.column(col, width=200)
        self.holidays_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        # Прокрутка
        scrollbar = ttk.Scrollbar(self.holidays_frame, orient=tk.VERTICAL, command=self.holidays_tree.yview)
        self.holidays_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.load_holidays_data()

    def create_archive_tab(self):
        """Создание вкладки архива расписаний"""
        # Создаем директорию для архива, если ее нет
        self.archive_dir = "schedule_archive"
        if not os.path.exists(self.archive_dir):
            os.makedirs(self.archive_dir)

        # Фрейм для кнопок
        btn_frame = ttk.Frame(self.archive_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btn_frame, text="💾 Сохранить текущее расписание", command=self.save_current_schedule).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="📂 Загрузить расписание", command=self.load_archived_schedule).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="🗑️ Удалить", command=self.delete_archived_schedule).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="📤 Экспорт в Excel", command=self.export_archived_schedule).pack(side=tk.RIGHT, padx=(5, 0))

        # Таблица архивных расписаний
        columns = ('Имя файла', 'Дата создания', 'Группы', 'Преподаватели', 'Аудитории', 'Предметы', 'Занятий')
        self.archive_tree = ttk.Treeview(self.archive_frame, columns=columns, show='headings', height=20)
        
        # Настройка заголовков и ширины столбцов
        column_widths = {
            'Имя файла': 200,
            'Дата создания': 150,
            'Группы': 80,
            'Преподаватели': 80,
            'Аудитории': 80,
            'Предметы': 80,
            'Занятий': 80
        }
        for col in columns:
            self.archive_tree.heading(col, text=col)
            self.archive_tree.column(col, width=column_widths.get(col, 100))

        self.archive_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=10, pady=10)
        
        # Прокрутка
        scrollbar = ttk.Scrollbar(self.archive_frame, orient=tk.VERTICAL, command=self.archive_tree.yview)
        self.archive_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Загрузка списка архивных расписаний
        self.load_archive_list()

    def load_groups_data(self):
        self.groups_tree.delete(*self.groups_tree.get_children())
        for group in self.groups:
            self.groups_tree.insert('', tk.END, values=(
                group['id'], group['name'], group.get('type', 'основное'), 
                group.get('students', 0), group.get('course', ''), 
                group.get('specialty', '')
            ))
        # Обновление комбобокса фильтра групп (если schedule_frame уже создан)
        if hasattr(self, 'schedule_frame') and self.schedule_frame.winfo_children():
            group_names = [group['name'] for group in self.groups]
            self.group_filter_var.set('')
            # Найдем комбобокс групп в расписании
            for child in self.schedule_frame.winfo_children():
                if isinstance(child, ttk.Frame):  # filter_frame
                    for widget in child.winfo_children():
                        if isinstance(widget, ttk.Combobox) and 'group' in str(widget):
                            widget['values'] = [''] + group_names
                            break

    def load_teachers_data(self):
        self.teachers_tree.delete(*self.teachers_tree.get_children())
        for teacher in self.teachers:
            self.teachers_tree.insert('', tk.END, values=(
                teacher['id'], teacher['name'], teacher.get('subjects', ''), 
                teacher.get('max_hours', 0), teacher.get('qualification', ''), 
                teacher.get('experience', 0)
            ))
        # Обновление комбобокса фильтра преподавателей (если schedule_frame уже создан)
        if hasattr(self, 'schedule_frame') and self.schedule_frame.winfo_children():
            teacher_names = [teacher['name'] for teacher in self.teachers]
            self.teacher_filter_var.set('')
            # Найдем комбобокс преподавателей в расписании
            for child in self.schedule_frame.winfo_children():
                if isinstance(child, ttk.Frame):  # filter_frame
                    for widget in child.winfo_children():
                        if isinstance(widget, ttk.Combobox) and 'teacher' in str(widget):
                            widget['values'] = [''] + teacher_names
                            break

    def load_classrooms_data(self):
        self.classrooms_tree.delete(*self.classrooms_tree.get_children())
        for classroom in self.classrooms:
            self.classrooms_tree.insert('', tk.END, values=(
                classroom['id'], classroom['name'], classroom.get('capacity', 0), 
                classroom.get('type', 'обычная'), classroom.get('equipment', ''), 
                classroom.get('location', '')
            ))
        # Обновление комбобокса фильтра аудиторий (если schedule_frame уже создан)
        if hasattr(self, 'schedule_frame') and self.schedule_frame.winfo_children():
            classroom_names = [classroom['name'] for classroom in self.classrooms]
            self.classroom_filter_var.set('')
            # Найдем комбобокс аудиторий в расписании
            for child in self.schedule_frame.winfo_children():
                if isinstance(child, ttk.Frame):  # filter_frame
                    for widget in child.winfo_children():
                        if isinstance(widget, ttk.Combobox) and 'classroom' in str(widget):
                            widget['values'] = [''] + classroom_names
                            break

    def load_subjects_data(self):
        self.subjects_tree.delete(*self.subjects_tree.get_children())
        for subject in self.subjects:
            self.subjects_tree.insert('', tk.END, values=(
                subject['id'], subject['name'], subject.get('group_type', 'основное'), 
                subject.get('hours_per_week', 0), subject.get('assessment', ''), 
                subject.get('department', '')
            ))

    def load_holidays_data(self):
        self.holidays_tree.delete(*self.holidays_tree.get_children())
        for holiday in self.holidays:
            self.holidays_tree.insert('', tk.END, values=(
                holiday.get('date', ''), 
                holiday.get('name', ''), 
                holiday.get('type', 'Государственный')
            ))

    def add_group(self):
        # Создание диалогового окна для добавления группы
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить группу")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        ttk.Label(dialog, text="Название:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Тип:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        type_var = tk.StringVar(value="основное")
        type_combo = ttk.Combobox(dialog, textvariable=type_var, values=["основное", "углубленное", "вечернее"], state="readonly")
        type_combo.grid(row=1, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Студентов:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        students_var = tk.StringVar(value="25")
        students_spin = ttk.Spinbox(dialog, from_=1, to=100, textvariable=students_var, width=10)
        students_spin.grid(row=2, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Курс:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        course_var = tk.StringVar()
        course_entry = ttk.Entry(dialog, textvariable=course_var, width=30)
        course_entry.grid(row=3, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Специальность:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        specialty_var = tk.StringVar()
        specialty_entry = ttk.Entry(dialog, textvariable=specialty_var, width=30)
        specialty_entry.grid(row=4, column=1, padx=10, pady=5)
        def save_group():
            if name_entry.get():
                new_id = max([g['id'] for g in self.groups], default=0) + 1
                self.groups.append({
                    'id': new_id,
                    'name': name_entry.get(),
                    'type': type_var.get(),
                    'students': int(students_var.get()),
                    'course': course_var.get(),
                    'specialty': specialty_var.get()
                })
                self.load_groups_data()
                dialog.destroy()
                self.create_backup()
            else:
                messagebox.showwarning("Предупреждение", "Введите название группы")
        ttk.Button(dialog, text="Сохранить", command=save_group).grid(row=5, column=0, columnspan=2, pady=20)

    def edit_group(self):
        selected = self.groups_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите группу для редактирования")
            return
        item = self.groups_tree.item(selected[0])
        group_id = item['values'][0]
        group = next((g for g in self.groups if g['id'] == group_id), None)
        if not group:
            return
        # Создание диалогового окна для редактирования группы
        dialog = tk.Toplevel(self.root)
        dialog.title("Редактировать группу")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        ttk.Label(dialog, text="Название:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=5)
        name_entry.insert(0, group['name'])
        ttk.Label(dialog, text="Тип:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        type_var = tk.StringVar(value=group.get('type', 'основное'))
        type_combo = ttk.Combobox(dialog, textvariable=type_var, values=["основное", "углубленное", "вечернее"], state="readonly")
        type_combo.grid(row=1, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Студентов:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        students_var = tk.StringVar(value=str(group.get('students', 25)))
        students_spin = ttk.Spinbox(dialog, from_=1, to=100, textvariable=students_var, width=10)
        students_spin.grid(row=2, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Курс:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        course_var = tk.StringVar(value=group.get('course', ''))
        course_entry = ttk.Entry(dialog, textvariable=course_var, width=30)
        course_entry.grid(row=3, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Специальность:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        specialty_var = tk.StringVar(value=group.get('specialty', ''))
        specialty_entry = ttk.Entry(dialog, textvariable=specialty_var, width=30)
        specialty_entry.grid(row=4, column=1, padx=10, pady=5)
        def save_group():
            if name_entry.get():
                group['name'] = name_entry.get()
                group['type'] = type_var.get()
                group['students'] = int(students_var.get())
                group['course'] = course_var.get()
                group['specialty'] = specialty_var.get()
                self.load_groups_data()
                dialog.destroy()
                self.create_backup()
            else:
                messagebox.showwarning("Предупреждение", "Введите название группы")
        ttk.Button(dialog, text="Сохранить", command=save_group).grid(row=5, column=0, columnspan=2, pady=20)

    def delete_group(self):
        selected = self.groups_tree.selection()
        if selected:
            if messagebox.askyesno("Подтверждение", "Удалить выбранную группу?"):
                item = self.groups_tree.item(selected[0])
                group_id = item['values'][0]
                self.groups = [g for g in self.groups if g['id'] != group_id]
                self.load_groups_data()
                self.create_backup()
        else:
            messagebox.showinfo("Информация", "Выберите группу для удаления")

    def add_teacher(self):
        # Создание диалогового окна для добавления преподавателя
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить преподавателя")
        dialog.geometry("450x480")  # Увеличена высота для новых полей
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        ttk.Label(dialog, text="ФИО:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        name_entry = ttk.Entry(dialog, width=35)
        name_entry.grid(row=0, column=1, padx=10, pady=5)

        ttk.Label(dialog, text="Предметы (через запятую):").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        subjects_entry = ttk.Entry(dialog, width=35)
        subjects_entry.grid(row=1, column=1, padx=10, pady=5)

        ttk.Label(dialog, text="Макс. часов/неделю:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        max_hours_var = tk.StringVar(value="20")
        max_hours_spin = ttk.Spinbox(dialog, from_=1, to=50, textvariable=max_hours_var, width=10)
        max_hours_spin.grid(row=2, column=1, padx=10, pady=5)

        ttk.Label(dialog, text="Квалификация:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        qualification_var = tk.StringVar()
        qualification_entry = ttk.Entry(dialog, textvariable=qualification_var, width=35)
        qualification_entry.grid(row=3, column=1, padx=10, pady=5)

        ttk.Label(dialog, text="Стаж (лет):").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        experience_var = tk.StringVar(value="0")
        experience_spin = ttk.Spinbox(dialog, from_=0, to=50, textvariable=experience_var, width=10)
        experience_spin.grid(row=4, column=1, padx=10, pady=5)

        ttk.Label(dialog, text="Контакты:").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        contacts_var = tk.StringVar()
        contacts_entry = ttk.Entry(dialog, textvariable=contacts_var, width=35)
        contacts_entry.grid(row=5, column=1, padx=10, pady=5)

        # --- НОВЫЕ ПОЛЯ: Приоритеты и ограничения ---
        ttk.Label(dialog, text="Запрещенные дни:").grid(row=6, column=0, sticky=tk.W, padx=10, pady=5)
        forbidden_days_var = tk.StringVar()
        forbidden_days_entry = ttk.Entry(dialog, textvariable=forbidden_days_var, width=35)
        forbidden_days_entry.grid(row=6, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Пример: Понедельник, Среда").grid(row=7, column=1, sticky=tk.W, padx=10, pady=(0, 5))

        ttk.Label(dialog, text="Предпочтительные дни:").grid(row=8, column=0, sticky=tk.W, padx=10, pady=5)
        preferred_days_var = tk.StringVar()
        preferred_days_entry = ttk.Entry(dialog, textvariable=preferred_days_var, width=35)
        preferred_days_entry.grid(row=8, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Пример: Вторник, Четверг").grid(row=9, column=1, sticky=tk.W, padx=10, pady=(0, 5))

        ttk.Label(dialog, text="Макс. пар в день:").grid(row=10, column=0, sticky=tk.W, padx=10, pady=5)
        max_lessons_per_day_var = tk.StringVar(value="6")
        max_lessons_per_day_spin = ttk.Spinbox(dialog, from_=1, to=12, textvariable=max_lessons_per_day_var, width=10)
        max_lessons_per_day_spin.grid(row=10, column=1, padx=10, pady=5)
        # --- КОНЕЦ НОВЫХ ПОЛЕЙ ---

        def save_teacher():
            if name_entry.get():
                new_id = max([t['id'] for t in self.teachers], default=0) + 1
                self.teachers.append({
                    'id': new_id,
                    'name': name_entry.get(),
                    'subjects': subjects_entry.get(),
                    'max_hours': int(max_hours_var.get()),
                    'qualification': qualification_var.get(),
                    'experience': int(experience_var.get()),
                    'contacts': contacts_var.get(),
                    # --- НОВЫЕ ПОЛЯ ---
                    'forbidden_days': forbidden_days_var.get(),
                    'preferred_days': preferred_days_var.get(),
                    'max_lessons_per_day': int(max_lessons_per_day_var.get())
                    # --- КОНЕЦ НОВЫХ ПОЛЕЙ ---
                })
                self.load_teachers_data()
                dialog.destroy()
                self.create_backup()
            else:
                messagebox.showwarning("Предупреждение", "Введите ФИО преподавателя")

        ttk.Button(dialog, text="Сохранить", command=save_teacher).grid(row=11, column=0, columnspan=2, pady=20)    

    def edit_teacher(self):
        selected = self.teachers_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите преподавателя для редактирования")
            return
        item = self.teachers_tree.item(selected[0])
        teacher_id = item['values'][0]
        teacher = next((t for t in self.teachers if t['id'] == teacher_id), None)
        if not teacher:
            return
        # Создание диалогового окна для редактирования преподавателя
        dialog = tk.Toplevel(self.root)
        dialog.title("Редактировать преподавателя")
        dialog.geometry("450x520")  # Увеличена высота для новых полей
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        ttk.Label(dialog, text="ФИО:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        name_entry = ttk.Entry(dialog, width=35)
        name_entry.grid(row=0, column=1, padx=10, pady=5)
        name_entry.insert(0, teacher['name'])

        ttk.Label(dialog, text="Предметы (через запятую):").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        subjects_entry = ttk.Entry(dialog, width=35)
        subjects_entry.grid(row=1, column=1, padx=10, pady=5)
        subjects_entry.insert(0, teacher.get('subjects', ''))

        ttk.Label(dialog, text="Макс. часов/неделю:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        max_hours_var = tk.StringVar(value=str(teacher.get('max_hours', 20)))
        max_hours_spin = ttk.Spinbox(dialog, from_=1, to=50, textvariable=max_hours_var, width=10)
        max_hours_spin.grid(row=2, column=1, padx=10, pady=5)

        ttk.Label(dialog, text="Квалификация:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        qualification_var = tk.StringVar(value=teacher.get('qualification', ''))
        qualification_entry = ttk.Entry(dialog, textvariable=qualification_var, width=35)
        qualification_entry.grid(row=3, column=1, padx=10, pady=5)

        ttk.Label(dialog, text="Стаж (лет):").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        experience_var = tk.StringVar(value=str(teacher.get('experience', 0)))
        experience_spin = ttk.Spinbox(dialog, from_=0, to=50, textvariable=experience_var, width=10)
        experience_spin.grid(row=4, column=1, padx=10, pady=5)

        ttk.Label(dialog, text="Контакты:").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        contacts_var = tk.StringVar(value=teacher.get('contacts', ''))
        contacts_entry = ttk.Entry(dialog, textvariable=contacts_var, width=35)
        contacts_entry.grid(row=5, column=1, padx=10, pady=5)

        # --- НОВЫЕ ПОЛЯ: Приоритеты и ограничения ---
        ttk.Label(dialog, text="Запрещенные дни:").grid(row=6, column=0, sticky=tk.W, padx=10, pady=5)
        forbidden_days_var = tk.StringVar(value=teacher.get('forbidden_days', ''))
        forbidden_days_entry = ttk.Entry(dialog, textvariable=forbidden_days_var, width=35)
        forbidden_days_entry.grid(row=6, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Пример: Понедельник, Среда").grid(row=7, column=1, sticky=tk.W, padx=10, pady=(0, 5))

        ttk.Label(dialog, text="Предпочтительные дни:").grid(row=8, column=0, sticky=tk.W, padx=10, pady=5)
        preferred_days_var = tk.StringVar(value=teacher.get('preferred_days', ''))
        preferred_days_entry = ttk.Entry(dialog, textvariable=preferred_days_var, width=35)
        preferred_days_entry.grid(row=8, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Пример: Вторник, Четверг").grid(row=9, column=1, sticky=tk.W, padx=10, pady=(0, 5))

        ttk.Label(dialog, text="Макс. пар в день:").grid(row=10, column=0, sticky=tk.W, padx=10, pady=5)
        max_lessons_per_day_var = tk.StringVar(value=str(teacher.get('max_lessons_per_day', 6)))
        max_lessons_per_day_spin = ttk.Spinbox(dialog, from_=1, to=12, textvariable=max_lessons_per_day_var, width=10)
        max_lessons_per_day_spin.grid(row=10, column=1, padx=10, pady=5)
        # --- КОНЕЦ НОВЫХ ПОЛЕЙ ---

        def save_teacher():
            if name_entry.get():
                teacher['name'] = name_entry.get()
                teacher['subjects'] = subjects_entry.get()
                teacher['max_hours'] = int(max_hours_var.get())
                teacher['qualification'] = qualification_var.get()
                teacher['experience'] = int(experience_var.get())
                teacher['contacts'] = contacts_var.get()
                # --- НОВЫЕ ПОЛЯ ---
                teacher['forbidden_days'] = forbidden_days_var.get()
                teacher['preferred_days'] = preferred_days_var.get()
                teacher['max_lessons_per_day'] = int(max_lessons_per_day_var.get())
                # --- КОНЕЦ НОВЫХ ПОЛЕЙ ---
                self.load_teachers_data()
                dialog.destroy()
                self.create_backup()
            else:
                messagebox.showwarning("Предупреждение", "Введите ФИО преподавателя")

        ttk.Button(dialog, text="Сохранить", command=save_teacher).grid(row=11, column=0, columnspan=2, pady=20)

    def delete_teacher(self):
        selected = self.teachers_tree.selection()
        if selected:
            if messagebox.askyesno("Подтверждение", "Удалить выбранного преподавателя?"):
                item = self.teachers_tree.item(selected[0])
                teacher_id = item['values'][0]
                self.teachers = [t for t in self.teachers if t['id'] != teacher_id]
                self.load_teachers_data()
                self.create_backup()
        else:
            messagebox.showinfo("Информация", "Выберите преподавателя для удаления")

    def update_all_experience(self):
        """Обновить стаж всех преподавателей"""
        if messagebox.askyesno("Подтверждение", 
                              "Вы уверены, что хотите увеличить стаж всех преподавателей на 1 год?"):
            updated_count = 0
            for teacher in self.teachers:
                teacher['experience'] = teacher.get('experience', 0) + 1
                updated_count += 1
            self.load_teachers_data()
            self.create_backup()
            messagebox.showinfo("Успех", f"Стаж обновлен у {updated_count} преподавателей")

    def add_classroom(self):
        # Создание диалогового окна для добавления аудитории
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить аудиторию")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        ttk.Label(dialog, text="Номер:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Вместимость:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        capacity_var = tk.StringVar(value="30")
        capacity_spin = ttk.Spinbox(dialog, from_=1, to=200, textvariable=capacity_var, width=10)
        capacity_spin.grid(row=1, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Тип:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        type_var = tk.StringVar(value="обычная")
        type_combo = ttk.Combobox(dialog, textvariable=type_var, 
                                 values=["обычная", "компьютерный класс", "спортзал", "лаборатория", "специализированная"], 
                                 state="readonly")
        type_combo.grid(row=2, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Оборудование:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        equipment_var = tk.StringVar()
        equipment_entry = ttk.Entry(dialog, textvariable=equipment_var, width=30)
        equipment_entry.grid(row=3, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Расположение:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        location_var = tk.StringVar()
        location_entry = ttk.Entry(dialog, textvariable=location_var, width=30)
        location_entry.grid(row=4, column=1, padx=10, pady=5)
        def save_classroom():
            if name_entry.get():
                new_id = max([c['id'] for c in self.classrooms], default=0) + 1
                self.classrooms.append({
                    'id': new_id,
                    'name': name_entry.get(),
                    'capacity': int(capacity_var.get()),
                    'type': type_var.get(),
                    'equipment': equipment_var.get(),
                    'location': location_var.get()
                })
                self.load_classrooms_data()
                dialog.destroy()
                self.create_backup()
            else:
                messagebox.showwarning("Предупреждение", "Введите номер аудитории")
        ttk.Button(dialog, text="Сохранить", command=save_classroom).grid(row=5, column=0, columnspan=2, pady=20)

    def edit_classroom(self):
        selected = self.classrooms_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите аудиторию для редактирования")
            return
        item = self.classrooms_tree.item(selected[0])
        classroom_id = item['values'][0]
        classroom = next((c for c in self.classrooms if c['id'] == classroom_id), None)
        if not classroom:
            return
        # Создание диалогового окна для редактирования аудитории
        dialog = tk.Toplevel(self.root)
        dialog.title("Редактировать аудиторию")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        ttk.Label(dialog, text="Номер:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=5)
        name_entry.insert(0, classroom['name'])
        ttk.Label(dialog, text="Вместимость:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        capacity_var = tk.StringVar(value=str(classroom.get('capacity', 30)))
        capacity_spin = ttk.Spinbox(dialog, from_=1, to=200, textvariable=capacity_var, width=10)
        capacity_spin.grid(row=1, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Тип:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        type_var = tk.StringVar(value=classroom.get('type', 'обычная'))
        type_combo = ttk.Combobox(dialog, textvariable=type_var, 
                                 values=["обычная", "компьютерный класс", "спортзал", "лаборатория", "специализированная"], 
                                 state="readonly")
        type_combo.grid(row=2, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Оборудование:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        equipment_var = tk.StringVar(value=classroom.get('equipment', ''))
        equipment_entry = ttk.Entry(dialog, textvariable=equipment_var, width=30)
        equipment_entry.grid(row=3, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Расположение:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        location_var = tk.StringVar(value=classroom.get('location', ''))
        location_entry = ttk.Entry(dialog, textvariable=location_var, width=30)
        location_entry.grid(row=4, column=1, padx=10, pady=5)
        def save_classroom():
            if name_entry.get():
                classroom['name'] = name_entry.get()
                classroom['capacity'] = int(capacity_var.get())
                classroom['type'] = type_var.get()
                classroom['equipment'] = equipment_var.get()
                classroom['location'] = location_var.get()
                self.load_classrooms_data()
                dialog.destroy()
                self.create_backup()
            else:
                messagebox.showwarning("Предупреждение", "Введите номер аудитории")
        ttk.Button(dialog, text="Сохранить", command=save_classroom).grid(row=5, column=0, columnspan=2, pady=20)

    def delete_classroom(self):
        selected = self.classrooms_tree.selection()
        if selected:
            if messagebox.askyesno("Подтверждение", "Удалить выбранную аудиторию?"):
                item = self.classrooms_tree.item(selected[0])
                classroom_id = item['values'][0]
                self.classrooms = [c for c in self.classrooms if c['id'] != classroom_id]
                self.load_classrooms_data()
                self.create_backup()
        else:
            messagebox.showinfo("Информация", "Выберите аудиторию для удаления")

    def add_subject(self):
        # Создание диалогового окна для добавления предмета
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить предмет")
        dialog.geometry("450x350")
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        ttk.Label(dialog, text="Название:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        name_entry = ttk.Entry(dialog, width=35)
        name_entry.grid(row=0, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Тип группы:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        group_type_var = tk.StringVar(value="основное")
        group_type_combo = ttk.Combobox(dialog, textvariable=group_type_var, 
                                       values=["основное", "углубленное", "вечернее"], state="readonly")
        group_type_combo.grid(row=1, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Часов/неделю:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        hours_var = tk.StringVar(value="4")
        hours_spin = ttk.Spinbox(dialog, from_=1, to=20, textvariable=hours_var, width=10)
        hours_spin.grid(row=2, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Форма контроля:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        assessment_var = tk.StringVar(value="экзамен")
        assessment_combo = ttk.Combobox(dialog, textvariable=assessment_var, 
                                       values=["экзамен", "зачет", "зачет с оценкой", "курсовая работа", "дифференцированный зачет"], 
                                       state="readonly")
        assessment_combo.grid(row=3, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Кафедра:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        department_var = tk.StringVar()
        department_entry = ttk.Entry(dialog, textvariable=department_var, width=35)
        department_entry.grid(row=4, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Описание:").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        description_var = tk.StringVar()
        description_entry = ttk.Entry(dialog, textvariable=description_var, width=35)
        description_entry.grid(row=5, column=1, padx=10, pady=5)
        def save_subject():
            if name_entry.get():
                new_id = max([s['id'] for s in self.subjects], default=0) + 1
                self.subjects.append({
                    'id': new_id,
                    'name': name_entry.get(),
                    'group_type': group_type_var.get(),
                    'hours_per_week': int(hours_var.get()),
                    'assessment': assessment_var.get(),
                    'department': department_var.get(),
                    'description': description_var.get()
                })
                self.load_subjects_data()
                dialog.destroy()
                self.create_backup()
            else:
                messagebox.showwarning("Предупреждение", "Введите название предмета")
        ttk.Button(dialog, text="Сохранить", command=save_subject).grid(row=6, column=0, columnspan=2, pady=20)

    def edit_subject(self):
        selected = self.subjects_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите предмет для редактирования")
            return
        item = self.subjects_tree.item(selected[0])
        subject_id = item['values'][0]
        subject = next((s for s in self.subjects if s['id'] == subject_id), None)
        if not subject:
            return
        # Создание диалогового окна для редактирования предмета
        dialog = tk.Toplevel(self.root)
        dialog.title("Редактировать предмет")
        dialog.geometry("450x350")
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        ttk.Label(dialog, text="Название:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        name_entry = ttk.Entry(dialog, width=35)
        name_entry.grid(row=0, column=1, padx=10, pady=5)
        name_entry.insert(0, subject['name'])
        ttk.Label(dialog, text="Тип группы:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        group_type_var = tk.StringVar(value=subject.get('group_type', 'основное'))
        group_type_combo = ttk.Combobox(dialog, textvariable=group_type_var, 
                                       values=["основное", "углубленное", "вечернее"], state="readonly")
        group_type_combo.grid(row=1, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Часов/неделю:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        hours_var = tk.StringVar(value=str(subject.get('hours_per_week', 4)))
        hours_spin = ttk.Spinbox(dialog, from_=1, to=20, textvariable=hours_var, width=10)
        hours_spin.grid(row=2, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Форма контроля:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        assessment_var = tk.StringVar(value=subject.get('assessment', 'экзамен'))
        assessment_combo = ttk.Combobox(dialog, textvariable=assessment_var, 
                                       values=["экзамен", "зачет", "зачет с оценкой", "курсовая работа", "дифференцированный зачет"], 
                                       state="readonly")
        assessment_combo.grid(row=3, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Кафедра:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        department_var = tk.StringVar(value=subject.get('department', ''))
        department_entry = ttk.Entry(dialog, textvariable=department_var, width=35)
        department_entry.grid(row=4, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Описание:").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        description_var = tk.StringVar(value=subject.get('description', ''))
        description_entry = ttk.Entry(dialog, textvariable=description_var, width=35)
        description_entry.grid(row=5, column=1, padx=10, pady=5)
        def save_subject():
            if name_entry.get():
                subject['name'] = name_entry.get()
                subject['group_type'] = group_type_var.get()
                subject['hours_per_week'] = int(hours_var.get())
                subject['assessment'] = assessment_var.get()
                subject['department'] = department_var.get()
                subject['description'] = description_var.get()
                self.load_subjects_data()
                dialog.destroy()
                self.create_backup()
            else:
                messagebox.showwarning("Предупреждение", "Введите название предмета")
        ttk.Button(dialog, text="Сохранить", command=save_subject).grid(row=6, column=0, columnspan=2, pady=20)

    def delete_subject(self):
        selected = self.subjects_tree.selection()
        if selected:
            if messagebox.askyesno("Подтверждение", "Удалить выбранный предмет?"):
                item = self.subjects_tree.item(selected[0])
                subject_id = item['values'][0]
                self.subjects = [s for s in self.subjects if s['id'] != subject_id]
                self.load_subjects_data()
                self.create_backup()
        else:
            messagebox.showinfo("Информация", "Выберите предмет для удаления")

    def generate_schedule_thread(self):
        """Запуск генерации расписания в отдельном потоке"""
        self.progress.start()
        self.status_var.set("Генерация расписания...")
        threading.Thread(target=self.generate_schedule, daemon=True).start()

    def generate_schedule(self):
        """Генерация расписания с учетом праздников и настраиваемого расписания звонков"""
        try:
            # --- Добавлены проверки наличия обязательных данных ---
            if not hasattr(self, 'groups') or not self.groups:
                messagebox.showerror("Ошибка", "Необходимо добавить хотя бы одну группу")
                self.progress.stop()
                return
            if not hasattr(self, 'subjects') or not self.subjects:
                messagebox.showerror("Ошибка", "Необходимо добавить хотя бы один предмет")
                self.progress.stop()
                return
            if not hasattr(self, 'teachers') or not self.teachers:
                messagebox.showerror("Ошибка", "Необходимо добавить хотя бы одного преподавателя")
                self.progress.stop()
                return
            if not hasattr(self, 'classrooms') or not self.classrooms:
                messagebox.showerror("Ошибка", "Необходимо добавить хотя бы одну аудиторию")
                self.progress.stop()
                return
            # -------------------------------------------------------
            # Обновление настроек
            self.settings['days_per_week'] = int(self.days_var.get())
            self.settings['lessons_per_day'] = int(self.lessons_var.get())
            self.settings['weeks'] = int(self.weeks_var.get())
            
            # Получаем список праздничных дат
            holiday_dates = []
            for h in self.holidays:
                try:
                    holiday_dates.append(datetime.strptime(h['date'], '%Y-%m-%d').date())
                except ValueError:
                    # Пропускаем некорректные даты
                    continue

            # Создание структуры расписания, исключая праздники
            days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
            # Генерируем список всех дат в периоде расписания
            start_date = datetime.strptime(self.settings['start_date'], '%Y-%m-%d').date()
            all_dates = []
            for week in range(self.settings['weeks']):
                for day_index, day_name in enumerate(days):
                    current_date = start_date + timedelta(weeks=week, days=day_index)
                    # Проверяем, не является ли дата праздником
                    if current_date not in holiday_dates:
                        all_dates.append((week+1, day_name, current_date))

            # Создаем расписание только для рабочих дней
            schedule_data = []
            lesson_id = 1
            
            # Получаем настройки времени звонков, если они есть
            bell_schedule_str = self.settings.get('bell_schedule', '8:00-8:45,8:55-9:40,9:50-10:35,10:45-11:30,11:40-12:25,12:35-13:20')
            times = [slot.strip() for slot in bell_schedule_str.split(',')]  # Убираем лишние пробелы
            
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
            
            # Назначение предметов группам
            self.assign_subjects_to_groups()
            # Назначение преподавателей и аудиторий
            self.assign_teachers_and_classrooms()
            
            # Обновление интерфейса в основном потоке
            self.root.after(0, self.on_schedule_generated)
        
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Ошибка генерации расписания: {str(e)}"))
            self.root.after(0, self.progress.stop)

    def assign_subjects_to_groups(self):
        """Назначение предметов группам"""
        # --- Добавлена проверка на наличие предметов ---
        if not hasattr(self, 'subjects') or not self.subjects:
            print("Нет предметов для назначения.")
            return
        if not hasattr(self, 'groups') or not self.groups:
            print("Нет групп для назначения предметов.")
            return
        # -----------------------------------------------
        for group in self.groups:
            group_subjects = [s for s in self.subjects if s.get('group_type') in [group['type'], 'общий']]
            for subject in group_subjects:
                hours_per_week = subject.get('hours_per_week', 0)
                if hours_per_week <= 0:
                    continue
                # Находим свободные слоты для этой группы
                free_slots = self.schedule[
                    (self.schedule['group_id'] == group['id']) & 
                    (self.schedule['status'] == 'свободно')
                ].index
                # Выбираем случайные слоты для предмета
                if len(free_slots) >= hours_per_week:
                    selected_slots = random.sample(list(free_slots), hours_per_week)
                    for slot in selected_slots:
                        self.schedule.loc[slot, 'subject_id'] = subject['id']
                        self.schedule.loc[slot, 'subject_name'] = subject['name']
                        self.schedule.loc[slot, 'status'] = 'запланировано' 
                else:
                    print(f"Недостаточно свободных слотов для предмета {subject['name']} у группы {group['name']}")

    def assign_teachers_and_classrooms(self):
        """Назначение преподавателей и аудиторий"""
        # Для каждого запланированного занятия назначаем преподавателя и аудиторию
        for idx, lesson in self.schedule.iterrows():
            if lesson['status'] == 'запланировано':
                # Находим преподавателя по предмету
                subject = next((s for s in self.subjects if s['id'] == lesson['subject_id']), None)
                if subject:
                    available_teachers = [t for t in self.teachers if subject['name'] in t['subjects']]
                    if available_teachers:
                        # Проверяем нагрузку преподавателя
                        teacher = random.choice(available_teachers)
                        self.schedule.loc[idx, 'teacher_id'] = teacher['id']
                        self.schedule.loc[idx, 'teacher_name'] = teacher['name']
                        # Назначаем аудиторию
                        if self.classrooms:
                            classroom = random.choice(self.classrooms)
                            self.schedule.loc[idx, 'classroom_id'] = classroom['id']
                            self.schedule.loc[idx, 'classroom_name'] = classroom['name']
                        self.schedule.loc[idx, 'status'] = 'подтверждено'

    def on_schedule_generated(self):
        """Обработка завершения генерации расписания"""
        self.progress.stop()
        self.status_var.set("Расписание сгенерировано")
        messagebox.showinfo("Успех", "Расписание успешно сгенерировано!")
        self.filter_schedule()
        self.update_reports()
        self.create_backup()

    def filter_schedule(self, event=None):
        """Фильтрация расписания"""
        self.schedule_tree.delete(*self.schedule_tree.get_children())
        # Получение выбранных значений
        week_text = self.week_var.get()
        week_num = int(week_text.split()[1]) if week_text and "Неделя" in week_text else 1
        group_name = self.group_filter_var.get()
        teacher_name = self.teacher_filter_var.get()
        classroom_name = self.classroom_filter_var.get()
        # Фильтрация данных
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
            # Создание простой таблицы расписания
            if not filtered_schedule.empty:
                days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
                times = sorted(filtered_schedule['time'].unique())
                for time_slot in times:
                    row_data = [time_slot]
                    day_lessons = {}
                    # Получаем занятия для каждого дня
                    for day in days:
                        lesson = filtered_schedule[
                            (filtered_schedule['time'] == time_slot) & 
                            (filtered_schedule['day'] == day) & 
                            (filtered_schedule['status'] == 'подтверждено')
                        ]
                        if not lesson.empty:
                            lesson_info = lesson.iloc[0]
                            day_lessons[day] = f"{lesson_info['group_name']}\n{lesson_info['subject_name']}\n{lesson_info['teacher_name']}\n{lesson_info['classroom_name']}"
                        else:
                            day_lessons[day] = ""
                    # Добавляем данные для каждого дня
                    for day in days:
                        row_data.append(day_lessons.get(day, ""))
                    self.schedule_tree.insert('', tk.END, values=row_data)
            else:
                # Если нет данных, показываем пустую таблицу
                self.show_empty_schedule()
        else:
            # Если нет данных, показываем пустую таблицу
            self.show_empty_schedule()

    def show_empty_schedule(self):
        """Показать пустое расписание"""
        times = [f"{8+i}:00-{8+i}:45" for i in range(self.settings['lessons_per_day'])]
        days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
        for time_slot in times:
            row_data = [time_slot] + [''] * len(days)
            self.schedule_tree.insert('', tk.END, values=row_data)

    # --- НАЧАЛО РЕАЛИЗАЦИИ РУЧНОГО РЕЖИМА ---

    def add_lesson(self):
        """Добавить занятие в расписание вручную"""
        if not self.groups or not self.subjects or not self.teachers or not self.classrooms:
            messagebox.showwarning("Предупреждение", "Для добавления занятия необходимо, чтобы были созданы хотя бы одна группа, один предмет, один преподаватель и одна аудитория.")
            return

        # Создание диалогового окна
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить занятие")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        # Поля ввода
        ttk.Label(dialog, text="Неделя:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        week_var = tk.StringVar(value="1")
        week_combo = ttk.Combobox(dialog, textvariable=week_var, values=[str(i) for i in range(1, self.settings['weeks'] + 1)], state="readonly", width=5)
        week_combo.grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="День недели:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        day_var = tk.StringVar()
        days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
        day_combo = ttk.Combobox(dialog, textvariable=day_var, values=days, state="readonly", width=15)
        day_combo.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="Время:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        time_var = tk.StringVar()
        times = [f"{8+i}:00-{8+i}:45" for i in range(self.settings['lessons_per_day'])]
        time_combo = ttk.Combobox(dialog, textvariable=time_var, values=times, state="readonly", width=15)
        time_combo.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="Группа:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        group_var = tk.StringVar()
        group_names = [g['name'] for g in self.groups]
        group_combo = ttk.Combobox(dialog, textvariable=group_var, values=group_names, state="readonly", width=30)
        group_combo.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="Предмет:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        subject_var = tk.StringVar()
        subject_names = [s['name'] for s in self.subjects]
        subject_combo = ttk.Combobox(dialog, textvariable=subject_var, values=subject_names, state="readonly", width=30)
        subject_combo.grid(row=4, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="Преподаватель:").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        teacher_var = tk.StringVar()
        teacher_names = [t['name'] for t in self.teachers]
        teacher_combo = ttk.Combobox(dialog, textvariable=teacher_var, values=teacher_names, state="readonly", width=30)
        teacher_combo.grid(row=5, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="Аудитория:").grid(row=6, column=0, sticky=tk.W, padx=10, pady=5)
        classroom_var = tk.StringVar()
        classroom_names = [c['name'] for c in self.classrooms]
        classroom_combo = ttk.Combobox(dialog, textvariable=classroom_var, values=classroom_names, state="readonly", width=30)
        classroom_combo.grid(row=6, column=1, sticky=tk.W, padx=10, pady=5)

        def save_lesson():
            # Проверка, что все поля заполнены
            if not all([week_var.get(), day_var.get(), time_var.get(), group_var.get(), subject_var.get(), teacher_var.get(), classroom_var.get()]):
                messagebox.showwarning("Предупреждение", "Заполните все поля")
                return

            # Получение ID по именам
            selected_group = next((g for g in self.groups if g['name'] == group_var.get()), None)
            selected_subject = next((s for s in self.subjects if s['name'] == subject_var.get()), None)
            selected_teacher = next((t for t in self.teachers if t['name'] == teacher_var.get()), None)
            selected_classroom = next((c for c in self.classrooms if c['name'] == classroom_var.get()), None)

            if not all([selected_group, selected_subject, selected_teacher, selected_classroom]):
                messagebox.showerror("Ошибка", "Не удалось найти выбранные элементы в базе данных")
                return
            
            # --- НОВАЯ ПРОВЕРКА ОГРАНИЧЕНИЙ ---
            is_valid, error_message = self.validate_teacher_constraints(selected_teacher['id'], day_var.get(), int(week_var.get()))
            if not is_valid:
                messagebox.showerror("Ошибка ограничений", error_message)
                return
            # --- КОНЕЦ НОВОЙ ПРОВЕРКИ ---
            
            # Проверка, свободен ли слот
            existing_lesson = self.schedule[
                (self.schedule['week'] == int(week_var.get())) &
                (self.schedule['day'] == day_var.get()) &
                (self.schedule['time'] == time_var.get()) &
                (self.schedule['group_id'] == selected_group['id']) &
                (self.schedule['status'] != 'свободно')
            ]

            if not existing_lesson.empty:
                if not messagebox.askyesno("Подтверждение", "В выбранное время у этой группы уже есть занятие. Заменить его?"):
                    return

            # Поиск строки в DataFrame для обновления
            target_row = self.schedule[
                (self.schedule['week'] == int(week_var.get())) &
                (self.schedule['day'] == day_var.get()) &
                (self.schedule['time'] == time_var.get()) &
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
                messagebox.showinfo("Успех", "Занятие успешно добавлено!")
                dialog.destroy()
            else:
                messagebox.showerror("Ошибка", "Не удалось найти подходящий слот в расписании")

        ttk.Button(dialog, text="Сохранить", command=save_lesson).grid(row=7, column=0, columnspan=2, pady=20)

    def edit_lesson(self):
        """Редактировать выбранное занятие в расписании"""
        selected_item = self.schedule_tree.selection()
        if not selected_item:
            messagebox.showinfo("Информация", "Выберите занятие в таблице для редактирования")
            return

        # Получаем данные из строки таблицы
        values = self.schedule_tree.item(selected_item[0], 'values')
        time_slot = values[0]  # Первый столбец - время

        # Определяем текущую неделю и день из фильтров
        week_text = self.week_var.get()
        current_week = int(week_text.split()[1]) if week_text and "Неделя" in week_text else 1
        days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]

        # Находим день, в котором есть данные (не пустая строка)
        current_day = None
        for i, day in enumerate(days):
            if values[i + 1].strip():  # Пропускаем столбец времени (индекс 0)
                current_day = day
                break

        if not current_day:
            messagebox.showerror("Ошибка", "Не удалось определить день для выбранного занятия")
            return

        # Ищем запись в DataFrame
        lesson_row = self.schedule[
            (self.schedule['week'] == current_week) &
            (self.schedule['day'] == current_day) &
            (self.schedule['time'] == time_slot) &
            (self.schedule['status'] == 'подтверждено')
        ]

        if lesson_row.empty:
            messagebox.showerror("Ошибка", "Выбранное занятие не найдено в базе данных")
            return

        idx = lesson_row.index[0]
        lesson_data = self.schedule.loc[idx]

        # Создание диалогового окна (аналогично add_lesson, но с предзаполнением)
        dialog = tk.Toplevel(self.root)
        dialog.title("Редактировать занятие")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        ttk.Label(dialog, text="Неделя:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        week_var = tk.StringVar(value=str(current_week))
        week_combo = ttk.Combobox(dialog, textvariable=week_var, values=[str(i) for i in range(1, self.settings['weeks'] + 1)], state="readonly", width=5)
        week_combo.grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="День недели:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        day_var = tk.StringVar(value=current_day)
        days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
        day_combo = ttk.Combobox(dialog, textvariable=day_var, values=days, state="readonly", width=15)
        day_combo.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="Время:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        time_var = tk.StringVar(value=time_slot)
        times = [f"{8+i}:00-{8+i}:45" for i in range(self.settings['lessons_per_day'])]
        time_combo = ttk.Combobox(dialog, textvariable=time_var, values=times, state="readonly", width=15)
        time_combo.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="Группа:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        group_var = tk.StringVar(value=lesson_data['group_name'])
        group_names = [g['name'] for g in self.groups]
        group_combo = ttk.Combobox(dialog, textvariable=group_var, values=group_names, state="readonly", width=30)
        group_combo.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="Предмет:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        subject_var = tk.StringVar(value=lesson_data['subject_name'])
        subject_names = [s['name'] for s in self.subjects]
        subject_combo = ttk.Combobox(dialog, textvariable=subject_var, values=subject_names, state="readonly", width=30)
        subject_combo.grid(row=4, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="Преподаватель:").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        teacher_var = tk.StringVar(value=lesson_data['teacher_name'])
        teacher_names = [t['name'] for t in self.teachers]
        teacher_combo = ttk.Combobox(dialog, textvariable=teacher_var, values=teacher_names, state="readonly", width=30)
        teacher_combo.grid(row=5, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Label(dialog, text="Аудитория:").grid(row=6, column=0, sticky=tk.W, padx=10, pady=5)
        classroom_var = tk.StringVar(value=lesson_data['classroom_name'])
        classroom_names = [c['name'] for c in self.classrooms]
        classroom_combo = ttk.Combobox(dialog, textvariable=classroom_var, values=classroom_names, state="readonly", width=30)
        classroom_combo.grid(row=6, column=1, sticky=tk.W, padx=10, pady=5)

        def save_lesson():
            # Проверка, что все поля заполнены
            if not all([week_var.get(), day_var.get(), time_var.get(), group_var.get(), subject_var.get(), teacher_var.get(), classroom_var.get()]):
                messagebox.showwarning("Предупреждение", "Заполните все поля")
                return

            # Получение ID по именам
            selected_group = next((g for g in self.groups if g['name'] == group_var.get()), None)
            selected_subject = next((s for s in self.subjects if s['name'] == subject_var.get()), None)
            selected_teacher = next((t for t in self.teachers if t['name'] == teacher_var.get()), None)
            selected_classroom = next((c for c in self.classrooms if c['name'] == classroom_var.get()), None)

            if not all([selected_group, selected_subject, selected_teacher, selected_classroom]):
                messagebox.showerror("Ошибка", "Не удалось найти выбранные элементы в базе данных")
                return

            # Проверка, не создает ли редактирование конфликт (если группа/время меняются)
            if (group_var.get() != lesson_data['group_name'] or
                week_var.get() != str(current_week) or
                day_var.get() != current_day or
                time_var.get() != time_slot):

                conflict_check = self.schedule[
                    (self.schedule['week'] == int(week_var.get())) &
                    (self.schedule['day'] == day_var.get()) &
                    (self.schedule['time'] == time_var.get()) &
                    (self.schedule['group_id'] == selected_group['id']) &
                    (self.schedule['status'] != 'свободно')
                ]

                if not conflict_check.empty:
                    if not messagebox.askyesno("Подтверждение", "В выбранное время у этой группы уже есть занятие. Заменить его?"):
                        return

            # Обновляем данные в DataFrame
            self.schedule.loc[idx, 'week'] = int(week_var.get())
            self.schedule.loc[idx, 'day'] = day_var.get()
            self.schedule.loc[idx, 'time'] = time_var.get()
            self.schedule.loc[idx, 'group_id'] = selected_group['id']
            self.schedule.loc[idx, 'group_name'] = selected_group['name']
            self.schedule.loc[idx, 'subject_id'] = selected_subject['id']
            self.schedule.loc[idx, 'subject_name'] = selected_subject['name']
            self.schedule.loc[idx, 'teacher_id'] = selected_teacher['id']
            self.schedule.loc[idx, 'teacher_name'] = selected_teacher['name']
            self.schedule.loc[idx, 'classroom_id'] = selected_classroom['id']
            self.schedule.loc[idx, 'classroom_name'] = selected_classroom['name']

            self.filter_schedule()
            self.update_reports()
            self.create_backup()
            messagebox.showinfo("Успех", "Занятие успешно обновлено!")
            dialog.destroy()

        ttk.Button(dialog, text="Сохранить", command=save_lesson).grid(row=7, column=0, columnspan=2, pady=20)

    def delete_lesson(self):
        """Удалить выбранное занятие из расписания"""
        selected_item = self.schedule_tree.selection()
        if not selected_item:
            messagebox.showinfo("Информация", "Выберите занятие в таблице для удаления")
            return

        # Получаем данные из строки таблицы
        values = self.schedule_tree.item(selected_item[0], 'values')
        time_slot = values[0]

        # Определяем текущую неделю и день из фильтров
        week_text = self.week_var.get()
        current_week = int(week_text.split()[1]) if week_text and "Неделя" in week_text else 1
        days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]

        # Находим день, в котором есть данные
        current_day = None
        for i, day in enumerate(days):
            if values[i + 1].strip():
                current_day = day
                break

        if not current_day:
            messagebox.showerror("Ошибка", "Не удалось определить день для выбранного занятия")
            return

        # Ищем запись в DataFrame
        lesson_row = self.schedule[
            (self.schedule['week'] == current_week) &
            (self.schedule['day'] == current_day) &
            (self.schedule['time'] == time_slot) &
            (self.schedule['status'] == 'подтверждено')
        ]

        if lesson_row.empty:
            messagebox.showerror("Ошибка", "Выбранное занятие не найдено в базе данных")
            return

        if messagebox.askyesno("Подтверждение", "Вы уверены, что хотите удалить выбранное занятие?"):
            idx = lesson_row.index[0]
            # Очищаем данные и возвращаем статус в 'свободно'
            self.schedule.loc[idx, ['subject_id', 'subject_name', 'teacher_id', 'teacher_name', 'classroom_id', 'classroom_name']] = [None, '', None, '', None, '']
            self.schedule.loc[idx, 'status'] = 'свободно'

            self.filter_schedule()
            self.update_reports()
            self.create_backup()
            messagebox.showinfo("Успех", "Занятие успешно удалено!")

    # --- КОНЕЦ РЕАЛИЗАЦИИ РУЧНОГО РЕЖИМА ---

    def substitute_lesson(self):
        """Заменить преподавателя на занятии"""
        selected_item = self.schedule_tree.selection()
        if not selected_item:
            messagebox.showinfo("Информация", "Выберите занятие в таблице для замены")
            return
        # Получаем данные из строки таблицы
        values = self.schedule_tree.item(selected_item[0], 'values')
        time_slot = values[0]  # Первый столбец - время
        # Определяем текущую неделю и день из фильтров
        week_text = self.week_var.get()
        current_week = int(week_text.split()[1]) if week_text and "Неделя" in week_text else 1
        days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
        # Находим день, в котором есть данные
        current_day = None
        for i, day in enumerate(days):
            if values[i + 1].strip():
                current_day = day
                break
        if not current_day:
            messagebox.showerror("Ошибка", "Не удалось определить день для выбранного занятия")
            return
        # Ищем запись в DataFrame
        lesson_row = self.schedule[
            (self.schedule['week'] == current_week) &
            (self.schedule['day'] == current_day) &
            (self.schedule['time'] == time_slot) &
            (self.schedule['status'] == 'подтверждено')
        ]
        if lesson_row.empty:
            messagebox.showerror("Ошибка", "Выбранное занятие не найдено в базе данных")
            return
        idx = lesson_row.index[0]
        lesson_data = self.schedule.loc[idx]
        # Создание диалогового окна для замены
        dialog = tk.Toplevel(self.root)
        dialog.title("Замена занятия")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        # Информация о текущем занятии
        ttk.Label(dialog, text="Информация о занятии:", font=('Arial', 10, 'bold')).grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=(10, 5))
        ttk.Label(dialog, text=f"Группа: {lesson_data['group_name']}").grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=10)
        ttk.Label(dialog, text=f"Предмет: {lesson_data['subject_name']}").grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=10)
        ttk.Label(dialog, text=f"Оригинальный преподаватель: {lesson_data['teacher_name']}").grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=10)
        ttk.Label(dialog, text=f"Аудитория: {lesson_data['classroom_name']}").grid(row=4, column=0, columnspan=2, sticky=tk.W, padx=10)
        ttk.Label(dialog, text=f"Время: {time_slot}, {current_day}, Неделя {current_week}").grid(row=5, column=0, columnspan=2, sticky=tk.W, padx=10)
        ttk.Separator(dialog, orient='horizontal').grid(row=6, column=0, columnspan=2, sticky='ew', padx=10, pady=10)
        # Выбор заменяющего преподавателя
        ttk.Label(dialog, text="Заменяющий преподаватель:").grid(row=7, column=0, sticky=tk.W, padx=10, pady=5)
        replacement_teacher_var = tk.StringVar()
        teacher_names = [t['name'] for t in self.teachers if t['name'] != lesson_data['teacher_name']]  # Исключаем текущего
        if not teacher_names:
            messagebox.showwarning("Предупреждение", "Нет других преподавателей для замены")
            dialog.destroy()
            return
        teacher_combo = ttk.Combobox(dialog, textvariable=replacement_teacher_var, values=teacher_names, state="readonly", width=30)
        teacher_combo.grid(row=7, column=1, sticky=tk.W, padx=10, pady=5)
        # Причина замены
        ttk.Label(dialog, text="Причина замены:").grid(row=8, column=0, sticky=tk.W, padx=10, pady=5)
        reason_var = tk.StringVar()
        reason_combo = ttk.Combobox(dialog, textvariable=reason_var, 
                                   values=["Болезнь", "Отпуск", "Командировка", "Другое"], state="readonly", width=30)
        reason_combo.grid(row=8, column=1, sticky=tk.W, padx=10, pady=5)
        # Дополнительное поле для "Другое"
        ttk.Label(dialog, text="Детали (если 'Другое'):").grid(row=9, column=0, sticky=tk.W, padx=10, pady=5)
        details_entry = ttk.Entry(dialog, width=30)
        details_entry.grid(row=9, column=1, sticky=tk.W, padx=10, pady=5)
        # Флаг для отметки, что замена временная
        is_temporary_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(dialog, text="Временная замена (вернуть оригинального преподавателя позже)", 
                       variable=is_temporary_var).grid(row=10, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)
        def save_substitution():
            if not replacement_teacher_var.get():
                messagebox.showwarning("Предупреждение", "Выберите заменяющего преподавателя")
                return
            # Получаем объект нового преподавателя
            new_teacher = next((t for t in self.teachers if t['name'] == replacement_teacher_var.get()), None)
            if not new_teacher:
                messagebox.showerror("Ошибка", "Не удалось найти выбранного преподавателя")
                return
            # Собираем причину
            reason = reason_var.get()
            if reason == "Другое":
                reason = details_entry.get() or "Не указано"
            # Обновляем расписание
            original_teacher_id = self.schedule.loc[idx, 'teacher_id']
            original_teacher_name = self.schedule.loc[idx, 'teacher_name']
            self.schedule.loc[idx, 'teacher_id'] = new_teacher['id']
            self.schedule.loc[idx, 'teacher_name'] = new_teacher['name']
            # Записываем замену в журнал
            substitution_record = {
                'date': datetime.now().strftime('%Y-%m-%d'),
                'week': current_week,
                'day': current_day,
                'time': time_slot,
                'group': lesson_data['group_name'],
                'subject': lesson_data['subject_name'],
                'original_teacher': original_teacher_name,
                'replacement_teacher': new_teacher['name'],
                'reason': reason,
                'is_temporary': is_temporary_var.get(),
                'original_teacher_id': original_teacher_id,  # Для возможности отмены
                'schedule_index': int(idx)  # Сохраняем индекс строки в DataFrame для отмены
            }
            self.substitutions.append(substitution_record)
            # Обновляем интерфейс
            self.filter_schedule()
            self.create_backup()
            messagebox.showinfo("Успех", f"Занятие успешно заменено!\nНовый преподаватель: {new_teacher['name']}")
            dialog.destroy()
        ttk.Button(dialog, text="Подтвердить замену", command=save_substitution).grid(row=11, column=0, columnspan=2, pady=20)

    def show_calendar(self):
        """Показать календарь"""
        # Создание окна календаря
        dialog = tk.Toplevel(self.root)
        dialog.title("Календарь расписания")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        # Получаем текущую дату
        today = datetime.now()
        year = today.year
        month = today.month
        # Создаем фрейм для навигации
        nav_frame = ttk.Frame(dialog)
        nav_frame.pack(fill=tk.X, padx=10, pady=10)
        # Кнопки навигации
        prev_button = ttk.Button(nav_frame, text="◀ Назад", command=lambda: self.change_month(dialog, year, month-1))
        prev_button.pack(side=tk.LEFT)
        month_label = ttk.Label(nav_frame, text=f"{calendar.month_name[month]} {year}", font=('Arial', 14, 'bold'))
        month_label.pack(side=tk.LEFT, padx=20)
        next_button = ttk.Button(nav_frame, text="Вперед ▶", command=lambda: self.change_month(dialog, year, month+1))
        next_button.pack(side=tk.RIGHT)
        # Создаем таблицу календаря
        cal_frame = ttk.Frame(dialog)
        cal_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        # Заголовки дней недели
        days = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
        for i, day in enumerate(days):
            ttk.Label(cal_frame, text=day, font=('Arial', 10, 'bold')).grid(row=0, column=i, padx=5, pady=5)
        # Получаем календарь для текущего месяца
        cal = calendar.monthcalendar(year, month)
        # Заполняем календарь
        for week_num, week in enumerate(cal):
            for day_num, day in enumerate(week):
                if day == 0:
                    # Пустая ячейка
                    ttk.Label(cal_frame, text="", width=8, anchor='center').grid(row=week_num+1, column=day_num, padx=2, pady=2)
                else:
                    # Ячейка с датой
                    date_str = f"{year}-{month:02d}-{day:02d}"
                    # Проверяем, есть ли занятия в этот день
                    day_schedule = self.schedule[
                        (self.schedule['week'] == 1) &  # Для примера берем первую неделю
                        (self.schedule['day'] == days[day_num]) &
                        (self.schedule['status'] == 'подтверждено')
                    ]
                    if not day_schedule.empty:
                        # Есть занятия - выделяем цветом
                        day_button = tk.Button(cal_frame, text=str(day), width=8, 
                                              bg='#2ecc71', fg='white', 
                                              command=lambda d=date_str: self.show_day_schedule(d))
                    else:
                        # Нет занятий
                        day_button = tk.Button(cal_frame, text=str(day), width=8, 
                                              command=lambda d=date_str: self.show_day_schedule(d))
                    day_button.grid(row=week_num+1, column=day_num, padx=2, pady=2)

    def change_month(self, dialog, year, month):
        """Изменить месяц в календаре"""
        # Закрываем текущее окно
        dialog.destroy()
        # Открываем новое с новым месяцем
        self.show_calendar()

    def show_day_schedule(self, date):
        """Показать расписание на конкретный день"""
        messagebox.showinfo("Расписание", f"Расписание на {date} будет показано здесь")

    def check_conflicts(self):
        """Проверка конфликтов"""
        if self.schedule.empty:
            messagebox.showinfo("Информация", "Сначала сгенерируйте расписание")
            return
        # Проверка конфликтов преподавателей
        if 'teacher_id' in self.schedule.columns:
            teacher_conflicts = self.schedule[
                (self.schedule['status'] == 'подтверждено') &
                (self.schedule.duplicated(['teacher_id', 'day', 'time', 'week'], keep=False))
            ]
        else:
            teacher_conflicts = pd.DataFrame()
        # Проверка конфликтов аудиторий
        if 'classroom_id' in self.schedule.columns:
            classroom_conflicts = self.schedule[
                (self.schedule['status'] == 'подтверждено') &
                (self.schedule.duplicated(['classroom_id', 'day', 'time', 'week'], keep=False))
            ]
        else:
            classroom_conflicts = pd.DataFrame()
        # Проверка конфликтов групп
        if 'group_id' in self.schedule.columns:
            group_conflicts = self.schedule[
                (self.schedule['status'] == 'подтверждено') &
                (self.schedule.duplicated(['group_id', 'day', 'time', 'week'], keep=False))
            ]
        else:
            group_conflicts = pd.DataFrame()
        # Отображение результатов
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
        self.conflicts_text.delete(1.0, tk.END)
        self.conflicts_text.insert(1.0, conflict_text)
        messagebox.showinfo("Проверка конфликтов", 
                           f"Найдено конфликтов:\n"
                           f"Преподавателей: {len(teacher_conflicts)}\n"
                           f"Аудиторий: {len(classroom_conflicts)}\n"
                           f"Групп: {len(group_conflicts)}")

    def optimize_schedule(self):
        """Оптимизация расписания"""
        if self.schedule.empty:
            messagebox.showinfo("Информация", "Сначала сгенерируйте расписание")
            return
        self.progress.start()
        self.status_var.set("Оптимизация расписания...")
        # Простая оптимизация - переназначение конфликтных занятий
        if 'teacher_id' in self.schedule.columns:
            conflicts = self.schedule[
                (self.schedule['status'] == 'подтверждено') &
                (self.schedule.duplicated(['teacher_id', 'day', 'time', 'week'], keep=False))
            ]
        else:
            conflicts = pd.DataFrame()
        optimized_count = 0
        for idx in conflicts.index[:10]:  # Ограничиваем количество оптимизаций
            # Ищем свободный слот для этой группы в той же неделе
            group_id = self.schedule.loc[idx, 'group_id']
            week = self.schedule.loc[idx, 'week']
            free_slots = self.schedule[
                (self.schedule['group_id'] == group_id) & 
                (self.schedule['week'] == week) & 
                (self.schedule['status'] == 'свободно')
            ].index
            if len(free_slots) > 0:
                new_slot = random.choice(free_slots)
                # Копируем данные в новый слот
                self.schedule.loc[new_slot, 'subject_id'] = self.schedule.loc[idx, 'subject_id']
                self.schedule.loc[new_slot, 'subject_name'] = self.schedule.loc[idx, 'subject_name']
                self.schedule.loc[new_slot, 'teacher_id'] = self.schedule.loc[idx, 'teacher_id']
                self.schedule.loc[new_slot, 'teacher_name'] = self.schedule.loc[idx, 'teacher_name']
                self.schedule.loc[new_slot, 'classroom_id'] = self.schedule.loc[idx, 'classroom_id']
                self.schedule.loc[new_slot, 'classroom_name'] = self.schedule.loc[idx, 'classroom_name']
                self.schedule.loc[new_slot, 'status'] = 'подтверждено'
                # Освобождаем старый слот
                self.schedule.loc[idx, ['subject_id', 'subject_name', 'teacher_id', 'teacher_name', 
                                      'classroom_id', 'classroom_name']] = [None, '', None, '', None, '']
                self.schedule.loc[idx, 'status'] = 'свободно'
                optimized_count += 1
        self.progress.stop()
        self.status_var.set("Оптимизация завершена")
        messagebox.showinfo("Оптимизация", f"Оптимизировано {optimized_count} занятий")
        self.filter_schedule()
        self.update_reports()
        self.create_backup()

    def update_reports(self):
        """Обновление отчетов"""
        if self.schedule.empty:
            return
        # Очистка таблиц
        self.teacher_report_tree.delete(*self.teacher_report_tree.get_children())
        self.group_report_tree.delete(*self.group_report_tree.get_children())
        self.summary_text.delete(1.0, tk.END)
        # Нагрузка преподавателей
        if 'teacher_name' in self.schedule.columns and not self.schedule[self.schedule['status'] == 'подтверждено'].empty:
            teacher_load = self.schedule[self.schedule['status'] == 'подтверждено'].groupby('teacher_name').size()
            for teacher, hours in teacher_load.items():
                # Получаем группы и предметы преподавателя
                teacher_data = self.schedule[
                    (self.schedule['teacher_name'] == teacher) & 
                    (self.schedule['status'] == 'подтверждено')
                ]
                groups = ', '.join(teacher_data['group_name'].unique())
                subjects = ', '.join(teacher_data['subject_name'].unique())
                self.teacher_report_tree.insert('', tk.END, values=(teacher, hours, groups, subjects))
        # Нагрузка групп
        if 'group_name' in self.schedule.columns and not self.schedule[self.schedule['status'] == 'подтверждено'].empty:
            group_load = self.schedule[self.schedule['status'] == 'подтверждено'].groupby('group_name').size()
            for group, hours in group_load.items():
                # Получаем предметы и преподавателей группы
                group_data = self.schedule[
                    (self.schedule['group_name'] == group) & 
                    (self.schedule['status'] == 'подтверждено')
                ]
                subjects = ', '.join(group_data['subject_name'].unique())
                teachers = ', '.join(group_data['teacher_name'].unique())
                self.group_report_tree.insert('', tk.END, values=(group, hours, subjects, teachers))
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
        self.summary_text.insert(1.0, summary_text)

    def export_to_excel(self):
        """Экспорт в Excel"""
        if self.schedule.empty:
            messagebox.showinfo("Информация", "Сначала сгенерируйте расписание")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Сохранить расписание как"
        )
        if filename:
            try:
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    # Основное расписание
                    if not self.schedule.empty:
                        confirmed_schedule = self.schedule[self.schedule['status'] == 'подтверждено']
                        if not confirmed_schedule.empty:
                            # Создаем простую таблицу вместо сводной
                            confirmed_schedule.to_excel(writer, sheet_name='Расписание', index=False)
                        else:
                            # Если нет подтвержденных занятий, сохраняем все
                            self.schedule.to_excel(writer, sheet_name='Расписание', index=False)
                    else:
                        # Создаем пустую таблицу
                        pd.DataFrame(columns=['id', 'week', 'day', 'time', 'group_id', 'group_name', 
                                            'subject_id', 'subject_name', 'teacher_id', 'teacher_name', 
                                            'classroom_id', 'classroom_name', 'status']).to_excel(
                            writer, sheet_name='Расписание', index=False)
                    # Исходные данные
                    pd.DataFrame(self.groups).to_excel(writer, sheet_name='Группы', index=False)
                    pd.DataFrame(self.teachers).to_excel(writer, sheet_name='Преподаватели', index=False)
                    pd.DataFrame(self.classrooms).to_excel(writer, sheet_name='Аудитории', index=False)
                    pd.DataFrame(self.subjects).to_excel(writer, sheet_name='Предметы', index=False)
                    pd.DataFrame(self.holidays).to_excel(writer, sheet_name='Праздники', index=False)
                    # Отчеты
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
                messagebox.showinfo("Экспорт", f"Расписание успешно экспортировано в {filename}")
                self.create_backup()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка экспорта: {str(e)}")

    def export_to_html(self):
        """Экспорт расписания в HTML для размещения на сайте"""
        if self.schedule.empty:
            messagebox.showinfo("Информация", "Сначала сгенерируйте расписание")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
            title="Сохранить расписание как HTML"
        )
        if not filename:
            return

        try:
            # Фильтруем только подтвержденные занятия
            confirmed_schedule = self.schedule[self.schedule['status'] == 'подтверждено']
            if confirmed_schedule.empty:
                messagebox.showwarning("Предупреждение", "Нет подтвержденных занятий для экспорта")
                return

            # Получаем уникальные группы
            unique_groups = confirmed_schedule['group_name'].unique()
            unique_groups.sort()

            # Генерация HTML-контента
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

            # Генерируем таблицу для каждой группы
            for group_name in unique_groups:
                group_schedule = confirmed_schedule[confirmed_schedule['group_name'] == group_name]
                
                # Создаем структуру расписания для группы
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

            # Добавляем футер
            html_content += f"""
    </div>

    <div class="footer">
        <p>Расписание сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}</p>
        <p>Директор: {self.settings.get('director', 'Не указан')}</p>
    </div>
</body>
</html>
"""

            # Сохраняем HTML-файл
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(html_content)

            messagebox.showinfo("Экспорт", f"Расписание успешно экспортировано в HTML-файл:\n{filename}")
            self.create_backup() # Создаем бэкап после экспорта

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка экспорта в HTML: {str(e)}")

    def save_data(self):
        """Сохранение данных в файл"""
        data = {
            'settings': self.settings,
            'groups': self.groups,
            'teachers': self.teachers,
            'classrooms': self.classrooms,
            'subjects': self.subjects,
            'substitutions': self.substitutions,
            'holidays': self.holidays
        }
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Сохранить данные как"
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("Сохранение", "Данные успешно сохранены!")
                self.create_backup() # Создаем бэкап после сохранения
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка сохранения: {str(e)}")

    def load_data(self):
        """Загрузка данных из файла"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Загрузить данные"
        )
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
                # Обновление интерфейса
                self.load_groups_data()
                self.load_teachers_data()
                self.load_classrooms_data()
                self.load_subjects_data()
                self.load_holidays_data()
                messagebox.showinfo("Загрузка", "Данные успешно загружены")
                self.create_backup()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка загрузки: {str(e)}")

    def open_settings(self):
        """Открыть настройки приложения"""
        # Создание диалогового окна для настроек
        dialog = tk.Toplevel(self.root)
        dialog.title("Настройки приложения")
        dialog.geometry("550x600")  # Увеличил ширину для лучшего размещения
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Создаем основной фрейм с прокруткой (если нужно)
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === ОСНОВНЫЕ НАСТРОЙКИ ===
        basic_frame = ttk.LabelFrame(main_frame, text="Основные настройки", padding="10")
        basic_frame.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(basic_frame, text="Название учреждения:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        school_name_var = tk.StringVar(value=self.settings.get('school_name', ''))
        school_name_entry = ttk.Entry(basic_frame, textvariable=school_name_var, width=40)
        school_name_entry.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(basic_frame, text="Директор:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        director_var = tk.StringVar(value=self.settings.get('director', ''))
        director_entry = ttk.Entry(basic_frame, textvariable=director_var, width=40)
        director_entry.grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Label(basic_frame, text="Учебный год:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        academic_year_var = tk.StringVar(value=self.settings.get('academic_year', ''))
        academic_year_entry = ttk.Entry(basic_frame, textvariable=academic_year_var, width=40)
        academic_year_entry.grid(row=2, column=1, padx=5, pady=2)
        
        ttk.Label(basic_frame, text="Дата начала года:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        start_date_var = tk.StringVar(value=self.settings.get('start_date', datetime.now().date().isoformat()))
        start_date_entry = ttk.Entry(basic_frame, textvariable=start_date_var, width=40)
        start_date_entry.grid(row=3, column=1, padx=5, pady=2)
        
        # === ПАРАМЕТРЫ РАСПИСАНИЯ ===
        schedule_frame = ttk.LabelFrame(main_frame, text="Параметры расписания", padding="10")
        schedule_frame.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(schedule_frame, text="Дней в неделю:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        days_per_week_var = tk.StringVar(value=str(self.settings.get('days_per_week', 5)))
        days_per_week_spin = ttk.Spinbox(schedule_frame, from_=1, to=7, textvariable=days_per_week_var, width=10)
        days_per_week_spin.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(schedule_frame, text="Занятий в день:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        lessons_per_day_var = tk.StringVar(value=str(self.settings.get('lessons_per_day', 6)))
        lessons_per_day_spin = ttk.Spinbox(schedule_frame, from_=1, to=12, textvariable=lessons_per_day_var, width=10)
        lessons_per_day_spin.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(schedule_frame, text="Недель:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        weeks_var = tk.StringVar(value=str(self.settings.get('weeks', 2)))
        weeks_spin = ttk.Spinbox(schedule_frame, from_=1, to=52, textvariable=weeks_var, width=10)
        weeks_spin.grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)
        
        # === РАСПИСАНИЕ ЗВОНКОВ ===
        bell_frame = ttk.LabelFrame(main_frame, text="Расписание звонков", padding="10")
        bell_frame.grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(bell_frame, text="Текущее расписание:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        
        # Текстовое поле для отображения текущего расписания (только для чтения)
        bell_schedule_var = tk.StringVar(value=self.settings.get('bell_schedule', '8:00-8:45,8:55-9:40,9:50-10:35,10:45-11:30,11:40-12:25,12:35-13:20'))
        bell_schedule_display = ttk.Entry(bell_frame, textvariable=bell_schedule_var, width=50, state='readonly')
        bell_schedule_display.grid(row=1, column=0, columnspan=2, padx=5, pady=2, sticky=tk.W)
        
        # Кнопка для открытия редактора
        open_editor_btn = ttk.Button(bell_frame, text="⚙️ Открыть редактор...", 
                                   command=lambda: self.open_bell_schedule_editor(bell_schedule_var, dialog))
        open_editor_btn.grid(row=2, column=0, columnspan=2, pady=(5, 2), padx=5, sticky=tk.W)
        
        # Подпись под кнопкой
        ttk.Label(bell_frame, text="Редактор расписания звонков", font=('Segoe UI', 9, 'italic')).grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=5, pady=(0, 2))
        
        # === НАСТРОЙКИ АВТО-БЭКАПА ===
        backup_frame = ttk.LabelFrame(main_frame, text="Настройки авто-бэкапа", padding="10")
        backup_frame.grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(backup_frame, text="Авто-бэкап:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        auto_backup_var = tk.BooleanVar(value=self.settings.get('auto_backup', True))
        auto_backup_check = ttk.Checkbutton(backup_frame, variable=auto_backup_var)
        auto_backup_check.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(backup_frame, text="Интервал бэкапа (мин):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        backup_interval_var = tk.StringVar(value=str(self.settings.get('backup_interval', 30)))
        backup_interval_spin = ttk.Spinbox(backup_frame, from_=1, to=1440, textvariable=backup_interval_var, width=10)
        backup_interval_spin.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(backup_frame, text="Макс. бэкапов:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        max_backups_var = tk.StringVar(value=str(self.settings.get('max_backups', 10)))
        max_backups_spin = ttk.Spinbox(backup_frame, from_=1, to=100, textvariable=max_backups_var, width=10)
        max_backups_spin.grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)
        
        def save_settings():
            self.settings['school_name'] = school_name_var.get()
            self.settings['director'] = director_var.get()
            self.settings['academic_year'] = academic_year_var.get()
            self.settings['start_date'] = start_date_var.get()
            self.settings['days_per_week'] = int(days_per_week_var.get())
            self.settings['lessons_per_day'] = int(lessons_per_day_var.get())
            self.settings['weeks'] = int(weeks_var.get())
            self.settings['auto_backup'] = auto_backup_var.get()
            self.settings['backup_interval'] = int(backup_interval_var.get())
            self.settings['max_backups'] = int(max_backups_var.get())
            # Обновляем переменные в основном окно
            self.days_var.set(days_per_week_var.get())
            self.lessons_var.set(lessons_per_day_var.get())
            self.weeks_var.set(weeks_var.get())
            # Перезапуск таймера бэкапа
            self.restart_auto_backup()
            # Обновляем индикатор бэкапа
            self.update_backup_indicator()
            dialog.destroy()
        
        # Кнопка сохранения
        save_button = ttk.Button(main_frame, text="Сохранить", command=save_settings)
        save_button.grid(row=4, column=0, pady=20, padx=5, sticky=tk.E)
        
    def open_bell_schedule_editor(self, bell_schedule_var, parent_dialog):
        """Открывает редактор расписания звонков"""
        current_schedule = bell_schedule_var.get()
        editor = BellScheduleEditor(self.root, current_schedule)
        self.root.wait_window(editor.dialog)  # Ждем закрытия диалога редактора
        
        if editor.result is not None:
            # Обновляем переменную и, соответственно, поле в настройках
            bell_schedule_var.set(editor.result)
            # Также обновляем настройки приложения
            self.settings['bell_schedule'] = editor.result

    def show_reports(self):
        """Показать отчеты"""
        # Переключаемся на вкладку отчетов
        notebook = self.reports_frame.master
        notebook.select(self.reports_frame)
        self.update_reports()

    def open_substitutions(self):
        """Открыть журнал замен"""
        # Создание окна для журнала замен
        dialog = tk.Toplevel(self.root)
        dialog.title("Журнал замен")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        # Создание таблицы замен
        columns = ('Дата', 'Группа', 'Пара', 'Предмет', 'Преподаватель', 'Замена', 'Причина')
        self.substitutions_tree = ttk.Treeview(dialog, columns=columns, show='headings', height=20)
        for col in columns:
            self.substitutions_tree.heading(col, text=col)
            self.substitutions_tree.column(col, width=120)
        self.substitutions_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=10, pady=10)
        # Прокрутка
        scrollbar = ttk.Scrollbar(dialog, orient=tk.VERTICAL, command=self.substitutions_tree.yview)
        self.substitutions_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # Кнопки
        buttons_frame = ttk.Frame(dialog)
        buttons_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Button(buttons_frame, text="➕ Добавить замену", command=self.add_substitution).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="✏️ Редактировать", command=self.edit_substitution).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="🗑️ Удалить", command=self.delete_substitution).pack(side=tk.LEFT)
        ttk.Button(buttons_frame, text="📤 Экспорт", command=self.export_substitutions).pack(side=tk.RIGHT)
        # Загрузка данных
        self.load_substitutions_data()

    def load_substitutions_data(self):
        """Загрузка данных журнала замен"""
        self.substitutions_tree.delete(*self.substitutions_tree.get_children())
        for substitution in self.substitutions:
            self.substitutions_tree.insert('', tk.END, values=(
                substitution.get('date', ''),
                substitution.get('group', ''),
                substitution.get('lesson', ''),
                substitution.get('subject', ''),
                substitution.get('teacher', ''),
                substitution.get('replacement', ''),
                substitution.get('reason', '')
            ))

    def add_substitution(self):
        """Добавить замену в журнал вручную"""
        if not self.teachers or len(self.teachers) < 2:
            messagebox.showwarning("Предупреждение", "Необходимо наличие как минимум двух преподавателей")
            return
        # Создание диалогового окна
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить замену")
        dialog.geometry("500x500")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        ttk.Label(dialog, text="Дата замены (ГГГГ-ММ-ДД):").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        date_entry = ttk.Entry(dialog, textvariable=date_var, width=30)
        date_entry.grid(row=0, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Неделя:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        week_var = tk.StringVar(value="1")
        week_combo = ttk.Combobox(dialog, textvariable=week_var, values=[str(i) for i in range(1, self.settings['weeks'] + 1)], state="readonly", width=5)
        week_combo.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="День недели:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        day_var = tk.StringVar()
        days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
        day_combo = ttk.Combobox(dialog, textvariable=day_var, values=days, state="readonly", width=15)
        day_combo.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Время:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        time_var = tk.StringVar()
        times = [f"{8+i}:00-{8+i}:45" for i in range(self.settings['lessons_per_day'])]
        time_combo = ttk.Combobox(dialog, textvariable=time_var, values=times, state="readonly", width=15)
        time_combo.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Группа:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        group_var = tk.StringVar()
        group_names = [g['name'] for g in self.groups]
        group_combo = ttk.Combobox(dialog, textvariable=group_var, values=group_names, state="readonly", width=30)
        group_combo.grid(row=4, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Предмет:").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        subject_var = tk.StringVar()
        subject_names = [s['name'] for s in self.subjects]
        subject_combo = ttk.Combobox(dialog, textvariable=subject_var, values=subject_names, state="readonly", width=30)
        subject_combo.grid(row=5, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Оригинальный преподаватель:").grid(row=6, column=0, sticky=tk.W, padx=10, pady=5)
        original_teacher_var = tk.StringVar()
        teacher_names = [t['name'] for t in self.teachers]
        original_teacher_combo = ttk.Combobox(dialog, textvariable=original_teacher_var, values=teacher_names, state="readonly", width=30)
        original_teacher_combo.grid(row=6, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Заменяющий преподаватель:").grid(row=7, column=0, sticky=tk.W, padx=10, pady=5)
        replacement_teacher_var = tk.StringVar()
        replacement_teacher_combo = ttk.Combobox(dialog, textvariable=replacement_teacher_var, values=teacher_names, state="readonly", width=30)
        replacement_teacher_combo.grid(row=7, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Причина замены:").grid(row=8, column=0, sticky=tk.W, padx=10, pady=5)
        reason_var = tk.StringVar()
        reason_combo = ttk.Combobox(dialog, textvariable=reason_var, 
                                   values=["Болезнь", "Отпуск", "Командировка", "Другое"], state="readonly", width=30)
        reason_combo.grid(row=8, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Детали (если 'Другое'):").grid(row=9, column=0, sticky=tk.W, padx=10, pady=5)
        details_entry = ttk.Entry(dialog, width=30)
        details_entry.grid(row=9, column=1, sticky=tk.W, padx=10, pady=5)
        is_temporary_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(dialog, text="Временная замена", variable=is_temporary_var).grid(row=10, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)
        def save_manual_substitution():
            # Проверка заполнения обязательных полей
            required_fields = [
                date_var.get(), week_var.get(), day_var.get(), time_var.get(),
                group_var.get(), subject_var.get(), original_teacher_var.get(), replacement_teacher_var.get()
            ]
            if not all(required_fields):
                messagebox.showwarning("Предупреждение", "Заполните все обязательные поля")
                return
            if original_teacher_var.get() == replacement_teacher_var.get():
                messagebox.showwarning("Предупреждение", "Оригинальный и заменяющий преподаватели не могут быть одинаковыми")
                return
            # Собираем причину
            reason = reason_var.get()
            if reason == "Другое":
                reason = details_entry.get() or "Не указано"
            # Находим ID преподавателей
            original_teacher = next((t for t in self.teachers if t['name'] == original_teacher_var.get()), None)
            replacement_teacher = next((t for t in self.teachers if t['name'] == replacement_teacher_var.get()), None)
            if not original_teacher or not replacement_teacher:
                messagebox.showerror("Ошибка", "Не удалось найти преподавателей")
                return
            # Создаем запись
            substitution_record = {
                'date': date_var.get(),
                'week': int(week_var.get()),
                'day': day_var.get(),
                'time': time_var.get(),
                'group': group_var.get(),
                'subject': subject_var.get(),
                'original_teacher': original_teacher_var.get(),
                'replacement_teacher': replacement_teacher_var.get(),
                'reason': reason,
                'is_temporary': is_temporary_var.get(),
                'original_teacher_id': original_teacher['id'],
                'schedule_index': -1  # -1 означает, что замена не привязана к конкретной строке в текущем расписании
            }
            self.substitutions.append(substitution_record)
            # Пытаемся найти и обновить занятие в текущем расписании
            target_lesson = self.schedule[
                (self.schedule['week'] == int(week_var.get())) &
                (self.schedule['day'] == day_var.get()) &
                (self.schedule['time'] == time_var.get()) &
                (self.schedule['group_name'] == group_var.get()) &
                (self.schedule['subject_name'] == subject_var.get()) &
                (self.schedule['teacher_name'] == original_teacher_var.get()) &
                (self.schedule['status'] == 'подтверждено')
            ]
            if not target_lesson.empty:
                idx = target_lesson.index[0]
                self.schedule.loc[idx, 'teacher_id'] = replacement_teacher['id']
                self.schedule.loc[idx, 'teacher_name'] = replacement_teacher['name']
                # Обновляем индекс в записи замены
                self.substitutions[-1]['schedule_index'] = int(idx)
                self.filter_schedule()
            self.load_substitutions_data()
            self.create_backup()
            messagebox.showinfo("Успех", "Замена успешно добавлена в журнал!")
            dialog.destroy()
        ttk.Button(dialog, text="Сохранить", command=save_manual_substitution).grid(row=11, column=0, columnspan=2, pady=20)

    def edit_substitution(self):
        """Редактировать выбранную замену в журнале"""
        selected = self.substitutions_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите замену для редактирования")
            return
        item = self.substitutions_tree.item(selected[0])
        values = item['values']
        # Находим запись в списке по дате, группе и времени (это не идеально, но для примера подойдет)
        substitution_to_edit = None
        for sub in self.substitutions:
            if (sub.get('date') == values[0] and
                sub.get('group') == values[1] and
                sub.get('time') == values[2]):
                substitution_to_edit = sub
                break
        if not substitution_to_edit:
            messagebox.showerror("Ошибка", "Замена не найдена")
            return
        # Создание диалогового окна
        dialog = tk.Toplevel(self.root)
        dialog.title("Редактировать замену")
        dialog.geometry("500x500")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        ttk.Label(dialog, text="Дата замены (ГГГГ-ММ-ДД):").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        date_var = tk.StringVar(value=substitution_to_edit.get('date', ''))
        date_entry = ttk.Entry(dialog, textvariable=date_var, width=30)
        date_entry.grid(row=0, column=1, padx=10, pady=5)
        ttk.Label(dialog, text="Неделя:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        week_var = tk.StringVar(value=str(substitution_to_edit.get('week', 1)))
        week_combo = ttk.Combobox(dialog, textvariable=week_var, values=[str(i) for i in range(1, self.settings['weeks'] + 1)], state="readonly", width=5)
        week_combo.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="День недели:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        day_var = tk.StringVar(value=substitution_to_edit.get('day', ''))
        days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'][:self.settings['days_per_week']]
        day_combo = ttk.Combobox(dialog, textvariable=day_var, values=days, state="readonly", width=15)
        day_combo.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Время:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        time_var = tk.StringVar(value=substitution_to_edit.get('time', ''))
        times = [f"{8+i}:00-{8+i}:45" for i in range(self.settings['lessons_per_day'])]
        time_combo = ttk.Combobox(dialog, textvariable=time_var, values=times, state="readonly", width=15)
        time_combo.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Группа:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        group_var = tk.StringVar(value=substitution_to_edit.get('group', ''))
        group_names = [g['name'] for g in self.groups]
        group_combo = ttk.Combobox(dialog, textvariable=group_var, values=group_names, state="readonly", width=30)
        group_combo.grid(row=4, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Предмет:").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        subject_var = tk.StringVar(value=substitution_to_edit.get('subject', ''))
        subject_names = [s['name'] for s in self.subjects]
        subject_combo = ttk.Combobox(dialog, textvariable=subject_var, values=subject_names, state="readonly", width=30)
        subject_combo.grid(row=5, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Оригинальный преподаватель:").grid(row=6, column=0, sticky=tk.W, padx=10, pady=5)
        original_teacher_var = tk.StringVar(value=substitution_to_edit.get('original_teacher', ''))
        teacher_names = [t['name'] for t in self.teachers]
        original_teacher_combo = ttk.Combobox(dialog, textvariable=original_teacher_var, values=teacher_names, state="readonly", width=30)
        original_teacher_combo.grid(row=6, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Заменяющий преподаватель:").grid(row=7, column=0, sticky=tk.W, padx=10, pady=5)
        replacement_teacher_var = tk.StringVar(value=substitution_to_edit.get('replacement_teacher', ''))
        replacement_teacher_combo = ttk.Combobox(dialog, textvariable=replacement_teacher_var, values=teacher_names, state="readonly", width=30)
        replacement_teacher_combo.grid(row=7, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Причина замены:").grid(row=8, column=0, sticky=tk.W, padx=10, pady=5)
        reason_var = tk.StringVar(value="Другое") # Упрощаем, всегда "Другое"
        reason_combo = ttk.Combobox(dialog, textvariable=reason_var, 
                                   values=["Болезнь", "Отпуск", "Командировка", "Другое"], state="readonly", width=30)
        reason_combo.grid(row=8, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Label(dialog, text="Детали:").grid(row=9, column=0, sticky=tk.W, padx=10, pady=5)
        details_var = tk.StringVar(value=substitution_to_edit.get('reason', ''))
        details_entry = ttk.Entry(dialog, textvariable=details_var, width=30)
        details_entry.grid(row=9, column=1, sticky=tk.W, padx=10, pady=5)
        is_temporary_var = tk.BooleanVar(value=substitution_to_edit.get('is_temporary', True))
        ttk.Checkbutton(dialog, text="Временная замена", variable=is_temporary_var).grid(row=10, column=0, columnspan=2, sticky=tk.W, padx=10, pady=5)
        def update_substitution():
            # Проверка заполнения
            required_fields = [
                date_var.get(), week_var.get(), day_var.get(), time_var.get(),
                group_var.get(), subject_var.get(), original_teacher_var.get(), replacement_teacher_var.get()
            ]
            if not all(required_fields):
                messagebox.showwarning("Предупреждение", "Заполните все обязательные поля")
                return
            if original_teacher_var.get() == replacement_teacher_var.get():
                messagebox.showwarning("Предупреждение", "Оригинальный и заменяющий преподаватели не могут быть одинаковыми")
                return
            # Обновляем данные в записи
            substitution_to_edit['date'] = date_var.get()
            substitution_to_edit['week'] = int(week_var.get())
            substitution_to_edit['day'] = day_var.get()
            substitution_to_edit['time'] = time_var.get()
            substitution_to_edit['group'] = group_var.get()
            substitution_to_edit['subject'] = subject_var.get()
            substitution_to_edit['original_teacher'] = original_teacher_var.get()
            substitution_to_edit['replacement_teacher'] = replacement_teacher_var.get()
            substitution_to_edit['reason'] = details_var.get()
            substitution_to_edit['is_temporary'] = is_temporary_var.get()
            # Если замена привязана к расписанию (schedule_index >= 0), обновляем и его
            if substitution_to_edit.get('schedule_index', -1) >= 0:
                idx = substitution_to_edit['schedule_index']
                # Находим нового преподавателя по имени
                new_teacher = next((t for t in self.teachers if t['name'] == replacement_teacher_var.get()), None)
                if new_teacher:
                    self.schedule.loc[idx, 'teacher_id'] = new_teacher['id']
                    self.schedule.loc[idx, 'teacher_name'] = new_teacher['name']
                    self.filter_schedule()
            self.load_substitutions_data()
            self.create_backup()
            messagebox.showinfo("Успех", "Замена успешно обновлена!")
            dialog.destroy()
        ttk.Button(dialog, text="Сохранить", command=update_substitution).grid(row=11, column=0, columnspan=2, pady=20)

    def delete_substitution(self):
        """Удалить выбранную замену из журнала"""
        selected = self.substitutions_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите замену для удаления")
            return
        item = self.substitutions_tree.item(selected[0])
        values = item['values']
        # Находим запись для удаления
        substitution_to_delete = None
        for i, sub in enumerate(self.substitutions):
            if (sub.get('date') == values[0] and
                sub.get('group') == values[1] and
                sub.get('time') == values[2]):
                substitution_to_delete = sub
                delete_index = i
                break
        if not substitution_to_delete:
            messagebox.showerror("Ошибка", "Замена не найдена")
            return
        if messagebox.askyesno("Подтверждение", "Вы уверены, что хотите удалить выбранную замену?"):
            # Если это временная замена и она привязана к расписанию, возвращаем оригинального преподавателя
            if (substitution_to_delete.get('is_temporary', False) and
                substitution_to_delete.get('schedule_index', -1) >= 0):
                idx = substitution_to_delete['schedule_index']
                original_teacher_id = substitution_to_delete.get('original_teacher_id')
                original_teacher_name = substitution_to_delete.get('original_teacher')
                # Находим преподавателя по ID
                original_teacher = next((t for t in self.teachers if t['id'] == original_teacher_id), None)
                if original_teacher:
                    self.schedule.loc[idx, 'teacher_id'] = original_teacher_id
                    self.schedule.loc[idx, 'teacher_name'] = original_teacher_name
                    self.filter_schedule()
            # Удаляем из списка
            del self.substitutions[delete_index]
            self.load_substitutions_data()
            self.create_backup()
            messagebox.showinfo("Успех", "Замена успешно удалена!")

    def export_substitutions(self):
        """Экспорт журнала замен в Excel"""
        if not self.substitutions:
            messagebox.showinfo("Информация", "Журнал замен пуст")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Сохранить журнал замен как"
        )
        if filename:
            try:
                # Создаем DataFrame из списка замен
                df = pd.DataFrame(self.substitutions)
                # Выбираем и переименовываем колонки для экспорта
                export_df = df[['date', 'week', 'day', 'time', 'group', 'subject', 
                               'original_teacher', 'replacement_teacher', 'reason', 'is_temporary']]
                export_df.columns = ['Дата', 'Неделя', 'День', 'Время', 'Группа', 'Предмет',
                                    'Оригинальный преподаватель', 'Заменяющий преподаватель', 'Причина', 'Временная']
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    export_df.to_excel(writer, sheet_name='Журнал замен', index=False)
                messagebox.showinfo("Экспорт", f"Журнал замен успешно экспортирован в {filename}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка экспорта: {str(e)}")

    def open_backup_manager(self):
        """Открыть менеджер бэкапов"""
        # Создание диалогового окна для менеджера бэкапов
        dialog = tk.Toplevel(self.root)
        dialog.title("Менеджер бэкапов")
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        ttk.Label(dialog, text="🛡️ Менеджер бэкапов", font=('Arial', 14, 'bold')).pack(pady=10)
        # Создание фрейма для списка бэкапов
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        # Таблица бэкапов
        columns = ('Имя файла', 'Дата создания', 'Размер')
        self.backup_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.backup_tree.heading(col, text=col)
            self.backup_tree.column(col, width=180)
        self.backup_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # Прокрутка
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.backup_tree.yview)
        self.backup_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # --- ДОБАВЛЕНО: Контекстное меню ---
        context_menu = tk.Menu(self.backup_tree, tearoff=0)
        context_menu.add_command(label="📂 Восстановить", command=self.restore_backup)
        context_menu.add_command(label="🗑️ Удалить", command=self.delete_backup)
        context_menu.add_separator()
        context_menu.add_command(label="🔄 Обновить список", command=self.load_backup_list)
        def show_context_menu(event):
            item = self.backup_tree.identify_row(event.y)
            if item:
                self.backup_tree.selection_set(item)
                context_menu.post(event.x_root, event.y_root)
        self.backup_tree.bind("<Button-3>", show_context_menu)  # ПКМ для вызова меню
        # --- КОНЕЦ ДОБАВЛЕНИЯ ---
        # Кнопки управления (используем стиль Accent.TButton)
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        buttons = [
            ("🔄 Обновить", self.load_backup_list),
            ("💾 Создать бэкап", self.create_backup),
            ("📂 Восстановить", self.restore_backup),
            ("🗑️ Удалить", self.delete_backup)
        ]
        for text, command in buttons:
            btn = ttk.Button(button_frame, text=text, command=command, style='Accent.TButton')
            btn.pack(side=tk.LEFT, padx=(0, 5))
        # Загрузка списка бэкапов
        self.load_backup_list()

    def load_backup_list(self):
        """Загрузить список бэкапов"""
        # Очистка таблицы
        for item in self.backup_tree.get_children():
            self.backup_tree.delete(item)
        # Получение списка файлов из директории бэкапов
        try:
            backup_files = [f for f in os.listdir(self.backup_dir) if f.endswith('.zip')]
            backup_files.sort(reverse=True) # Сортировка по дате (новые первые)
            for filename in backup_files:
                filepath = os.path.join(self.backup_dir, filename)
                stat = os.stat(filepath)
                creation_time = datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
                size = stat.st_size
                # Форматирование размера
                if size < 1024:
                    size_str = f"{size} B"
                elif size < 1024**2:
                    size_str = f"{size/1024:.1f} KB"
                else:
                    size_str = f"{size/(1024**2):.1f} MB"
                self.backup_tree.insert('', tk.END, values=(filename, creation_time, size_str))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка загрузки списка бэкапов: {e}")

    def restore_backup(self):
        """Восстановить из бэкапа"""
        selected = self.backup_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите бэкап для восстановления")
            return
        item = self.backup_tree.item(selected[0])
        filename = item['values'][0]
        filepath = os.path.join(self.backup_dir, filename)
        if messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите восстановить данные из {filename}?\nТекущие данные будут потеряны."):
            try:
                # Распаковка архива
                with zipfile.ZipFile(filepath, 'r') as zip_ref:
                    zip_ref.extractall('.') # Распаковка в текущую директорию
                messagebox.showinfo("Успех", f"Данные успешно восстановлены из {filename}")
                # Перезагрузка данных
                self.load_data()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка восстановления: {e}")

    def delete_backup(self):
        """Удалить бэкап"""
        selected = self.backup_tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите бэкап для удаления")
            return
        item = self.backup_tree.item(selected[0])
        filename = item['values'][0]
        filepath = os.path.join(self.backup_dir, filename)
        if messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите удалить {filename}?"):
            try:
                os.remove(filepath)
                self.load_backup_list() # Обновление списка
                messagebox.showinfo("Успех", f"Бэкап {filename} успешно удален")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка удаления: {e}")

    def create_backup(self):
        """Создать бэкап данных"""
        try:
            # Создание имени файла бэкапа
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"backup_{timestamp}.zip"
            backup_filepath = os.path.join(self.backup_dir, backup_filename)
            # Создание временного файла данных
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
            # Создание архива
            with zipfile.ZipFile(backup_filepath, 'w') as zipf:
                zipf.write(temp_data_file)
            # Удаление временного файла
            os.remove(temp_data_file)
            # Удаление старых бэкапов, если их больше максимального количества
            self.cleanup_old_backups()
            # Обновление времени последнего бэкапа
            self.last_backup_time = datetime.now()
            self.update_backup_indicator()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка создания бэкапа: {str(e)}")

    def cleanup_old_backups(self):
        """Удалить старые бэкапы, если их больше максимального количества"""
        try:
            backup_files = [f for f in os.listdir(self.backup_dir) if f.endswith('.zip')]
            backup_files.sort(key=lambda x: os.path.getctime(os.path.join(self.backup_dir, x)))
            # Удаление самых старых файлов, если их больше максимума
            while len(backup_files) > self.settings.get('max_backups', 10):
                oldest_file = backup_files.pop(0)
                os.remove(os.path.join(self.backup_dir, oldest_file))
        except Exception as e:
            pass  # Игнорируем ошибки очистки

    def start_auto_backup(self):
        """Запустить таймер авто-бэкапа"""
        if self.settings.get('auto_backup', True):
            interval = self.settings.get('backup_interval', 30) * 60 * 1000  # Преобразование минут в миллисекунды
            self.backup_timer = self.root.after(interval, self.auto_backup)

    def auto_backup(self):
        """Автоматический бэкап"""
        self.create_backup()
        # Перезапуск таймера
        self.start_auto_backup()

    def restart_auto_backup(self):
        """Перезапустить таймер авто-бэкапа"""
        if self.backup_timer:
            self.root.after_cancel(self.backup_timer)
        self.start_auto_backup()

    def update_backup_indicator(self):
        """Обновить индикатор бэкапа"""
        if self.settings.get('auto_backup', True):
            self.backup_status_label.config(text="Авто-бэкап: ВКЛ", style='BackupActive.TLabel')
            if self.last_backup_time:
                next_backup = self.last_backup_time + timedelta(minutes=self.settings.get('backup_interval', 30))
                self.next_backup_time = next_backup
                self.backup_info_label.config(text=f"Следующий: {next_backup.strftime('%H:%M')}")
            else:
                self.backup_info_label.config(text="Следующий: --:--")
        else:
            self.backup_status_label.config(text="Авто-бэкап: ВЫКЛ", style='BackupInactive.TLabel')
            self.backup_info_label.config(text="")

    def check_and_update_experience(self):
        """Проверить и обновить стаж преподавателей при необходимости"""
        current_year = datetime.now().year
        last_update_year = self.settings.get('last_academic_year_update', current_year)
        if current_year > last_update_year:
            # Обновляем стаж всех преподавателей
            for teacher in self.teachers:
                teacher['experience'] = teacher.get('experience', 0) + (current_year - last_update_year)
            # Обновляем год последнего обновления
            self.settings['last_academic_year_update'] = current_year
            # Обновляем интерфейс
            self.load_teachers_data()
            # Создаем бэкап после обновления
            self.create_backup()

    def show_about(self):
        """Показать обновленную информацию о программе"""
        # Создание диалогового окна
        dialog = tk.Toplevel(self.root)
        dialog.title("О программе")
        dialog.geometry("605x650")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False) # Запрещаем изменение размера
        
        # Центрирование окна
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        # Основной фрейм с прокруткой
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Создаем Canvas для прокрутки
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Логотип/Иконка
        try:
            logo_label = ttk.Label(scrollable_frame, text="🎓", font=('Arial', 48, 'bold'))
            logo_label.pack(pady=(0, 10))
        except:
            pass

        # Название программы
        title_label = ttk.Label(scrollable_frame, text="Система автоматического составления расписания", 
                               font=('Arial', 16, 'bold'))
        title_label.pack(pady=(0, 5))

        # Версия
        version_label = ttk.Label(scrollable_frame, text="Версия 2.0", font=('Arial', 11, 'bold'))
        version_label.pack(pady=(0, 15))

        # Описание
        description_text = (
            "Это профессиональное приложение предназначено для автоматизации и управления\n"
            "процессом составления учебного расписания в образовательных учреждениях.\n\n"
            "Программа объединяет в себе мощный функционал автоматической генерации и гибкие\n"
            "инструменты ручного управления, что делает ее незаменимым помощником для\n"
            "администраторов и завучей."
        )
        desc_label = ttk.Label(scrollable_frame, text=description_text, justify=tk.LEFT)
        desc_label.pack(pady=(0, 20), anchor='w')

        # Основные возможности
        features_title = ttk.Label(scrollable_frame, text="🔑 Основные возможности:", font=('Arial', 12, 'bold'))
        features_title.pack(pady=(0, 10), anchor='w')

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
        features_label = ttk.Label(scrollable_frame, text=features_text, justify=tk.LEFT)
        features_label.pack(pady=(0, 20), anchor='w')

        # Информация о разработчике
        dev_frame = ttk.Frame(scrollable_frame)
        dev_frame.pack(fill=tk.X, pady=(0, 10), anchor='w')
        ttk.Label(dev_frame, text="👨‍💻 Разработчик:", font=('Arial', 10, 'bold')).pack(anchor='w')
        ttk.Label(dev_frame, text="Команда разработчиков Mikhail Lukomskiy").pack(anchor='w')
        ttk.Label(dev_frame, text="📧 Контактная почта: support@lukomsky.ru").pack(anchor='w')
        ttk.Label(dev_frame, text="🌐 Официальный сайт: www.lukomsky.ru").pack(anchor='w')
        ttk.Label(dev_frame, text="📅 Год выпуска: 2025").pack(anchor='w')

        # Информация об учреждении
        school_frame = ttk.Frame(scrollable_frame)
        school_frame.pack(fill=tk.X, pady=(20, 10), anchor='w')
        ttk.Label(school_frame, text="🏫 Текущее учреждение:", font=('Arial', 10, 'bold')).pack(anchor='w')
        ttk.Label(school_frame, text=f"{self.settings.get('school_name', 'Не указано')}").pack(anchor='w')
        ttk.Label(school_frame, text=f"Директор: {self.settings.get('director', 'Не указан')}").pack(anchor='w')
        ttk.Label(school_frame, text=f"Учебный год: {self.settings.get('academic_year', 'Не указан')}").pack(anchor='w')

        # Кнопка закрытия
        close_button = ttk.Button(scrollable_frame, text="Закрыть", command=dialog.destroy)
        close_button.pack(pady=(20, 0))

# Запуск приложения
if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()
