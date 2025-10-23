#!/usr/bin/env python

# --- Imports ---
import sys
import os
import json
import csv
import inspect
import webbrowser
from appdirs import user_data_dir
from functools import partial

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTreeWidget, QTreeWidgetItem, QHeaderView, QGroupBox, QGridLayout,
    QLabel, QLineEdit, QComboBox, QDialog, QDialogButtonBox, QMessageBox,
    QFileDialog, QInputDialog, QCheckBox, QStackedWidget, QCompleter, QFormLayout
)
from PyQt6.QtGui import QIcon, QDesktopServices, QAction
from PyQt6.QtCore import Qt, QUrl, QStringListModel

# --- Helper function for finding assets (Unchanged) ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# --- Function to get a writable application data directory (Unchanged) ---
def get_app_data_path(app_name, app_author):
    """ Get a writable, cross-platform path for application data. """
    path = user_data_dir(app_name, app_author) # appdirs handles all OS logic
    if not os.path.exists(path):
        os.makedirs(path)
    return path

# --- Data File (Unchanged) ---
APP_NAME = "CanvasSISPrepTool"
APP_AUTHOR = "Harrison Smith"
DATA_DIR = get_app_data_path(APP_NAME, APP_AUTHOR)
DATA_FILE = os.path.join(DATA_DIR, 'course_data.json')


# --- Core Data Models (Unchanged) ---
class ProgramArea:
    def __init__(self, name):
        self.name = name

    def to_dict(self):
        return {'name': self.name}

    @staticmethod
    def from_dict(data):
        return ProgramArea(data['name'])


class Person:
    def __init__(self, name, user_id, program_area_name=''):
        self.name = name
        self.user_id = user_id
        self.program_area_name = program_area_name

    def to_dict(self):
        return {'name': self.name, 'user_id': self.user_id, 'program_area_name': self.program_area_name}

    @staticmethod
    def from_dict(data):
        return Person(data['name'], data['user_id'], data.get('program_area_name', ''))


class Course:
    def __init__(self, short_name, long_name, course_id_portion, program_area_name=''):
        self.short_name = short_name
        self.long_name = long_name
        self.course_id_portion = course_id_portion.upper()
        self.program_area_name = program_area_name

    def to_dict(self):
        return {'short_name': self.short_name, 'long_name': self.long_name,
                'course_id_portion': self.course_id_portion, 'program_area_name': self.program_area_name}

    @staticmethod
    def from_dict(data):
        return Course(data['short_name'], data['long_name'], data['course_id_portion'],
                      data.get('program_area_name', ''))


class Term:
    def __init__(self, term_id, name, short_code):
        self.term_id = term_id
        self.name = name
        self.short_code = short_code

    def to_dict(self):
        return {'term_id': self.term_id, 'name': self.name, 'short_code': self.short_code}

    @staticmethod
    def from_dict(data):
        return Term(data.get('term_id', ''), data.get('name', ''), data.get('short_code', ''))


class Account:
    def __init__(self, account_id):
        self.account_id = account_id

    def to_dict(self):
        return {'account_id': self.account_id}

    @staticmethod
    def from_dict(data):
        return Account(data['account_id'])


class Enrollment:
    def __init__(self, user_id, role, status='active'):
        self.user_id = user_id
        self.role = role
        self.status = status

    def to_dict(self):
        return {'user_id': self.user_id, 'role': self.role, 'status': self.status}

    @staticmethod
    def from_dict(data):
        return Enrollment(data['user_id'], data['role'], data.get('status', 'active'))


class Section:
    def __init__(self, course_id_portion, term_name, account_id, section_number, status='active', start_date='',
                 end_date=''):
        self.course_id_portion = course_id_portion
        self.term_name = term_name
        self.account_id = account_id
        self.section_number = section_number
        self.status = status
        self.start_date = start_date
        self.end_date = end_date
        self.enrollments = []

    def add_enrollment(self, enrollment):
        self.enrollments.append(enrollment)

    def to_dict(self):
        return {'course_id_portion': self.course_id_portion, 'term_name': self.term_name, 'account_id': self.account_id,
                'section_number': self.section_number, 'status': self.status, 'start_date': self.start_date,
                'end_date': self.end_date, 'enrollments': [e.to_dict() for e in self.enrollments]}

    @staticmethod
    def from_dict(data):
        section = Section(data['course_id_portion'], data.get('term_name'), data['account_id'], data['section_number'],
                          data.get('status', 'active'), data.get('start_date', ''), data.get('end_date', ''))
        section.enrollments = [Enrollment.from_dict(e) for e in data.get('enrollments', [])]
        return section


# --- DataManager (Unchanged) ---
class DataManager:
    def __init__(self):
        self.people = {}
        self.courses = {}
        self.terms = {}
        self.accounts = {}
        self.program_areas = {}
        self.enrollment_roles = {}
        self.sections = []
        self.load_data()

    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'r') as f:
                    data = json.load(f)
                    self.people = {uid: Person.from_dict(p) for uid, p in data.get('people', {}).items()}
                    self.courses = {cid: Course.from_dict(c) for cid, c in data.get('courses', {}).items()}
                    # Migration from 'departments' to 'program_areas'
                    program_areas_data = data.get('program_areas', data.get('departments', {}))
                    self.program_areas = {dname: ProgramArea.from_dict(d) for dname, d in program_areas_data.items()}

                    raw_terms = data.get('terms', {})
                    self.terms = {}
                    term_id_to_name_map = {}
                    for _, term_data in raw_terms.items():
                        term_obj = Term.from_dict(term_data)
                        if not term_obj.name:
                            term_obj.name = term_obj.term_id
                        self.terms[term_obj.name] = term_obj
                        term_id_to_name_map[term_obj.term_id] = term_obj.name
                    self.accounts = {aid: Account.from_dict(a) for aid, a in data.get('accounts', {}).items()}
                    self.sections = [Section.from_dict(s) for s in data.get('sections', [])]
                    self.enrollment_roles = data.get('enrollment_roles', self._get_default_roles())

                    for section in self.sections:
                        if hasattr(section, 'term_id') and not hasattr(section, 'term_name'):
                            section.term_name = next(
                                (t.name for t in self.terms.values() if t.term_id == section.term_id), None)
                            delattr(section, 'term_id')
                        elif section.term_name and section.term_name not in self.terms:
                            if section.term_name in term_id_to_name_map:
                                section.term_name = term_id_to_name_map[section.term_name]
            except (json.JSONDecodeError, IOError, StopIteration) as e:
                print(f"Error loading data: {e}. Starting fresh.")
                self.initialize_empty()
        else:
            self.initialize_empty()

    def _get_default_roles(self):
        return {
            "Student": "student",
            "Teaching Assistant": "ta",
            "Instructor": "teacher",
            "Program Manager": "Program Manager"
        }

    def initialize_empty(self):
        self.people = {}
        self.courses = {}
        self.terms = {}
        self.accounts = {}
        self.program_areas = {}
        self.enrollment_roles = self._get_default_roles()
        self.sections = []

    def save_data(self):
        data = {'people': {uid: p.to_dict() for uid, p in self.people.items()},
                'courses': {cid: c.to_dict() for cid, c in self.courses.items()},
                'terms': {tname: t.to_dict() for tname, t in self.terms.items()},
                'accounts': {aid: a.to_dict() for aid, a in self.accounts.items()},
                'program_areas': {dname: d.to_dict() for dname, d in self.program_areas.items()},
                'enrollment_roles': self.enrollment_roles,
                'sections': [s.to_dict() for s in self.sections]}
        try:
            with open(DATA_FILE, 'w') as f:
                json.dump(data, f, indent=4)
            return True
        except IOError as e:
            print(f"Error saving data: {e}")
            return False

    def clear_all(self):
        self.initialize_empty()
        if os.path.exists(DATA_FILE):
            try:
                os.remove(DATA_FILE)
                return True
            except OSError as e:
                print(f"Error removing data file: {e}")
                return False
        return True

    def import_from_csv_file(self, file_path, data_type):
        importer_map = {
            "people": (self.import_people_from_csv, file_path),
            "courses": (self.import_courses_from_csv, file_path),
            "terms": (self.import_terms_from_csv, file_path),
            "accounts": (self.import_accounts_from_csv, file_path),
            "program_areas": (self.import_program_areas_from_csv, file_path)
        }
        if data_type in importer_map:
            func, path = importer_map[data_type]
            return func(path)
        return {'error': f'Unknown data type: {data_type}'}

    def _import_csv_data(self, file_path, required_headers, data_dict, key_field, constructor):
        added_count, skipped_count = 0, 0
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                csv_headers = reader.fieldnames
                if not all(h in csv_headers for h in required_headers):
                    return {'error': f"CSV is missing one of the required headers: {', '.join(required_headers)}"}

                sig = inspect.signature(constructor)
                constructor_params = list(sig.parameters.keys())

                for row in reader:
                    key = row.get(key_field)
                    if not key or key in data_dict:
                        skipped_count += 1
                        continue
                    if any(not row.get(h) for h in required_headers):
                        skipped_count += 1
                        continue

                    args = {k: v for k, v in row.items() if k in constructor_params}
                    data_dict[key] = constructor(**args)
                    added_count += 1
        except Exception as e:
            return {'error': f"An error occurred: {e}"}
        return {'added': added_count, 'skipped': skipped_count}

    def import_people_from_csv(self, file_path):
        return self._import_csv_data(file_path, ['user_id', 'name'], self.people, 'user_id', Person)

    def import_courses_from_csv(self, file_path):
        return self._import_csv_data(file_path, ['course_id_portion', 'short_name', 'long_name'], self.courses,
                                     'course_id_portion', Course)

    def import_program_areas_from_csv(self, file_path):
        return self._import_csv_data(file_path, ['name'], self.program_areas, 'name', ProgramArea)

    def import_terms_from_csv(self, file_path):
        return self._import_csv_data(file_path, ['name', 'term_id', 'short_code'], self.terms, 'name', Term)

    def import_accounts_from_csv(self, file_path):
        return self._import_csv_data(file_path, ['account_id'], self.accounts, 'account_id', Account)

    def import_roles_from_csv(self, file_path):
        added_count, updated_count = 0, 0
        required_headers = ['display_name', 'canvas_role']
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                if not all(h in reader.fieldnames for h in required_headers):
                    return {'error': f"CSV is missing required headers: {', '.join(required_headers)}"}
                for row in reader:
                    display_name = row.get('display_name')
                    canvas_role = row.get('canvas_role')
                    if display_name and canvas_role:
                        if display_name in self.enrollment_roles:
                            updated_count += 1
                        else:
                            added_count += 1
                        self.enrollment_roles[display_name] = canvas_role
            self.save_data()
            return {'added': added_count, 'updated': updated_count}
        except Exception as e:
            return {'error': str(e)}

    def export_data_to_csvs(self, data_types_to_export, directory):
        export_map = {
            'people': (self.people.values(), ['user_id', 'name', 'program_area_name']),
            'courses': (self.courses.values(), ['course_id_portion', 'short_name', 'long_name', 'program_area_name']),
            'terms': (self.terms.values(), ['name', 'term_id', 'short_code']),
            'accounts': (self.accounts.values(), ['account_id']),
            'program_areas': (self.program_areas.values(), ['name'])
        }
        try:
            for data_type in data_types_to_export:
                if data_type in export_map:
                    data_objects, headers = export_map[data_type]
                    if not data_objects:
                        continue
                    file_path = os.path.join(directory, f"{data_type}.csv")
                    with open(file_path, 'w', newline='', encoding='utf-8') as f:
                        writer = csv.DictWriter(f, fieldnames=headers)
                        writer.writeheader()
                        for obj in data_objects:
                            writer.writerow(obj.to_dict())
            return f"Successfully exported selected files to:\n{directory}"
        except IOError as e:
            return f"Error writing files: {e}"

    def generate_csv_files(self, directory, prefix):
        if not self.sections:
            return "No sections created. Cannot generate CSV files."
        courses_data, sections_data, enrollment_data = [], [], []
        for section in self.sections:
            course_obj = self.courses.get(section.course_id_portion)
            term_obj = self.terms.get(section.term_name)
            if not course_obj or not term_obj:
                continue
            course_id = f"{section.course_id_portion}-{term_obj.short_code}-{section.section_number}"
            section_id = f"{term_obj.short_code}-{section.course_id_portion}-{section.section_number}"
            long_name_with_section = f"{course_obj.long_name}-{section.section_number}"
            effective_start_date = section.start_date
            effective_end_date = section.end_date
            courses_data.append(
                {'course_id': course_id, 'short_name': course_obj.short_name, 'long_name': long_name_with_section,
                 'account_id': section.account_id, 'term_id': term_obj.term_id, 'status': section.status,
                 'start_date': effective_start_date, 'end_date': effective_end_date})
            sections_data.append({'section_id': section_id, 'course_id': course_id, 'name': long_name_with_section,
                                  'status': section.status, 'start_date': section.start_date,
                                  'end_date': section.end_date})
            for enrollment in section.enrollments:
                export_role = self.enrollment_roles.get(enrollment.role, enrollment.role)
                enrollment_data.append(
                    {'section_id': section_id, 'user_id': enrollment.user_id, 'role': export_role,
                     'status': enrollment.status, 'course_id': course_id})
        prefix_str = f"{prefix} " if prefix else ""
        try:
            courses_path = os.path.join(directory, f"{prefix_str}courses.csv")
            with open(courses_path, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['course_id', 'short_name', 'long_name', 'account_id', 'term_id',
                                                       'status', 'start_date', 'end_date'])
                writer.writeheader()
                writer.writerows(courses_data)
            sections_path = os.path.join(directory, f"{prefix_str}sections.csv")
            with open(sections_path, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['section_id', 'course_id', 'name', 'status', 'start_date',
                                                       'end_date'])
                writer.writeheader()
                writer.writerows(sections_data)
            enrollments_path = os.path.join(directory, f"{prefix_str}enrollments.csv")
            with open(enrollments_path, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['course_id', 'section_id', 'user_id', 'role', 'status'])
                writer.writeheader()
                writer.writerows(enrollment_data)
            return f"Successfully generated files in:\n{directory}"
        except IOError as e:
            return f"Error writing files: {e}"


# --- Custom Autocomplete Combobox (PyQt6 Version) ---
class AutocompleteCombobox(QComboBox):
    """
    A QComboBox that supports autocompletion and filtering as the user types.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        self.completer = QCompleter(self)
        self.completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self.completer.setFilterMode(Qt.MatchFlag.MatchContains)
        self.completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.setCompleter(self.completer)

        self._model = QStringListModel()
        self.completer.setModel(self._model)

    def set_completion_list(self, items):
        """Public method to update the master list of options."""
        self._model.setStringList(items)
        self.setModel(self._model) # Use the same model for the combobox itself
        self.setCurrentIndex(-1) # Clear selection

    def is_valid(self):
        """Check if the current value is an exact match in the master completion list."""
        return self.currentText() in self._model.stringList()

# --- GUI Themes ---
LIGHT_THEME = """
    QWidget {
        background-color: #F2F2F2;
        color: #222222;
    }
    QMainWindow, QDialog {
        background-color: #F2F2F2;
    }
    QTabWidget::pane {
        border: 1px solid #E0E0E0;
        background-color: #FFFFFF;
    }
    QTabBar::tab {
        background-color: #F2F2F2;
        color: #222222;
        padding: 10px 15px;
        border: 1px solid #E0E0E0;
        border-bottom: none;
        margin-right: -1px;
    }
    QTabBar::tab:selected {
        background-color: #FFFFFF;
        color: #D94F4F;
        border-bottom: 1px solid #FFFFFF;
    }
    QTabBar::tab:!selected:hover {
        background-color: #E8E8E8;
    }
    QTreeWidget {
        background-color: #FFFFFF;
        color: #222222;
        border: 1px solid #E0E0E0;
        alternate-background-color: #F8F8F8;
    }
    QTreeWidget::item:selected {
        background-color: #4D9FE0;
        color: #FFFFFF;
    }
    QHeaderView::section {
        background-color: #E0E0E0;
        color: #222222;
        padding: 5px;
        border: 1px solid #D0D0D0;
        font-weight: bold;
    }
    QPushButton {
        background-color: #E1E1E1;
        color: #222222;
        border: 1px solid #C0C0C0;
        padding: 5px 15px;
        border-radius: 3px;
    }
    QPushButton:hover {
        background-color: #4D9FE0;
        color: #FFFFFF;
    }
    QPushButton:pressed {
        background-color: #3C8ACF;
    }
    QPushButton[accent="true"] {
        background-color: #D94F4F;
        color: #FFFFFF;
        font-weight: bold;
    }
    QPushButton[accent="true"]:hover {
        background-color: #B84444;
    }
    QPushButton[accent="true"]:pressed {
        background-color: #A33B3B;
    }
    QGroupBox {
        font-weight: bold;
        background-color: #F2F2F2;
        border: 1px solid #E0E0E0;
        border-radius: 3px;
        margin-top: 10px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 0 5px 0 5px;
        background-color: #F2F2F2;
        color: #222222;
    }
    QLineEdit, QComboBox {
        background-color: #FFFFFF;
        border: 1px solid #C0C0C0;
        padding: 3px;
        border-radius: 3px;
    }
    QComboBox[readonly="true"] {
        background-color: #E1E1E1;
    }
    QLabel[required="true"] {
        color: red;
        font-weight: bold;
    }
"""

DARK_THEME = """
    QWidget {
        background-color: #212121;
        color: #F5F5F5;
    }
    QMainWindow, QDialog {
        background-color: #212121;
    }
    QTabWidget::pane {
        border: 1px solid #383838;
        background-color: #2C2C2C;
    }
    QTabBar::tab {
        background-color: #212121;
        color: #F5F5F5;
        padding: 10px 15px;
        border: 1px solid #383838;
        border-bottom: none;
        margin-right: -1px;
    }
    QTabBar::tab:selected {
        background-color: #2C2C2C;
        color: #D94F4F;
        border-bottom: 1px solid #2C2C2C;
    }
    QTabBar::tab:!selected:hover {
        background-color: #333333;
    }
    QTreeWidget {
        background-color: #2C2C2C;
        color: #F5F5F5;
        border: 1px solid #383838;
        alternate-background-color: #333333;
    }
    QTreeWidget::item:selected {
        background-color: #4D9FE0;
        color: #FFFFFF;
    }
    QHeaderView::section {
        background-color: #383838;
        color: #F5F5F5;
        padding: 5px;
        border: 1px solid #4A4A4A;
        font-weight: bold;
    }
    QPushButton {
        background-color: #4F4F4F;
        color: #F5F5F5;
        border: 1px solid #606060;
        padding: 5px 15px;
        border-radius: 3px;
    }
    QPushButton:hover {
        background-color: #4D9FE0;
        color: #FFFFFF;
    }
    QPushButton:pressed {
        background-color: #3C8ACF;
    }
    QPushButton[accent="true"] {
        background-color: #D94F4F;
        color: #FFFFFF;
        font-weight: bold;
    }
    QPushButton[accent="true"]:hover {
        background-color: #B84444;
    }
    QPushButton[accent="true"]:pressed {
        background-color: #A33B3B;
    }
    QGroupBox {
        font-weight: bold;
        background-color: #212121;
        border: 1px solid #383838;
        border-radius: 3px;
        margin-top: 10px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 0 5px 0 5px;
        background-color: #212121;
        color: #F5F5F5;
    }
    QLineEdit, QComboBox {
        background-color: #3C3C3C;
        color: #F5F5F5;
        border: 1px solid #606060;
        padding: 3px;
        border-radius: 3px;
    }
    QComboBox[readonly="true"] {
        background-color: #4F4F4F;
    }
    QLabel[required="true"] {
        color: red;
        font-weight: bold;
    }
"""


# --- GUI Application (PyQt6 Version) ---
class App(QMainWindow):
    def __init__(self, data_manager):
        super().__init__()
        self.data_manager = data_manager
        self.current_theme = "light"

        self.setWindowTitle("Canvas SIS Prep Tool")
        self.setGeometry(100, 100, 1350, 750)
        try:
            self.setWindowIcon(QIcon(resource_path('app_icon.png')))
        except Exception:
            print("app_icon.png not found, skipping icon.")

        self.notebook = QTabWidget()
        self.setCentralWidget(self.notebook)

        self.create_sections_tab()
        self.create_settings_tab()
        self.create_people_tab()
        self.create_courses_tab()
        self.create_program_areas_tab()
        self.create_terms_tab()
        self.create_accounts_tab()
        self.create_about_tab()
        
        self.setup_styles()

    def closeEvent(self, event):
        """Save data on window close."""
        self.data_manager.save_data()
        event.accept()

    def setup_styles(self):
        """Applies the current theme stylesheet to the application."""
        theme = DARK_THEME if self.current_theme == "dark" else LIGHT_THEME
        self.setStyleSheet(theme)

    def toggle_theme(self):
        """Switches the application theme between light and dark."""
        self.current_theme = "dark" if self.current_theme == "light" else "light"
        self.setup_styles()

    def create_management_tab(self, title, columns, item_name, add_slot, edit_slot, delete_slot):
        """Helper to create a standard data management tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # --- Tree Widget ---
        tree = QTreeWidget()
        tree.setColumnCount(len(columns))
        tree.setHeaderLabels([col.replace('_', ' ').title() for col in columns])
        tree.header().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        tree.setAlternatingRowColors(True)
        tree.itemDoubleClicked.connect(edit_slot)
        layout.addWidget(tree)

        # --- Button Bar ---
        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.setContentsMargins(0, 5, 0, 0)
        
        add_btn = QPushButton(f"Add {item_name}")
        add_btn.clicked.connect(add_slot)
        
        edit_btn = QPushButton(f"Edit {item_name}")
        edit_btn.clicked.connect(edit_slot)
        
        delete_btn = QPushButton(f"Delete {item_name}")
        delete_btn.clicked.connect(delete_slot)
        
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        
        layout.addWidget(btn_widget)
        self.notebook.addTab(tab, title)
        return tree

    def refresh_all_views(self):
        self.refresh_people_list()
        self.refresh_courses_list()
        self.refresh_program_areas_list()
        self.refresh_terms_list()
        self.refresh_accounts_list()
        self.refresh_sections_list()

    # --- People Tab ---
    def create_people_tab(self):
        self.people_tree = self.create_management_tab(
            "People", ('user_id', 'name', 'program_area_name'), "Person",
            self.add_person, self.edit_person, self.delete_person
        )
        self.refresh_people_list()

    def refresh_people_list(self):
        self.people_tree.clear()
        for user_id, person in sorted(self.data_manager.people.items()):
            item = QTreeWidgetItem([str(user_id), person.name, person.program_area_name])
            self.people_tree.addTopLevelItem(item)

    def add_person(self): self.add_item("Person", self.people_tree)
    def edit_person(self): self.edit_item("Person", self.people_tree)
    def delete_person(self): self.delete_item("Person", self.people_tree)

    # --- Courses Tab ---
    def create_courses_tab(self):
        self.courses_tree = self.create_management_tab(
            "Courses", ('course_id_portion', 'short_name', 'long_name', 'program_area_name'), "Course",
            self.add_course, self.edit_course, self.delete_course
        )
        self.refresh_courses_list()

    def refresh_courses_list(self):
        self.courses_tree.clear()
        for cid, course in sorted(self.data_manager.courses.items()):
            item = QTreeWidgetItem([cid, course.short_name, course.long_name, course.program_area_name])
            self.courses_tree.addTopLevelItem(item)

    def add_course(self): self.add_item("Course", self.courses_tree)
    def edit_course(self): self.edit_item("Course", self.courses_tree)
    def delete_course(self): self.delete_item("Course", self.courses_tree)

    # --- Program Areas Tab ---
    def create_program_areas_tab(self):
        self.program_areas_tree = self.create_management_tab(
            "Program Areas", ('name',), "Program Area",
            self.add_program_area, self.edit_program_area, self.delete_program_area
        )
        self.refresh_program_areas_list()

    def refresh_program_areas_list(self):
        self.program_areas_tree.clear()
        for name, dept in sorted(self.data_manager.program_areas.items()):
            item = QTreeWidgetItem([name])
            self.program_areas_tree.addTopLevelItem(item)

    def add_program_area(self): self.add_item("Program Area", self.program_areas_tree)
    def edit_program_area(self): self.edit_item("Program Area", self.program_areas_tree)
    def delete_program_area(self): self.delete_item("Program Area", self.program_areas_tree)

    # --- Terms Tab ---
    def create_terms_tab(self):
        self.terms_tree = self.create_management_tab(
            "Terms", ('name', 'term_id', 'short_code'), "Term",
            self.add_term, self.edit_term, self.delete_term
        )
        self.refresh_terms_list()

    def refresh_terms_list(self):
        self.terms_tree.clear()
        for tname, term in sorted(self.data_manager.terms.items()):
            item = QTreeWidgetItem([tname, str(term.term_id), term.short_code])
            self.terms_tree.addTopLevelItem(item)

    def add_term(self): self.add_item("Term", self.terms_tree)
    def edit_term(self): self.edit_item("Term", self.terms_tree)
    def delete_term(self): self.delete_item("Term", self.terms_tree)

    # --- Accounts Tab ---
    def create_accounts_tab(self):
        self.accounts_tree = self.create_management_tab(
            "Accounts", ('account_id',), "Account",
            self.add_account, self.edit_account, self.delete_account
        )
        self.refresh_accounts_list()

    def refresh_accounts_list(self):
        self.accounts_tree.clear()
        for aid, acc in sorted(self.data_manager.accounts.items()):
            item = QTreeWidgetItem([str(aid)])
            self.accounts_tree.addTopLevelItem(item)
    
    def add_account(self): self.add_item("Account", self.accounts_tree)
    def edit_account(self): self.edit_item("Account", self.accounts_tree)
    def delete_account(self): self.delete_item("Account", self.accounts_tree)

    # --- Sections Tab ---
    def create_sections_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        self.sections_tree = QTreeWidget()
        columns = ('course', 'term', 'section', 'status', 'start_date', 'end_date')
        self.sections_tree.setColumnCount(len(columns))
        self.sections_tree.setHeaderLabels([c.replace('_', ' ').title() for c in columns])
        self.sections_tree.header().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.sections_tree.setAlternatingRowColors(True)
        self.sections_tree.itemDoubleClicked.connect(self.edit_section)
        layout.addWidget(self.sections_tree)

        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.setContentsMargins(0, 5, 0, 0)
        
        create_btn = QPushButton("Create Section")
        create_btn.clicked.connect(self.create_section)
        edit_btn = QPushButton("Edit Section")
        edit_btn.clicked.connect(self.edit_section)
        manage_btn = QPushButton("Manage Enrollments")
        manage_btn.clicked.connect(self.manage_enrollments)
        delete_btn = QPushButton("Delete Section")
        delete_btn.clicked.connect(self.delete_section)
        
        btn_layout.addWidget(create_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(manage_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        
        layout.addWidget(btn_widget)
        self.notebook.addTab(tab, "Sections")
        self.refresh_sections_list()

    def refresh_sections_list(self):
        self.sections_tree.clear()
        for i, sec in enumerate(self.data_manager.sections):
            course = self.data_manager.courses.get(sec.course_id_portion)
            course_name = course.short_name if course else 'N/A'
            term_obj = self.data_manager.terms.get(sec.term_name)
            term_display = f"{term_obj.name} ({term_obj.term_id})" if term_obj else 'N/A'
            
            values = [
                f"{course_name} ({sec.course_id_portion})",
                term_display,
                str(sec.section_number),
                sec.status,
                sec.start_date,
                sec.end_date
            ]
            item = QTreeWidgetItem(values)
            # Store the list index in the item
            item.setData(0, Qt.ItemDataRole.UserRole, i)
            self.sections_tree.addTopLevelItem(item)

    def _get_selected_section_index(self):
        """Helper to get the index of the selected section from the tree."""
        selected = self.sections_tree.selectedItems()
        if not selected:
            return None
        return selected[0].data(0, Qt.ItemDataRole.UserRole)

    def create_section(self):
        dialog = SectionDialog(self, "Create Section", self.data_manager)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            new_section = Section(**data)
            self.data_manager.sections.append(new_section)
            self.refresh_sections_list()
            self.data_manager.save_data()
            
            # Select the new section and open enrollment dialog
            new_index = len(self.data_manager.sections) - 1
            for i in range(self.sections_tree.topLevelItemCount()):
                item = self.sections_tree.topLevelItem(i)
                if item.data(0, Qt.ItemDataRole.UserRole) == new_index:
                    self.sections_tree.setCurrentItem(item)
                    break
            self.manage_enrollments()

    def edit_section(self):
        section_index = self._get_selected_section_index()
        if section_index is None:
            QMessageBox.warning(self, "Selection Error", "Please select a section to edit.")
            return
            
        section_obj = self.data_manager.sections[section_index]
        dialog = SectionDialog(self, "Edit Section", self.data_manager, initial_data=section_obj.to_dict())
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            for key, value in data.items():
                setattr(section_obj, key, value)
            self.refresh_sections_list()
            self.data_manager.save_data()

    def manage_enrollments(self):
        section_index = self._get_selected_section_index()
        if section_index is None:
            QMessageBox.warning(self, "Selection Error", "Please select a section to manage.")
            return
            
        section = self.data_manager.sections[section_index]
        term_obj = self.data_manager.terms.get(section.term_name)
        if term_obj:
            full_section_id = f"{term_obj.short_code}-{section.course_id_portion}-{section.section_number}"
            title = f"Enrollments for {full_section_id}"
        else:
            title = f"Enrollments for {section.course_id_portion}"
            
        # This dialog modifies the section object directly
        dialog = EnrollmentDialog(self, title, section, self.data_manager)
        dialog.exec()
        # Data is saved inside the dialog's add/delete methods

    def delete_section(self):
        selected_items = self.sections_tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Selection Error", "Please select a section to delete.")
            return

        reply = QMessageBox.question(self, "Confirm Delete",
                                     "Are you sure you want to delete this section and all its enrollments?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            # Get indices from items and sort in reverse order
            indices_to_delete = sorted([item.data(0, Qt.ItemDataRole.UserRole) for item in selected_items], reverse=True)
            for index in indices_to_delete:
                del self.data_manager.sections[index]
            self.refresh_sections_list()
            self.data_manager.save_data()

    # --- Generic Item Management ---
    def add_item(self, item_name, tree):
        fields_map = {
            "Person": [('user_id', 'User ID'), ('name', 'Name'), ('program_area_name', 'Program Area')],
            "Course": [('course_id_portion', 'Course ID Portion'), ('short_name', 'Short Name'),
                       ('long_name', 'Long Name'), ('program_area_name', 'Program Area')],
            "Term": [('name', 'Name (Unique)'), ('term_id', 'Term ID'), ('short_code', 'Short Code')],
            "Account": [('account_id', 'Account ID')],
            "Program Area": [('name', 'Program Area Name')]
        }
        key_field_map = {"Term": 'name', "Program Area": 'name', "Person": "user_id", "Course": "course_id_portion",
                         "Account": "account_id"}
        key_field = key_field_map.get(item_name)

        combobox_fields = None
        if item_name in ["Person", "Course"]:
            program_areas = [""] + sorted(list(self.data_manager.program_areas.keys()))
            combobox_fields = {'program_area_name': program_areas}

        dialog = ManagementDialog(self, f"Add {item_name}", fields_map[item_name], combobox_fields=combobox_fields)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        data = dialog.get_data()
        for key, value in data.items():
            if not value and key != 'program_area_name':
                QMessageBox.critical(self, "Input Error", f"{key.replace('_', ' ').title()} cannot be empty.")
                return

        key = data[key_field]
        data_map = {"Person": self.data_manager.people, "Course": self.data_manager.courses,
                    "Term": self.data_manager.terms, "Account": self.data_manager.accounts,
                    "Program Area": self.data_manager.program_areas}
        if key in data_map[item_name]:
            QMessageBox.critical(self, "Error", f"A {item_name.lower()} with this ID/Name already exists.")
            return

        constructors = {"Person": Person, "Course": Course, "Term": Term, "Account": Account,
                        "Program Area": ProgramArea}
        data_map[item_name][key] = constructors[item_name](**data)

        refresh_map = {"Person": self.refresh_people_list, "Course": self.refresh_courses_list,
                       "Term": self.refresh_terms_list, "Account": self.refresh_accounts_list,
                       "Program Area": self.refresh_program_areas_list}
        refresh_map[item_name]()
        self.data_manager.save_data()

    def edit_item(self, item_name, tree):
        selected = tree.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Selection Error", f"Please select a {item_name.lower()} to edit.")
            return

        fields_map = {
            "Person": [('user_id', 'User ID'), ('name', 'Name'), ('program_area_name', 'Program Area')],
            "Course": [('course_id_portion', 'Course ID Portion'), ('short_name', 'Short Name'),
                       ('long_name', 'Long Name'), ('program_area_name', 'Program Area')],
            "Term": [('name', 'Name (Unique)'), ('term_id', 'Term ID'), ('short_code', 'Short Code')],
            "Account": [('account_id', 'Account ID')],
            "Program Area": [('name', 'Program Area Name')]
        }
        key_field_map = {"Term": 'name', "Program Area": 'name', "Person": "user_id", "Course": "course_id_portion",
                         "Account": "account_id"}
        key_field = key_field_map.get(item_name)
        key_index = [f[0] for f in fields_map[item_name]].index(key_field)

        # .text() always returns a string, so no bug fix needed
        key = selected[0].text(key_index)

        combobox_fields = None
        if item_name in ["Person", "Course"]:
            program_areas = [""] + sorted(list(self.data_manager.program_areas.keys()))
            combobox_fields = {'program_area_name': program_areas}

        data_map = {"Person": self.data_manager.people, "Course": self.data_manager.courses,
                    "Term": self.data_manager.terms, "Account": self.data_manager.accounts,
                    "Program Area": self.data_manager.program_areas}
        item_obj = data_map[item_name][key]

        dialog = ManagementDialog(self, f"Edit {item_name}", fields_map[item_name], initial_data=item_obj.to_dict(),
                                  readonly_key=key_field, combobox_fields=combobox_fields)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        data = dialog.get_data()
        for k, value in data.items():
            if k != key_field and not value and k != 'program_area_name':
                QMessageBox.critical(self, "Input Error", f"{k.replace('_', ' ').title()} cannot be empty.")
                return

        new_key = data[key_field]
        if new_key != key:  # Key has changed
            if new_key in data_map[item_name]:
                QMessageBox.critical(self, "Error", f"A {item_name.lower()} with this ID/Name already exists.")
                return
            # Update references if needed
            if item_name == "Term":
                for section in self.data_manager.sections:
                    if section.term_name == key: section.term_name = new_key
            if item_name == "Program Area":
                for person in self.data_manager.people.values():
                    if person.program_area_name == key: person.program_area_name = new_key
                for course in self.data_manager.courses.values():
                    if course.program_area_name == key: course.program_area_name = new_key
            del data_map[item_name][key]  # Delete old entry

        constructors = {"Person": Person, "Course": Course, "Term": Term, "Account": Account,
                        "Program Area": ProgramArea}
        data_map[item_name][new_key] = constructors[item_name](**data)

        refresh_map = {"Person": self.refresh_people_list, "Course": self.refresh_courses_list,
                       "Term": self.refresh_terms_list, "Account": self.refresh_accounts_list,
                       "Program Area": self.refresh_program_areas_list}
        refresh_map[item_name]()

        if item_name in ["Term", "Program Area"]:
            self.refresh_sections_list()
            self.refresh_people_list()
            self.refresh_courses_list()

        self.data_manager.save_data()

    def delete_item(self, item_name, tree):
        selected = tree.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Selection Error", f"Please select a {item_name.lower()} to delete.")
            return

        fields_map = {"Person": [('user_id',)], "Course": [('course_id_portion',)], "Term": [('name',)],
                      "Account": [('account_id',)], "Program Area": [('name',)]}
        key_field = fields_map[item_name][0][0]
        
        # Get column index from header label
        key_index = -1
        for i in range(tree.columnCount()):
            if tree.headerItem().text(i).lower() == key_field.replace('_', ' ').lower():
                key_index = i
                break
        
        if key_index == -1: # Fallback just in case
            key_index = 0 

        key = selected[0].text(key_index)

        in_use_msg = f"Cannot delete this {item_name.lower()}. It is in use."
        if item_name == "Person" and any(e.user_id == key for s in self.data_manager.sections for e in s.enrollments):
            QMessageBox.critical(self, "Deletion Error", in_use_msg)
            return
        if item_name == "Program Area":
            if any(p.program_area_name == key for p in self.data_manager.people.values()) or any(
                    c.program_area_name == key for c in self.data_manager.courses.values()):
                QMessageBox.critical(self, "Deletion Error", in_use_msg)
                return
        if item_name in ["Course", "Term", "Account"]:
            check_key = {"Course": "course_id_portion", "Term": "term_name", "Account": "account_id"}[item_name]
            if any(getattr(s, check_key, None) == key for s in self.data_manager.sections):
                QMessageBox.critical(self, "Deletion Error", in_use_msg)
                return

        reply = QMessageBox.question(self, "Confirm Delete",
                                     f"Are you sure you want to delete this {item_name.lower()}?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            data_map = {"Person": self.data_manager.people, "Course": self.data_manager.courses,
                        "Term": self.data_manager.terms, "Account": self.data_manager.accounts,
                        "Program Area": self.data_manager.program_areas}
            if key in data_map[item_name]: del data_map[item_name][key]

            refresh_map = {"Person": self.refresh_people_list, "Course": self.refresh_courses_list,
                           "Term": self.refresh_terms_list, "Account": self.refresh_accounts_list,
                           "Program Area": self.refresh_program_areas_list}
            refresh_map[item_name]()
            self.data_manager.save_data()

    # --- Actions / Settings Tab ---
    def create_settings_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Use QStackedWidget to switch between main view and roles view
        self.action_stack = QStackedWidget()
        layout.addWidget(self.action_stack)

        # --- Main Actions View ---
        self.main_actions_frame = QWidget()
        main_layout = QVBoxLayout(self.main_actions_frame)
        main_layout.setContentsMargins(0,0,0,0)

        canvas_frame = QGroupBox("Canvas SIS Files")
        canvas_layout = QVBoxLayout(canvas_frame)
        gen_csv_btn = QPushButton("Generate Canvas CSV Files...")
        gen_csv_btn.setProperty("accent", True)
        gen_csv_btn.clicked.connect(self.generate_csv)
        canvas_layout.addWidget(gen_csv_btn)
        main_layout.addWidget(canvas_frame)

        data_frame = QGroupBox("Application Data")
        data_layout = QVBoxLayout(data_frame)
        import_btn = QPushButton("Import Data from CSVs...")
        import_btn.clicked.connect(self.open_import_dialog)
        export_btn = QPushButton("Export Data to CSVs...")
        export_btn.clicked.connect(self.open_export_dialog)
        data_layout.addWidget(import_btn)
        data_layout.addWidget(export_btn)
        main_layout.addWidget(data_frame)

        danger_frame = QGroupBox("Danger Zone")
        danger_layout = QVBoxLayout(danger_frame)
        roles_btn = QPushButton("Manage Enrollment Roles")
        roles_btn.clicked.connect(self.show_roles_management_view)
        clear_btn = QPushButton("Clear All Local Data")
        clear_btn.setProperty("accent", True)
        clear_btn.clicked.connect(self.clear_all_data)
        danger_layout.addWidget(roles_btn)
        danger_layout.addWidget(clear_btn)
        main_layout.addWidget(danger_frame)
        
        main_layout.addStretch()
        self.action_stack.addWidget(self.main_actions_frame)

        # --- Roles Management View ---
        self.roles_management_frame = QWidget()
        self.setup_roles_management_view()
        self.action_stack.addWidget(self.roles_management_frame)

        self.notebook.addTab(tab, "Actions")
        
    def setup_roles_management_view(self):
        layout = QVBoxLayout(self.roles_management_frame)
        back_btn = QPushButton("< Back to Actions")
        back_btn.clicked.connect(self.show_main_actions_view)
        layout.addWidget(back_btn, alignment=Qt.AlignmentFlag.AlignLeft)

        tree_frame = QGroupBox("Enrollment Roles")
        tree_layout = QVBoxLayout(tree_frame)
        
        self.roles_tree = QTreeWidget()
        self.roles_tree.setColumnCount(2)
        self.roles_tree.setHeaderLabels(['Display Name', 'Canvas Role Value'])
        self.roles_tree.header().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        tree_layout.addWidget(self.roles_tree)
        layout.addWidget(tree_frame)

        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.setContentsMargins(0, 5, 0, 0)
        
        add_btn = QPushButton("Add Role")
        add_btn.clicked.connect(self.add_role)
        edit_btn = QPushButton("Edit Role")
        edit_btn.clicked.connect(self.edit_role)
        delete_btn = QPushButton("Delete Role")
        delete_btn.clicked.connect(self.delete_role)
        import_btn = QPushButton("Import from CSV")
        import_btn.clicked.connect(self.import_roles)

        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(import_btn)
        layout.addWidget(btn_widget)

    def show_roles_management_view(self):
        self.refresh_roles_list()
        self.action_stack.setCurrentWidget(self.roles_management_frame)

    def show_main_actions_view(self):
        self.action_stack.setCurrentWidget(self.main_actions_frame)

    def refresh_roles_list(self):
        self.roles_tree.clear()
        for display, canvas_role in sorted(self.data_manager.enrollment_roles.items()):
            item = QTreeWidgetItem([display, canvas_role])
            self.roles_tree.addTopLevelItem(item)

    def add_role(self):
        dialog = RoleEditDialog(self, "Add New Role", is_new=True)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            display_name = data['display_name']
            canvas_role = data['canvas_role']
            if not display_name or not canvas_role:
                QMessageBox.critical(self, "Error", "Both fields are required.")
                return
            if display_name in self.data_manager.enrollment_roles:
                QMessageBox.critical(self, "Error", "A role with this display name already exists.")
                return
            self.data_manager.enrollment_roles[display_name] = canvas_role
            self.data_manager.save_data()
            self.refresh_roles_list()

    def edit_role(self):
        selected = self.roles_tree.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Selection Error", "Please select a role to edit.")
            return
            
        display_name = selected[0].text(0)
        canvas_role = self.data_manager.enrollment_roles[display_name]

        dialog = RoleEditDialog(self, "Edit Role", initial_data={'display_name': display_name, 'canvas_role': canvas_role})
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.data_manager.enrollment_roles[display_name] = dialog.get_data()['canvas_role']
            self.data_manager.save_data()
            self.refresh_roles_list()

    def delete_role(self):
        selected = self.roles_tree.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Selection Error", "Please select a role to delete.")
            return
            
        display_name = selected[0].text(0)
        reply = QMessageBox.question(self, "Confirm Delete",
                                     f"Are you sure you want to delete the role '{display_name}'?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            del self.data_manager.enrollment_roles[display_name]
            self.data_manager.save_data()
            self.refresh_roles_list()

    def import_roles(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Roles CSV File", "", "CSV Files (*.csv)")
        if not path:
            return
            
        result = self.data_manager.import_roles_from_csv(path)
        if 'error' in result:
            QMessageBox.critical(self, "Import Error", result['error'])
        else:
            QMessageBox.information(self, "Import Complete", f"Added: {result['added']}\nUpdated: {result['updated']}")
            self.refresh_roles_list()

    def generate_csv(self):
        prefix, ok = QInputDialog.getText(self, "File Prefix", "Enter an optional prefix for the CSV files:")
        if not ok:
            return
            
        directory = QFileDialog.getExistingDirectory(self, "Select Folder to Save CSV Files")
        if not directory:
            return
            
        result = self.data_manager.generate_csv_files(directory, prefix)
        QMessageBox.information(self, "CSV Generation", result)

    def open_import_dialog(self):
        dialog = ImportDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            summary = []
            file_paths = dialog.get_data()
            for data_type, file_path in file_paths.items():
                result = self.data_manager.import_from_csv_file(file_path, data_type)
                if 'error' in result:
                    summary.append(f"Error importing {data_type.title()}: {result['error']}")
                else:
                    summary.append(
                        f"Imported {data_type.title()}: {result['added']} added, {result['skipped']} skipped (duplicates or blank rows).")
            self.refresh_all_views()
            self.data_manager.save_data()
            QMessageBox.information(self, "Import Complete", "\n".join(summary))

    def open_export_dialog(self):
        dialog = ExportDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            result = dialog.get_data()
            data_types, directory = result['types'], result['directory']
            result_msg = self.data_manager.export_data_to_csvs(data_types, directory)
            QMessageBox.information(self, "Export Complete", result_msg)

    def clear_all_data(self):
        reply = QMessageBox.question(self, "Confirm Clear",
                                     "Are you sure you want to delete ALL local data? This action is irreversible.",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            if self.data_manager.clear_all():
                self.refresh_all_views()
                QMessageBox.information(self, "Success", "All local data has been cleared.")
            else:
                QMessageBox.critical(self, "Error", "Could not clear all data. Check file permissions.")

    # --- About Tab ---
    def create_about_tab(self):
        tab = QWidget()
        
        # Center the content
        main_layout = QHBoxLayout(tab)
        main_layout.addStretch()
        
        content_widget = QWidget()
        layout = QFormLayout(content_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)
        
        layout.addRow(QLabel("<b>App Name:</b>"), QLabel("Canvas SIS Prep Tool"))
        layout.addRow(QLabel("<b>Author:</b>"), QLabel("Harrison Smith"))
        layout.addRow(QLabel("<b>AI Assistant:</b>"), QLabel("Gemini"))
        layout.addRow(QLabel("<b>Last Update Date:</b>"), QLabel("October 23, 2025"))
        
        # Links
        links_widget = QWidget()
        links_layout = QHBoxLayout(links_widget)
        links_layout.setContentsMargins(0,0,0,0)
        github_link = QLabel('<a href="https://github.com/hsmith-dev">GitHub</a>')
        github_link.setOpenExternalLinks(True)
        linkedin_link = QLabel('<a href="https://linkedin.com/in/hsmith-dev">LinkedIn</a>')
        linkedin_link.setOpenExternalLinks(True)
        email_link = QLabel('<a href="mailto:harrison@hsmith.dev">Email</a>')
        email_link.setOpenExternalLinks(True)
        links_layout.addWidget(github_link)
        links_layout.addWidget(linkedin_link)
        links_layout.addWidget(email_link)
        links_layout.addStretch()
        layout.addRow(QLabel("<b>Links:</b>"), links_widget)

        description_text = "A desktop application designed to streamline the creation of CSV files for Canvas SIS imports currently focused at streamlining the course shell creations."
        desc_label = QLabel(description_text)
        desc_label.setWordWrap(True)
        layout.addRow(QLabel("<b>Description:</b>"), desc_label)
        
        theme_btn = QPushButton("Toggle Theme")
        theme_btn.clicked.connect(self.toggle_theme)
        layout.addRow(theme_btn)
        
        main_layout.addWidget(content_widget)
        main_layout.addStretch()
        
        self.notebook.addTab(tab, "About")


# --- Custom Dialogs (PyQt6 Version) ---

class BaseDialog(QDialog):
    """Base dialog for consistent button box and validation."""
    def __init__(self, parent, title):
        super().__init__(parent)
        self.setWindowTitle(title)
        
        self.main_layout = QVBoxLayout(self)
        self.form_widget = QWidget()
        self.form_layout = QFormLayout(self.form_widget)
        self.main_layout.addWidget(self.form_widget)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.main_layout.addWidget(self.button_box)
        
        self.widgets = {}

    def add_field(self, key, label_text, widget, is_required=False):
        """Adds a field to the form layout."""
        if is_required:
            label = QLabel(f'{label_text}: <span style="color:red;">*</span>')
        else:
            label = QLabel(f"{label_text}:")
            
        self.form_layout.addRow(label, widget)
        self.widgets[key] = widget

    def accept(self):
        """Overridden to provide validation."""
        if self.validate_form():
            super().accept()
        else:
            # Validation method should show its own error
            pass

    def validate_form(self):
        """Placeholder for validation logic. Returns True if valid."""
        return True

    def get_data(self):
        """Placeholder for data retrieval. Returns a dict."""
        return {key: w.currentText() if isinstance(w, QComboBox) else w.text() for key, w in self.widgets.items()}


class ManagementDialog(BaseDialog):
    def __init__(self, parent, title, fields, initial_data=None, readonly_key=None, combobox_fields=None):
        super().__init__(parent, title)
        self.fields = fields
        self.initial_data = initial_data or {}
        self.readonly_key = readonly_key
        self.combobox_fields = combobox_fields or {}
        
        self._create_body()

    def _create_body(self):
        for key, label in self.fields:
            initial_value = self.initial_data.get(key, "")
            is_required = (key != self.readonly_key and key != 'program_area_name')
            
            if key in self.combobox_fields:
                widget = AutocompleteCombobox()
                widget.set_completion_list(self.combobox_fields[key])
                widget.setCurrentText(initial_value)
            else:
                widget = QLineEdit()
                widget.setText(initial_value)
                
            if key == self.readonly_key:
                widget.setReadOnly(True)
                
            self.add_field(key, label, widget, is_required=is_required)

    def validate_form(self):
        for key, widget in self.widgets.items():
            if isinstance(widget, AutocompleteCombobox):
                if widget.currentText() and not widget.is_valid():
                    QMessageBox.warning(self, "Input Error",
                                        f"Please select a valid {key.replace('_', ' ')} from the list.")
                    return False
        return True

    def get_data(self):
        return {key: widget.currentText() if isinstance(widget, QComboBox) else widget.text() 
                for key, widget in self.widgets.items()}


class SectionDialog(BaseDialog):
    def __init__(self, parent, title, data_manager, initial_data=None):
        self.data_manager = data_manager
        self.initial_data = initial_data or {}
        super().__init__(parent, title)
        self._create_body()
        self.update_course_options() # Initial population
        
        # Load initial data
        if self.initial_data.get('course_id_portion'):
            course = self.data_manager.courses.get(self.initial_data['course_id_portion'])
            if course: 
                self.widgets['course'].setCurrentText(f"{course.short_name} ({self.initial_data['course_id_portion']})")
        if self.initial_data.get('term_name'):
            term = self.data_manager.terms.get(self.initial_data['term_name'])
            if term: 
                self.widgets['term'].setCurrentText(f"{term.name} ({term.term_id})")
        if self.initial_data.get('account_id'):
            self.widgets['account'].setCurrentText(self.initial_data['account_id'])
        if self.initial_data.get('section_number'):
            self.widgets['section_number'].setText(self.initial_data['section_number'])
        if self.initial_data.get('status'):
            self.widgets['status'].setCurrentText(self.initial_data['status'])
        if self.initial_data.get('start_date'):
            self.widgets['start_date'].setText(self.initial_data['start_date'])
        if self.initial_data.get('end_date'):
            self.widgets['end_date'].setText(self.initial_data['end_date'])

    def _create_body(self):
        pa_values = ["All Program Areas"] + sorted(list(self.data_manager.program_areas.keys()))
        pa_combo = QComboBox()
        pa_combo.addItems(pa_values)
        pa_combo.currentTextChanged.connect(self.update_course_options)
        self.add_field('pa_filter', "Filter by Program Area", pa_combo)

        course_combo = AutocompleteCombobox()
        self.add_field('course', "Course", course_combo, is_required=True)

        term_display = [f"{t.name} ({t.term_id})" for tname, t in self.data_manager.terms.items()]
        term_combo = AutocompleteCombobox()
        term_combo.set_completion_list(term_display)
        self.add_field('term', "Term", term_combo, is_required=True)

        account_combo = AutocompleteCombobox()
        account_combo.set_completion_list(list(self.data_manager.accounts.keys()))
        self.add_field('account', "Account", account_combo, is_required=True)

        self.add_field('section_number', "Section Number", QLineEdit(), is_required=True)
        
        status_combo = QComboBox()
        status_combo.addItems(['active', 'deleted', 'completed', 'published'])
        self.add_field('status', "Status", status_combo, is_required=True)
        
        self.add_field('start_date', "Start Date (YYYY-MM-DD)", QLineEdit())
        self.add_field('end_date', "End Date (YYYY-MM-DD)", QLineEdit())

    def update_course_options(self, selected_pa="All Program Areas"):
        selected_pa = self.widgets['pa_filter'].currentText()
        course_options = []
        for cid, course in sorted(self.data_manager.courses.items(), key=lambda item: item[1].short_name):
            if selected_pa == "All Program Areas" or course.program_area_name == selected_pa:
                course_options.append(f"{course.short_name} ({cid})")
        
        current_selection = self.widgets['course'].currentText()
        self.widgets['course'].set_completion_list(course_options)
        
        # Try to preserve selection if it's still in the list
        if current_selection in course_options:
            self.widgets['course'].setCurrentText(current_selection)

    def validate_form(self):
        if not self.widgets['course'].currentText() or not self.widgets['course'].is_valid():
            QMessageBox.warning(self, "Input Error", "A valid course must be selected from the list.")
            return False
        if not self.widgets['term'].currentText() or not self.widgets['term'].is_valid():
            QMessageBox.warning(self, "Input Error", "A valid term must be selected from the list.")
            return False
        if not self.widgets['account'].currentText() or not self.widgets['account'].is_valid():
            QMessageBox.warning(self, "Input Error", "A valid account must be selected from the list.")
            return False
        if not self.widgets['section_number'].text():
            QMessageBox.warning(self, "Input Error", "Section Number is a required field.")
            return False
        return True

    def get_data(self):
        course_str = self.widgets['course'].currentText()
        course_id = course_str[course_str.rfind('(') + 1:-1] if '(' in course_str else ''
        
        term_str = self.widgets['term'].currentText()
        term_name = term_str.rsplit(' (', 1)[0] if ' (' in term_str else ''
        
        return {
            'course_id_portion': course_id,
            'term_name': term_name,
            'account_id': self.widgets['account'].currentText(),
            'section_number': self.widgets['section_number'].text(),
            'status': self.widgets['status'].currentText(),
            'start_date': self.widgets['start_date'].text(),
            'end_date': self.widgets['end_date'].text()
        }


class EnrollmentDialog(QDialog):
    def __init__(self, parent, title, section, data_manager):
        super().__init__(parent)
        self.section = section
        self.data_manager = data_manager
        
        self.setWindowTitle(title)
        self.setMinimumSize(600, 400)

        layout = QVBoxLayout(self)

        # --- Enrollment List ---
        self.tree = QTreeWidget()
        self.tree.setColumnCount(4)
        self.tree.setHeaderLabels(['User ID', 'Name', 'Role', 'Status'])
        self.tree.header().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.tree)
        
        # --- Add Enrollment Form ---
        add_frame = QGroupBox("Add New Enrollment")
        form_layout = QFormLayout(add_frame)

        pa_values = ["All Program Areas"] + sorted(list(self.data_manager.program_areas.keys()))
        self.pa_combo = QComboBox()
        self.pa_combo.addItems(pa_values)
        self.pa_combo.currentTextChanged.connect(self.update_person_options)
        form_layout.addRow("Filter by Program Area:", self.pa_combo)

        self.person_combo = AutocompleteCombobox()
        form_layout.addRow(QLabel('Person: <span style="color:red;">*</span>'), self.person_combo)
        
        self.role_combo = QComboBox()
        self.role_combo.addItems(sorted(list(self.data_manager.enrollment_roles.keys())))
        form_layout.addRow(QLabel('Role: <span style="color:red;">*</span>'), self.role_combo)
        
        self.status_combo = QComboBox()
        self.status_combo.addItems(['active', 'completed', 'inactive', 'deleted'])
        form_layout.addRow(QLabel('Status: <span style="color:red;">*</span>'), self.status_combo)
        
        add_btn = QPushButton("Add")
        add_btn.clicked.connect(self.add_enrollment)
        form_layout.addRow(add_btn)
        
        layout.addWidget(add_frame)
        
        # --- Bottom Buttons ---
        bottom_layout = QHBoxLayout()
        delete_btn = QPushButton("Delete Selected")
        delete_btn.clicked.connect(self.delete_enrollment)
        bottom_layout.addWidget(delete_btn)
        bottom_layout.addStretch()
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        bottom_layout.addWidget(close_btn)
        
        layout.addLayout(bottom_layout)
        
        # --- Initial Load ---
        self.update_person_options()
        self.refresh_enrollments()

    def update_person_options(self, event=None):
        selected_pa = self.pa_combo.currentText()
        people_options = []
        for uid, person in sorted(self.data_manager.people.items(), key=lambda item: item[1].name):
            if selected_pa == "All Program Areas" or person.program_area_name == selected_pa:
                people_options.append(f"{person.name} ({uid})")
        self.person_combo.set_completion_list(people_options)

    def refresh_enrollments(self):
        self.tree.clear()
        for i, enroll in enumerate(self.section.enrollments):
            person = self.data_manager.people.get(enroll.user_id, Person("Unknown", enroll.user_id))
            item = QTreeWidgetItem([str(enroll.user_id), person.name, enroll.role, enroll.status])
            item.setData(0, Qt.ItemDataRole.UserRole, i) # Store list index
            self.tree.addTopLevelItem(item)

    def add_enrollment(self):
        person_str = self.person_combo.currentText()
        role = self.role_combo.currentText()
        
        if not person_str or not self.person_combo.is_valid():
            QMessageBox.warning(self, "Input Error", "Please select a valid person from the list.")
            return
        if not role:
            QMessageBox.warning(self, "Input Error", "Please select a role.")
            return
            
        user_id = person_str[person_str.rfind('(') + 1:-1]
        status = self.status_combo.currentText()
        
        # Check for duplicates
        if any(e.user_id == user_id for e in self.section.enrollments):
            QMessageBox.warning(self, "Duplicate Error", "This user is already enrolled in this section.")
            return

        enrollment = Enrollment(user_id, role, status)
        self.section.add_enrollment(enrollment)
        self.refresh_enrollments()
        self.data_manager.save_data() # Save changes immediately

    def delete_enrollment(self):
        selected = self.tree.selectedItems()
        if not selected:
            return
            
        reply = QMessageBox.question(self, "Confirm Delete", "Delete selected enrollment(s)?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                                     
        if reply == QMessageBox.StandardButton.Yes:
            # Get indices and sort in reverse
            indices_to_delete = sorted([item.data(0, Qt.ItemDataRole.UserRole) for item in selected], reverse=True)
            for index in indices_to_delete:
                del self.section.enrollments[index]
            self.refresh_enrollments()
            self.data_manager.save_data() # Save changes immediately


class ExportDialog(BaseDialog):
    def __init__(self, parent):
        super().__init__(parent, "Export to CSV Files")
        self.data_types = ['people', 'courses', 'program_areas', 'terms', 'accounts']
        self.dir_path = ""
        self._create_body()

    def _create_body(self):
        # Remove the default form layout widget
        self.main_layout.removeWidget(self.form_widget)
        self.form_widget.deleteLater()

        options_frame = QGroupBox("Select Data to Export")
        options_layout = QVBoxLayout(options_frame)
        self.widgets = {}
        for data_type in self.data_types:
            cb = QCheckBox(data_type.replace('_', ' ').title())
            cb.setChecked(True)
            options_layout.addWidget(cb)
            self.widgets[data_type] = cb
        self.main_layout.insertWidget(0, options_frame)

        dir_frame = QGroupBox("Select Export Location")
        dir_layout = QHBoxLayout(dir_frame)
        self.dir_label = QLineEdit("No directory selected")
        self.dir_label.setReadOnly(True)
        dir_button = QPushButton("Choose...")
        dir_button.clicked.connect(self.choose_directory)
        dir_layout.addWidget(self.dir_label)
        dir_layout.addWidget(dir_button)
        self.main_layout.insertWidget(1, dir_frame)

    def choose_directory(self):
        path = QFileDialog.getExistingDirectory(self, "Select Folder to Save CSV Files")
        if path:
            self.dir_path = path
            self.dir_label.setText(path)
            
    def validate_form(self):
        selected_types = [dt for dt, var in self.widgets.items() if var.isChecked()]
        if not selected_types:
            QMessageBox.warning(self, "Export Error", "Please select at least one data type to export.")
            return False
        if not self.dir_path:
            QMessageBox.warning(self, "Export Error", "Please select a directory to save the files.")
            return False
        return True

    def get_data(self):
        selected_types = [dt for dt, var in self.widgets.items() if var.isChecked()]
        return {'types': selected_types, 'directory': self.dir_path}


class ImportDialog(BaseDialog):
    def __init__(self, parent):
        super().__init__(parent, "Import from CSV Files")
        self.data_types = ['people', 'courses', 'program_areas', 'terms', 'accounts']
        self.file_paths = {}
        self._create_body()

    def _create_body(self):
        self.widgets = {}
        for data_type in self.data_types:
            path_widget = QWidget()
            path_layout = QHBoxLayout(path_widget)
            path_layout.setContentsMargins(0,0,0,0)
            
            entry = QLineEdit("No file selected")
            entry.setReadOnly(True)
            btn = QPushButton("Choose File...")
            # Use partial to pass data_type to the slot
            btn.clicked.connect(partial(self.choose_file, data_type))
            
            path_layout.addWidget(entry)
            path_layout.addWidget(btn)
            
            self.add_field(data_type, data_type.replace('_', ' ').title(), path_widget)
            self.widgets[data_type] = entry # Store the entry to update its text

    def choose_file(self, data_type):
        path, _ = QFileDialog.getOpenFileName(self, f"Select {data_type.title()} CSV File", "", "CSV Files (*.csv)")
        if path:
            self.file_paths[data_type] = path
            self.widgets[data_type].setText(os.path.basename(path))
            
    def validate_form(self):
        if not self.file_paths:
            QMessageBox.warning(self, "Import Error", "No files were selected to import.")
            return False
        return True

    def get_data(self):
        return self.file_paths


class RoleEditDialog(BaseDialog):
    def __init__(self, parent, title, initial_data=None, is_new=False):
        super().__init__(parent, title)
        self.initial_data = initial_data or {}
        self.is_new = is_new
        self._create_body()

    def _create_body(self):
        display_name_entry = QLineEdit()
        display_name_entry.setText(self.initial_data.get('display_name', ''))
        if not self.is_new:
            display_name_entry.setReadOnly(True)
        self.add_field('display_name', "Display Name", display_name_entry, is_required=True)
        
        canvas_role_entry = QLineEdit()
        canvas_role_entry.setText(self.initial_data.get('canvas_role', ''))
        self.add_field('canvas_role', "Canvas Role Value", canvas_role_entry, is_required=True)
        
    def validate_form(self):
        if not self.widgets['display_name'].text() or not self.widgets['canvas_role'].text():
            QMessageBox.warning(self, "Input Error", "Both fields are required.")
            return False
        return True
    
    def get_data(self):
        return {
            "display_name": self.widgets['display_name'].text(),
            "canvas_role": self.widgets['canvas_role'].text()
        }

# --- Main Execution (PyQt6 Version) ---
if __name__ == "__main__":
    # Create the application instance
    qapp = QApplication(sys.argv)
    
    # Initialize the data manager
    dm = DataManager()
    
    # Create and show the main window
    main_window = App(dm)
    main_window.show()
    
    # Start the application's event loop
    sys.exit(qapp.exec())