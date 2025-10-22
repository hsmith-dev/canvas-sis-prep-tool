import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import json
import csv
import os
import webbrowser
import sys
import shutil
import inspect
from appdirs import user_data_dir


# --- Helper function for finding assets ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# --- Function to get a writable application data directory ---
def get_app_data_path(app_name, app_author):
    """ Get a writable, cross-platform path for application data. """
    path = user_data_dir(app_name, app_author) # appdirs handles all OS logic
    if not os.path.exists(path):
        os.makedirs(path)
    return path

# --- Data File ---
APP_NAME = "CanvasSISPrepTool"
APP_AUTHOR = "Harrison Smith"
DATA_DIR = get_app_data_path(APP_NAME, APP_AUTHOR)
DATA_FILE = os.path.join(DATA_DIR, 'course_data.json')


# --- Core Data Models ---
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


# --- DataManager ---
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


# --- Custom Autocomplete Combobox ---
class AutocompleteCombobox(ttk.Combobox):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        # Use a different name for the full list to avoid confusion with the widget's 'values'
        self._full_completion_list = []
        self.set_completion_list(self.cget('values'))  # Initialize with any values passed at creation
        self.bind('<KeyRelease>', self.handle_keyrelease)
        self.bind('<<ComboboxSelected>>', self.handle_selection)
        self.bind("<FocusOut>", self.handle_focusout)

    def set_completion_list(self, new_list):
        """Public method to update the master list of options."""
        self._full_completion_list = sorted(list(new_list))
        # Initially, the displayed list is the full list
        self['values'] = self._full_completion_list

    def configure(self, cnf=None, **kw):
        """Override configure to properly handle 'values' updates."""
        if 'values' in kw:
            # When 'values' is configured, update our master list
            self.set_completion_list(kw['values'])
            if self.get() not in self._full_completion_list:
                self.set('')
        super().configure(cnf, **kw)

    def handle_keyrelease(self, event):
        """Filter the list as the user types and show the dropdown."""
        value = self.get()
        if value == '':
            self['values'] = self._full_completion_list
        else:
            filtered_values = [item for item in self._full_completion_list if value.lower() in item.lower()]
            if filtered_values:
                # Update the displayed list and pop open the dropdown
                self['values'] = filtered_values
                self.event_generate('<Down>')
            else:
                # If no matches, show the full list so they can see available options
                self['values'] = self._full_completion_list
        # This is needed to keep the dropdown open while typing
        self.focus_set()
        self.icursor(tk.END)

    def handle_selection(self, event):
        """Reset the full list in the dropdown after a user selects an item."""
        self['values'] = self._full_completion_list

    def handle_focusout(self, event):
        """When the widget loses focus, reset the value list for the next interaction."""
        self['values'] = self._full_completion_list

    def is_valid(self):
        """Check if the current value is an exact match in the master completion list."""
        return self.get() in self._full_completion_list


# --- GUI Application ---
class App(tk.Tk):
    def __init__(self, data_manager):
        super().__init__()
        self.sections_tree = None
        self.accounts_tree = None
        self.terms_tree = None
        self.program_areas_tree = None
        self.courses_tree = None
        self.people_tree = None
        self.roles_management_frame = None
        self.data_manager = data_manager
        self.title("Canvas SIS Prep Tool")
        self.geometry("1350x750")
        try:
            icon = tk.PhotoImage(file=resource_path('app_icon.png'))
            self.iconphoto(False, icon)
        except tk.TclError:
            print("app_icon.png not found, skipping icon.")
        self.current_theme = "light"
        self.setup_styles()
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, expand=True, fill="both")

        self.create_sections_tab()
        self.create_settings_tab()
        self.create_people_tab()
        self.create_courses_tab()
        self.create_program_areas_tab()
        self.create_terms_tab()
        self.create_accounts_tab()
        self.create_about_tab()

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def get_theme_colors(self):
        light_theme = {"accent": "#D94F4F", "secondary": "#F2F2F2", "contrast": "#222222", "highlight": "#4D9FE0",
                       "surface": "#FFFFFF", "tree_heading": "#E0E0E0", "dialog_bg": "#F2F2F2", "button_bg": "#E1E1E1"}
        dark_theme = {"accent": "#D94F4F", "secondary": "#212121", "contrast": "#F5F5F5", "highlight": "#4D9FE0",
                      "surface": "#2C2C2C", "tree_heading": "#383838", "dialog_bg": "#333333", "button_bg": "#4F4F4F"}
        return dark_theme if self.current_theme == "dark" else light_theme

    def toggle_theme(self):
        self.current_theme = "dark" if self.current_theme == "light" else "light"
        self.setup_styles()

    def setup_styles(self):
        colors = self.get_theme_colors()
        style = ttk.Style(self)
        style.theme_use("clam")
        self.configure(background=colors["secondary"])
        style.configure("Accent.TButton", background=colors["accent"], foreground="#FFFFFF",
                        font=('TkDefaultFont', 10, 'bold'), borderwidth=1, relief="raised")
        style.map("Accent.TButton", background=[('active', '#B84444')])
        style.configure("TButton", background=colors["button_bg"], foreground=colors["contrast"], borderwidth=1,
                        relief="raised")
        style.map("TButton", background=[('active', colors["highlight"])])
        style.configure("TNotebook", background=colors["secondary"], borderwidth=0)
        style.configure("TNotebook.Tab", background=colors["secondary"], foreground=colors["contrast"],
                        lightcolor=colors["secondary"], padding=[10, 5])
        style.map("TNotebook.Tab", background=[("selected", colors["surface"])],
                  foreground=[("selected", colors["accent"])])
        style.configure("TFrame", background=colors["secondary"])
        style.configure("TLabelFrame", background=colors["secondary"], borderwidth=1, relief="groove")
        style.configure("TLabelFrame.Label", background=colors["secondary"], foreground=colors["contrast"],
                        font=('TkDefaultFont', 10, 'bold'))
        style.configure("TLabel", background=colors["secondary"], foreground=colors["contrast"])
        style.configure("Required.TLabel", foreground="red", background=colors["secondary"])
        style.configure("link.TLabel", foreground=colors["highlight"], background=colors["secondary"],
                        font=('TkDefaultFont', 10, 'underline'))
        style.configure("Treeview", background=colors["surface"], foreground=colors["contrast"],
                        fieldbackground=colors["surface"], rowheight=25)
        style.map("Treeview", background=[('selected', colors["highlight"])], foreground=[('selected', "#FFFFFF")])
        style.configure("Treeview.Heading", background=colors["tree_heading"], foreground=colors["contrast"],
                        font=('TkDefaultFont', 10, 'bold'))
        style.map('TCombobox', fieldbackground=[('readonly', colors["surface"])],
                  selectbackground=[('readonly', colors["highlight"])],
                  selectforeground=[('readonly', colors["contrast"])], foreground=[('readonly', colors["contrast"])])

    def on_closing(self):
        self.data_manager.save_data()
        self.destroy()

    def create_management_tab(self, parent, title, columns, item_name, edit_command):
        frame = ttk.Frame(parent, padding="10")
        parent.add(frame, text=title)
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(expand=True, fill="both")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col.replace('_', ' ').title())
            tree.column(col, width=120)
        tree.pack(side="left", expand=True, fill="both")
        tree.bind("<Double-1>", lambda event: edit_command(tree))
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        scrollbar.pack(side="right", fill="y")
        tree.configure(yscrollcommand=scrollbar.set)
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text=f"Add {item_name}", command=lambda: self.add_item(item_name, tree)).pack(side="left",
                                                                                                            padx=5)
        ttk.Button(btn_frame, text=f"Edit {item_name}", command=lambda: edit_command(tree)).pack(side="left", padx=5)
        ttk.Button(btn_frame, text=f"Delete {item_name}", command=lambda: self.delete_item(item_name, tree)).pack(
            side="left", padx=5)
        return tree

    def refresh_all_views(self):
        self.refresh_people_list()
        self.refresh_courses_list()
        self.refresh_program_areas_list()
        self.refresh_terms_list()
        self.refresh_accounts_list()
        self.refresh_sections_list()

    def create_people_tab(self):
        self.people_tree = self.create_management_tab(self.notebook, "People", ('user_id', 'name', 'program_area_name'),
                                                      "Person", lambda tree: self.edit_item("Person", tree))
        self.refresh_people_list()

    def refresh_people_list(self):
        for item in self.people_tree.get_children(): self.people_tree.delete(item)
        for user_id, person in sorted(self.data_manager.people.items()):
            self.people_tree.insert("", "end", values=(user_id, person.name, person.program_area_name))

    def create_courses_tab(self):
        self.courses_tree = self.create_management_tab(self.notebook, "Courses",
                                                       ('course_id_portion', 'short_name', 'long_name',
                                                        'program_area_name'), "Course",
                                                       lambda tree: self.edit_item("Course", tree))
        self.refresh_courses_list()

    def refresh_courses_list(self):
        for item in self.courses_tree.get_children(): self.courses_tree.delete(item)
        for cid, course in sorted(self.data_manager.courses.items()):
            self.courses_tree.insert("", "end",
                                     values=(cid, course.short_name, course.long_name, course.program_area_name))

    def create_program_areas_tab(self):
        self.program_areas_tree = self.create_management_tab(self.notebook, "Program Areas", ('name',), "Program Area",
                                                             lambda tree: self.edit_item("Program Area", tree))
        self.refresh_program_areas_list()

    def refresh_program_areas_list(self):
        for item in self.program_areas_tree.get_children(): self.program_areas_tree.delete(item)
        for name, dept in sorted(self.data_manager.program_areas.items()):
            self.program_areas_tree.insert("", "end", values=(name,))

    def create_terms_tab(self):
        self.terms_tree = self.create_management_tab(self.notebook, "Terms", ('name', 'term_id', 'short_code'), "Term",
                                                     lambda tree: self.edit_item("Term", tree))
        self.refresh_terms_list()

    def create_accounts_tab(self):
        self.accounts_tree = self.create_management_tab(self.notebook, "Accounts", ('account_id',), "Account",
                                                        lambda tree: self.edit_item("Account", tree))
        self.refresh_accounts_list()

    def refresh_terms_list(self):
        for item in self.terms_tree.get_children(): self.terms_tree.delete(item)
        for tname, term in sorted(self.data_manager.terms.items()):
            self.terms_tree.insert("", "end", values=(tname, term.term_id, term.short_code))

    def refresh_accounts_list(self):
        for item in self.accounts_tree.get_children(): self.accounts_tree.delete(item)
        for aid, acc in sorted(self.data_manager.accounts.items()):
            self.accounts_tree.insert("", "end", values=(aid,))

    def create_sections_tab(self):
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Sections & Enrollments")
        self.sections_tree = ttk.Treeview(frame,
                                          columns=('course', 'term', 'section', 'status', 'start_date', 'end_date'),
                                          show="headings")
        self.sections_tree.heading('course', text='Course');
        self.sections_tree.heading('term', text='Term');
        self.sections_tree.heading('section', text='Section #');
        self.sections_tree.heading('status', text='Status');
        self.sections_tree.heading('start_date', text='Start Date');
        self.sections_tree.heading('end_date', text='End Date')
        self.sections_tree.pack(expand=True, fill="both")
        self.sections_tree.bind("<Double-1>", lambda event: self.edit_section())
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="Create Section", command=self.create_section).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Edit Section", command=self.edit_section).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Manage Enrollments", command=self.manage_enrollments).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Delete Section", command=self.delete_section).pack(side="left", padx=5)
        self.refresh_sections_list()

    def refresh_sections_list(self):
        for item in self.sections_tree.get_children(): self.sections_tree.delete(item)
        for i, sec in enumerate(self.data_manager.sections):
            course = self.data_manager.courses.get(sec.course_id_portion)
            course_name = course.short_name if course else 'N/A'
            term_obj = self.data_manager.terms.get(sec.term_name)
            term_display = f"{term_obj.name} ({term_obj.term_id})" if term_obj else 'N/A'
            self.sections_tree.insert("", "end", iid=i,
                                      values=(f"{course_name} ({sec.course_id_portion})", term_display,
                                              sec.section_number, sec.status, sec.start_date, sec.end_date))

    def create_section(self):
        dialog = SectionDialog(self, "Create Section", self.data_manager, self.get_theme_colors())
        if dialog.result:
            new_section = Section(**dialog.result)
            self.data_manager.sections.append(new_section)
            self.refresh_sections_list()
            self.data_manager.save_data()
            section_index = len(self.data_manager.sections) - 1
            self.sections_tree.selection_set(str(section_index))
            self.manage_enrollments()

    def edit_section(self, event=None):
        selected = self.sections_tree.selection()
        if not selected:
            if event: return
            messagebox.showwarning("Selection Error", "Please select a section to edit.");
            return
        section_index = int(selected[0]);
        section_obj = self.data_manager.sections[section_index]
        dialog = SectionDialog(self, "Edit Section", self.data_manager, self.get_theme_colors(),
                               initial_data=section_obj.to_dict())
        if dialog.result:
            for key, value in dialog.result.items(): setattr(section_obj, key, value)
            self.refresh_sections_list()
            self.data_manager.save_data()

    def manage_enrollments(self):
        selected = self.sections_tree.selection()
        if not selected:
            messagebox.showwarning("Selection Error", "Please select a section to manage.", parent=self)
            return
        section_index = int(selected[0])
        section = self.data_manager.sections[section_index]
        term_obj = self.data_manager.terms.get(section.term_name)
        if term_obj:
            full_section_id = f"{term_obj.short_code}-{section.course_id_portion}-{section.section_number}"
            title = f"Enrollments for {full_section_id}"
        else:
            title = f"Enrollments for {section.course_id_portion}"

        EnrollmentDialog(self, title, section, self.data_manager, self.get_theme_colors())

    def delete_section(self):
        selected = self.sections_tree.selection()
        if not selected: messagebox.showwarning("Selection Error", "Please select a section to delete."); return
        if messagebox.askyesno("Confirm Delete",
                               "Are you sure you want to delete this section and all its enrollments?"):
            indices_to_delete = sorted([int(s) for s in selected], reverse=True)
            for index in indices_to_delete: del self.data_manager.sections[index]
            self.refresh_sections_list()
            self.data_manager.save_data()

    def generate_csv(self):
        prefix = simpledialog.askstring("File Prefix", "Enter an optional prefix for the CSV files:", parent=self);
        if prefix is None: return
        directory = filedialog.askdirectory(title="Select Folder to Save CSV Files");
        if not directory: return
        result = self.data_manager.generate_csv_files(directory, prefix);
        messagebox.showinfo("CSV Generation", result)

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

        dialog = ManagementDialog(self, f"Add {item_name}", fields_map[item_name], theme_colors=self.get_theme_colors(),
                                  combobox_fields=combobox_fields)
        if not dialog.result: return

        for key, value in dialog.result.items():
            if not value and key != 'program_area_name':
                messagebox.showerror("Input Error", f"{key.replace('_', ' ').title()} cannot be empty.", parent=self)
                return

        key = dialog.result[key_field]
        data_map = {"Person": self.data_manager.people, "Course": self.data_manager.courses,
                    "Term": self.data_manager.terms, "Account": self.data_manager.accounts,
                    "Program Area": self.data_manager.program_areas}
        if key in data_map[item_name]:
            messagebox.showerror("Error", f"A {item_name.lower()} with this ID/Name already exists.", parent=self)
            return

        constructors = {"Person": Person, "Course": Course, "Term": Term, "Account": Account,
                        "Program Area": ProgramArea}
        data_map[item_name][key] = constructors[item_name](**dialog.result)

        refresh_map = {"Person": self.refresh_people_list, "Course": self.refresh_courses_list,
                       "Term": self.refresh_terms_list, "Account": self.refresh_accounts_list,
                       "Program Area": self.refresh_program_areas_list}
        refresh_map[item_name]()
        self.data_manager.save_data()

    def edit_item(self, item_name, tree):
        selected = tree.selection()
        if not selected: messagebox.showwarning("Selection Error",
                                                f"Please select a {item_name.lower()} to edit."); return

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

        # BUG FIX: Ensure the key is a string, as tkinter might convert numeric strings to integers
        key = str(tree.item(selected[0])['values'][key_index])

        combobox_fields = None
        if item_name in ["Person", "Course"]:
            program_areas = [""] + sorted(list(self.data_manager.program_areas.keys()))
            combobox_fields = {'program_area_name': program_areas}

        data_map = {"Person": self.data_manager.people, "Course": self.data_manager.courses,
                    "Term": self.data_manager.terms, "Account": self.data_manager.accounts,
                    "Program Area": self.data_manager.program_areas}
        item_obj = data_map[item_name][key]

        dialog = ManagementDialog(self, f"Edit {item_name}", fields_map[item_name], initial_data=item_obj.to_dict(),
                                  readonly_key=key_field, theme_colors=self.get_theme_colors(),
                                  combobox_fields=combobox_fields)
        if not dialog.result: return

        for k, value in dialog.result.items():
            if k != key_field and not value and k != 'program_area_name':
                messagebox.showerror("Input Error", f"{k.replace('_', ' ').title()} cannot be empty.", parent=self)
                return

        new_key = dialog.result[key_field]
        if new_key != key:  # Key has changed
            if new_key in data_map[item_name]:
                messagebox.showerror("Error", f"A {item_name.lower()} with this ID/Name already exists.");
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
        data_map[item_name][new_key] = constructors[item_name](**dialog.result)

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
        selected = tree.selection()
        if not selected: messagebox.showwarning("Selection Error",
                                                f"Please select a {item_name.lower()} to delete."); return

        fields_map = {"Person": [('user_id',)], "Course": [('course_id_portion',)], "Term": [('name',)],
                      "Account": [('account_id',)], "Program Area": [('name',)]}
        key_field = fields_map[item_name][0][0]
        key_index = tree['columns'].index(key_field)

        # BUG FIX: Ensure the key is a string, as tkinter might convert numeric strings to integers
        key = str(tree.item(selected[0])['values'][key_index])

        in_use_msg = f"Cannot delete this {item_name.lower()}. It is in use."
        if item_name == "Person" and any(e.user_id == key for s in self.data_manager.sections for e in s.enrollments):
            messagebox.showerror("Deletion Error", in_use_msg);
            return
        if item_name == "Program Area":
            if any(p.program_area_name == key for p in self.data_manager.people.values()) or any(
                    c.program_area_name == key for c in self.data_manager.courses.values()):
                messagebox.showerror("Deletion Error", in_use_msg);
                return
        if item_name in ["Course", "Term", "Account"]:
            check_key = {"Course": "course_id_portion", "Term": "term_name", "Account": "account_id"}[item_name]
            if any(getattr(s, check_key, None) == key for s in self.data_manager.sections):
                messagebox.showerror("Deletion Error", in_use_msg);
                return

        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete this {item_name.lower()}?"):
            data_map = {"Person": self.data_manager.people, "Course": self.data_manager.courses,
                        "Term": self.data_manager.terms, "Account": self.data_manager.accounts,
                        "Program Area": self.data_manager.program_areas}
            if key in data_map[item_name]: del data_map[item_name][key]

            refresh_map = {"Person": self.refresh_people_list, "Course": self.refresh_courses_list,
                           "Term": self.refresh_terms_list, "Account": self.refresh_accounts_list,
                           "Program Area": self.refresh_program_areas_list}
            refresh_map[item_name]()
            self.data_manager.save_data()

    def create_settings_tab(self):
        actions_tab_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(actions_tab_frame, text="Actions")

        # --- Main Actions View ---
        self.main_actions_frame = ttk.Frame(actions_tab_frame)
        self.main_actions_frame.pack(fill='both', expand=True)

        canvas_frame = ttk.LabelFrame(self.main_actions_frame, text="Canvas SIS Files", padding=10)
        canvas_frame.pack(fill="x", pady=(0, 10))
        ttk.Button(canvas_frame, text="Generate Canvas CSV Files...", style="Accent.TButton",
                   command=self.generate_csv).pack(pady=10, ipady=5, fill='x')

        data_frame = ttk.LabelFrame(self.main_actions_frame, text="Application Data", padding=10)
        data_frame.pack(fill="x", pady=10)
        ttk.Button(data_frame, text="Import Data from CSVs...", command=self.open_import_dialog).pack(pady=5, ipady=5,
                                                                                                      fill='x')
        ttk.Button(data_frame, text="Export Data to CSVs...", command=self.open_export_dialog).pack(pady=5, ipady=5,
                                                                                                    fill='x')

        danger_frame = ttk.LabelFrame(self.main_actions_frame, text="Danger Zone", padding=10)
        danger_frame.pack(fill="x", pady=10)
        ttk.Button(danger_frame, text="Manage Enrollment Roles", command=self.show_roles_management_view).pack(
            pady=10, ipady=5, fill='x')
        ttk.Button(danger_frame, text="Clear All Local Data", style="Accent.TButton", command=self.clear_all_data).pack(
            pady=10, ipady=5, fill='x')

        # --- Roles Management View (initially hidden) ---
        self.roles_management_frame = ttk.Frame(actions_tab_frame)
        self.setup_roles_management_view()

    def setup_roles_management_view(self):
        # This frame is a child of the main actions tab frame
        frame = self.roles_management_frame

        ttk.Button(frame, text="< Back to Actions", command=self.show_main_actions_view).pack(anchor="w", pady=(0, 10))

        tree_frame = ttk.LabelFrame(frame, text="Enrollment Roles")
        tree_frame.pack(fill="both", expand=True)

        self.roles_tree = ttk.Treeview(tree_frame, columns=('display', 'canvas_role'), show="headings")
        self.roles_tree.heading('display', text='Display Name')
        self.roles_tree.heading('canvas_role', text='Canvas Role Value')
        self.roles_tree.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.roles_tree.yview)
        scrollbar.pack(side="right", fill="y", pady=5, padx=(0, 5))
        self.roles_tree.configure(yscrollcommand=scrollbar.set)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=10)

        ttk.Button(btn_frame, text="Add Role", command=self.add_role).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Edit Role", command=self.edit_role).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Delete Role", command=self.delete_role).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Import from CSV", command=self.import_roles).pack(side="right", padx=5)

    def show_roles_management_view(self):
        self.main_actions_frame.pack_forget()
        self.refresh_roles_list()
        self.roles_management_frame.pack(fill='both', expand=True, padx=10, pady=10)

    def show_main_actions_view(self):
        self.roles_management_frame.pack_forget()
        self.main_actions_frame.pack(fill='both', expand=True)

    def refresh_roles_list(self):
        for item in self.roles_tree.get_children():
            self.roles_tree.delete(item)
        for display, canvas_role in sorted(self.data_manager.enrollment_roles.items()):
            self.roles_tree.insert("", "end", iid=display, values=(display, canvas_role))

    def add_role(self):
        dialog = RoleEditDialog(self, "Add New Role", self.get_theme_colors(), is_new=True)
        if dialog.result:
            display_name = dialog.result['display_name']
            canvas_role = dialog.result['canvas_role']
            if not display_name or not canvas_role:
                messagebox.showerror("Error", "Both fields are required.", parent=self)
                return
            if display_name in self.data_manager.enrollment_roles:
                messagebox.showerror("Error", "A role with this display name already exists.", parent=self)
                return
            self.data_manager.enrollment_roles[display_name] = canvas_role
            self.data_manager.save_data()
            self.refresh_roles_list()

    def edit_role(self):
        selected = self.roles_tree.selection()
        if not selected:
            messagebox.showwarning("Selection Error", "Please select a role to edit.", parent=self)
            return
        display_name = selected[0]
        canvas_role = self.data_manager.enrollment_roles[display_name]

        dialog = RoleEditDialog(self, "Edit Role", self.get_theme_colors(),
                                initial_data={'display_name': display_name, 'canvas_role': canvas_role})
        if dialog.result:
            self.data_manager.enrollment_roles[display_name] = dialog.result['canvas_role']
            self.data_manager.save_data()
            self.refresh_roles_list()

    def delete_role(self):
        selected = self.roles_tree.selection()
        if not selected:
            messagebox.showwarning("Selection Error", "Please select a role to delete.", parent=self)
            return
        display_name = selected[0]
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the role '{display_name}'?"):
            del self.data_manager.enrollment_roles[display_name]
            self.data_manager.save_data()
            self.refresh_roles_list()

    def import_roles(self):
        path = filedialog.askopenfilename(title="Select Roles CSV File", filetypes=[("CSV Files", "*.csv")],
                                          parent=self)
        if not path:
            return
        result = self.data_manager.import_roles_from_csv(path)
        if 'error' in result:
            messagebox.showerror("Import Error", result['error'], parent=self)
        else:
            messagebox.showinfo("Import Complete", f"Added: {result['added']}\nUpdated: {result['updated']}",
                                parent=self)
            self.refresh_roles_list()

    def open_import_dialog(self):
        dialog = ImportDialog(self, "Import from CSV Files", theme_colors=self.get_theme_colors())
        if not dialog.result: return
        summary = []
        for data_type, file_path in dialog.result.items():
            result = self.data_manager.import_from_csv_file(file_path, data_type)
            if 'error' in result:
                summary.append(f"Error importing {data_type.title()}: {result['error']}")
            else:
                summary.append(
                    f"Imported {data_type.title()}: {result['added']} added, {result['skipped']} skipped (duplicates or blank rows).")
        self.refresh_all_views()
        self.data_manager.save_data()
        messagebox.showinfo("Import Complete", "\n".join(summary))

    def open_export_dialog(self):
        dialog = ExportDialog(self, "Export to CSV Files", theme_colors=self.get_theme_colors())
        if dialog.result:
            data_types, directory = dialog.result['types'], dialog.result['directory']
            result = self.data_manager.export_data_to_csvs(data_types, directory)
            messagebox.showinfo("Export Complete", result)

    def clear_all_data(self):
        if messagebox.askyesno("Confirm Clear",
                               "Are you sure you want to delete ALL local data? This action is irreversible."):
            if self.data_manager.clear_all():
                self.refresh_all_views()
                messagebox.showinfo("Success", "All local data has been cleared.")
            else:
                messagebox.showerror("Error", "Could not clear all data. Check file permissions.")

    def create_about_tab(self):
        frame = ttk.Frame(self.notebook, padding="20");
        self.notebook.add(frame, text="About")
        content_frame = ttk.Frame(frame);
        content_frame.pack(anchor="center", expand=True)
        ttk.Label(content_frame, text="App Name:", font=('TkDefaultFont', 10, 'bold')).grid(row=0, column=0,
                                                                                            sticky="ne", padx=5, pady=5)
        ttk.Label(content_frame, text="Canvas SIS Prep Tool").grid(row=0, column=1, sticky="nw", padx=5, pady=5)
        ttk.Label(content_frame, text="Author:", font=('TkDefaultFont', 10, 'bold')).grid(row=1, column=0, sticky="ne",
                                                                                          padx=5, pady=5)
        ttk.Label(content_frame, text="Harrison Smith").grid(row=1, column=1, sticky="nw", padx=5, pady=5)
        ttk.Label(content_frame, text="AI Assistant:", font=('TkDefaultFont', 10, 'bold')).grid(row=2, column=0,
                                                                                                sticky="ne", padx=5,
                                                                                                pady=5)
        ttk.Label(content_frame, text="Gemini").grid(row=2, column=1, sticky="nw", padx=5, pady=5)
        links_frame = ttk.Frame(content_frame);
        links_frame.grid(row=1, column=2, rowspan=2, sticky="nsw", padx=20)
        github_link = ttk.Label(links_frame, text="GitHub", style="link.TLabel", cursor="hand2");
        github_link.pack(side="left", padx=5);
        github_link.bind("<Button-1>", lambda e: self.open_link("https://github.com/hsmith-dev"))
        linkedin_link = ttk.Label(links_frame, text="LinkedIn", style="link.TLabel", cursor="hand2");
        linkedin_link.pack(side="left", padx=5);
        linkedin_link.bind("<Button-1>", lambda e: self.open_link("https://linkedin.com/in/hsmith-dev"))
        email_link = ttk.Label(links_frame, text="Email", style="link.TLabel", cursor="hand2");
        email_link.pack(side="left", padx=5);
        email_link.bind("<Button-1>", lambda e: self.open_link("mailto:harrison@hsmith.dev"))
        ttk.Label(content_frame, text="Last Update Date:", font=('TkDefaultFont', 10, 'bold')).grid(row=3, column=0,
                                                                                                    sticky="ne", padx=5,
                                                                                                    pady=5)
        ttk.Label(content_frame, text="October 21, 2025").grid(row=3, column=1, sticky="nw", padx=5, pady=5)
        ttk.Label(content_frame, text="Description:", font=('TkDefaultFont', 10, 'bold')).grid(row=4, column=0,
                                                                                               sticky="ne", padx=5,
                                                                                               pady=5)
        description_text = "A desktop application designed to streamline the creation of CSV files for Canvas SIS imports currently focused at streamlining the course shell creations."
        ttk.Label(content_frame, text=description_text, wraplength=500, justify="left").grid(row=4, column=1,
                                                                                             columnspan=2, sticky="nw",
                                                                                             padx=5, pady=5)
        ttk.Button(content_frame, text="Toggle Theme", command=self.toggle_theme).grid(row=5, column=0, columnspan=3,
                                                                                       pady=20)

    def open_link(self, url):
        webbrowser.open_new(url)


# --- Custom Dialogs ---
class ManagementDialog(simpledialog.Dialog):
    def __init__(self, parent, title, fields, theme_colors, initial_data=None, readonly_key=None, combobox_fields=None):
        self.result = None
        self.fields = fields
        self.theme_colors = theme_colors
        self.initial_data = initial_data or {}
        self.readonly_key = readonly_key
        self.combobox_fields = combobox_fields or {}
        self.widgets = {}
        super().__init__(parent, title)

    def body(self, master):
        master.config(bg=self.theme_colors['dialog_bg'])
        for i, (key, label) in enumerate(self.fields):
            label_frame = ttk.Frame(master)
            label_frame.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            lbl = ttk.Label(label_frame, text=f"{label}:", background=self.theme_colors['dialog_bg'],
                            foreground=self.theme_colors['contrast'])
            lbl.pack(side=tk.LEFT)
            if key != self.readonly_key and 'program_area_name' not in key:
                req_lbl = ttk.Label(label_frame, text="*", style="Required.TLabel")
                req_lbl.pack(side=tk.LEFT)

            if key in self.combobox_fields:
                widget = AutocompleteCombobox(master, values=self.combobox_fields[key])
                widget.set(self.initial_data.get(key, ""))
            else:
                widget = ttk.Entry(master, width=30)
                widget.insert(0, self.initial_data.get(key, ""))

            widget.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
            if key == self.readonly_key:
                widget.config(state="readonly")
            self.widgets[key] = widget
        return self.widgets[self.fields[0][0]]

    def buttonbox(self):
        box = ttk.Frame(self, style="TFrame");
        box.pack(pady=10)
        ttk.Button(box, text="OK", width=10, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(box, text="Cancel", width=10, command=self.cancel).pack(side=tk.LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok);
        self.bind("<Escape>", self.cancel)

    def validate(self):
        # Check if any comboboxes have invalid entries
        for key, widget in self.widgets.items():
            if isinstance(widget, AutocompleteCombobox):
                # An empty value is fine, but if it's not empty, it must be a valid selection.
                if widget.get() and not widget.is_valid():
                    messagebox.showwarning("Input Error",
                                           f"Please select a valid {key.replace('_', ' ')} from the list.",
                                           parent=self)
                    return 0
        return 1

    def apply(self):
        self.result = {key: widget.get() for key, widget in self.widgets.items()}


class SectionDialog(simpledialog.Dialog):
    def __init__(self, parent, title, data_manager, theme_colors, initial_data=None):
        self.start_date_entry = None
        self.status_combo = None
        self.section_num_entry = None
        self.account_combo = None
        self.term_combo = None
        self.course_combo = None
        self.pa_combo = None
        self.end_date_entry = None
        self.data_manager = data_manager
        self.theme_colors = theme_colors
        self.initial_data = initial_data or {}
        self.result = None
        super().__init__(parent, title)

    def create_label_with_asterisk(self, parent, text, row, is_optional=False):
        label_frame = ttk.Frame(parent)
        label_frame.grid(row=row, column=0, sticky="w", padx=5, pady=2)
        lbl = ttk.Label(label_frame, text=f"{text}:", background=self.theme_colors['dialog_bg'],
                        foreground=self.theme_colors['contrast'])
        lbl.pack(side=tk.LEFT)
        if not is_optional:
            req_lbl = ttk.Label(label_frame, text="*", style="Required.TLabel")
            req_lbl.pack(side=tk.LEFT)

    def update_course_options(self, event=None):
        selected_pa = self.pa_combo.get()
        course_options = []
        for cid, course in sorted(self.data_manager.courses.items(), key=lambda item: item[1].short_name):
            if selected_pa == "All Program Areas" or course.program_area_name == selected_pa:
                course_options.append(f"{course.short_name} ({cid})")
        self.course_combo.configure(values=course_options)
        self.course_combo.set('')

    def body(self, master):
        master.config(bg=self.theme_colors['dialog_bg'])

        ttk.Label(master, text="Filter by Program Area:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        pa_values = ["All Program Areas"] + sorted(list(self.data_manager.program_areas.keys()))
        self.pa_combo = ttk.Combobox(master, state="readonly", values=pa_values)
        self.pa_combo.set("All Program Areas")
        self.pa_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        self.pa_combo.bind("<<ComboboxSelected>>", self.update_course_options)

        self.create_label_with_asterisk(master, "Course", 1)
        self.course_combo = AutocompleteCombobox(master)
        self.course_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        self.update_course_options()
        if self.initial_data.get('course_id_portion'):
            course = self.data_manager.courses.get(self.initial_data['course_id_portion'])
            if course: self.course_combo.set(f"{course.short_name} ({self.initial_data['course_id_portion']})")

        self.create_label_with_asterisk(master, "Term", 2)
        term_display = [f"{t.name} ({t.term_id})" for tname, t in self.data_manager.terms.items()]
        self.term_combo = AutocompleteCombobox(master, values=term_display)
        self.term_combo.grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        if self.initial_data.get('term_name'):
            term = self.data_manager.terms.get(self.initial_data['term_name'])
            if term: self.term_combo.set(f"{term.name} ({term.term_id})")

        self.create_label_with_asterisk(master, "Account", 3)
        self.account_combo = AutocompleteCombobox(master, values=list(self.data_manager.accounts.keys()))
        self.account_combo.grid(row=3, column=1, sticky="ew", padx=5, pady=2)
        self.account_combo.set(self.initial_data.get('account_id', ''))

        self.create_label_with_asterisk(master, "Section Number", 4)
        self.section_num_entry = ttk.Entry(master)
        self.section_num_entry.grid(row=4, column=1, sticky="ew", padx=5, pady=2)
        self.section_num_entry.insert(0, self.initial_data.get('section_number', ''))

        self.create_label_with_asterisk(master, "Status", 5)
        self.status_combo = ttk.Combobox(master, state="readonly",
                                         values=['active', 'deleted', 'completed', 'published'])
        self.status_combo.grid(row=5, column=1, sticky="ew", padx=5, pady=2)
        self.status_combo.set(self.initial_data.get('status', 'active'))

        self.create_label_with_asterisk(master, "Start Date (YYYY-MM-DD)", 6, is_optional=True)
        self.start_date_entry = ttk.Entry(master)
        self.start_date_entry.grid(row=6, column=1, sticky="ew", padx=5, pady=2)
        self.start_date_entry.insert(0, self.initial_data.get('start_date', ''))

        self.create_label_with_asterisk(master, "End Date (YYYY-MM-DD)", 7, is_optional=True)
        self.end_date_entry = ttk.Entry(master)
        self.end_date_entry.grid(row=7, column=1, sticky="ew", padx=5, pady=2)
        self.end_date_entry.insert(0, self.initial_data.get('end_date', ''))
        return self.pa_combo

    def buttonbox(self):
        box = ttk.Frame(self, style="TFrame");
        box.pack(pady=10)
        ttk.Button(box, text="OK", width=10, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(box, text="Cancel", width=10, command=self.cancel).pack(side=tk.LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok);
        self.bind("<Escape>", self.cancel)

    def validate(self):
        if not self.course_combo.get() or not self.course_combo.is_valid():
            messagebox.showwarning("Input Error", "A valid course must be selected from the list.", parent=self)
            return 0
        if not self.term_combo.get() or not self.term_combo.is_valid():
            messagebox.showwarning("Input Error", "A valid term must be selected from the list.", parent=self)
            return 0
        if not self.account_combo.get() or not self.account_combo.is_valid():
            messagebox.showwarning("Input Error", "A valid account must be selected from the list.", parent=self)
            return 0
        if not self.section_num_entry.get():
            messagebox.showwarning("Input Error", "Section Number is a required field.", parent=self)
            return 0
        return 1

    def apply(self):
        course_str = self.course_combo.get()
        course_id = course_str[course_str.rfind('(') + 1:-1] if '(' in course_str else ''
        term_str = self.term_combo.get()
        term_name = term_str.rsplit(' (', 1)[0] if ' (' in term_str else ''
        self.result = {'course_id_portion': course_id, 'term_name': term_name, 'account_id': self.account_combo.get(),
                       'section_number': self.section_num_entry.get(), 'status': self.status_combo.get(),
                       'start_date': self.start_date_entry.get(), 'end_date': self.end_date_entry.get()}


class EnrollmentDialog(simpledialog.Dialog):
    def __init__(self, parent, title, section, data_manager, theme_colors):
        self.person_combo = None
        self.role_combo = None
        self.pa_combo = None
        self.tree = None
        self.status_combo = None
        self.section = section;
        self.data_manager = data_manager;
        self.theme_colors = theme_colors;
        super().__init__(parent, title)

    def body(self, master):
        master.config(bg=self.theme_colors['dialog_bg'])
        self.tree = ttk.Treeview(master, columns=('user_id', 'name', 'role', 'status'), show="headings");
        self.tree.heading('user_id', text='User ID');
        self.tree.heading('name', text='Name');
        self.tree.heading('role', text='Role');
        self.tree.heading('status', text='Status');
        self.tree.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 10));
        self.refresh_enrollments()

        add_frame = ttk.LabelFrame(master, text="Add New Enrollment", padding=10);
        add_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=10)
        add_frame.columnconfigure(1, weight=1)

        ttk.Label(add_frame, text="Filter by Program Area:").grid(row=0, column=0, sticky="w", padx=5)
        pa_values = ["All Program Areas"] + sorted(list(self.data_manager.program_areas.keys()))
        self.pa_combo = ttk.Combobox(add_frame, state="readonly", values=pa_values)
        self.pa_combo.set("All Program Areas")
        self.pa_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        self.pa_combo.bind("<<ComboboxSelected>>", self.update_person_options)

        person_label_frame = ttk.Frame(add_frame);
        person_label_frame.grid(row=1, column=0, sticky="w", padx=5)
        ttk.Label(person_label_frame, text="Person:").pack(side=tk.LEFT)
        ttk.Label(person_label_frame, text="*", style="Required.TLabel").pack(side=tk.LEFT)

        self.person_combo = AutocompleteCombobox(add_frame)
        self.person_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        self.update_person_options()

        role_label_frame = ttk.Frame(add_frame);
        role_label_frame.grid(row=2, column=0, sticky="w", padx=5)
        ttk.Label(role_label_frame, text="Role:").pack(side=tk.LEFT)
        ttk.Label(role_label_frame, text="*", style="Required.TLabel").pack(side=tk.LEFT)

        self.role_combo = ttk.Combobox(add_frame, state="readonly",
                                       values=sorted(list(self.data_manager.enrollment_roles.keys())))
        self.role_combo.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        status_label_frame = ttk.Frame(add_frame);
        status_label_frame.grid(row=3, column=0, sticky="w", padx=5)
        ttk.Label(status_label_frame, text="Status:").pack(side=tk.LEFT)
        ttk.Label(status_label_frame, text="*", style="Required.TLabel").pack(side=tk.LEFT)

        self.status_combo = ttk.Combobox(add_frame, state="readonly",
                                         values=['active', 'completed', 'inactive', 'deleted']);
        self.status_combo.set('active');
        self.status_combo.grid(row=3, column=1, sticky="ew", padx=5, pady=2)

        ttk.Button(add_frame, text="Add", command=self.add_enrollment).grid(row=4, column=1, sticky="e", pady=5, padx=5)
        ttk.Button(master, text="Delete Selected", command=self.delete_enrollment).grid(row=2, column=0, sticky="w")

        return self.tree

    def update_person_options(self, event=None):
        selected_pa = self.pa_combo.get()
        people_options = []
        for uid, person in sorted(self.data_manager.people.items(), key=lambda item: item[1].name):
            if selected_pa == "All Program Areas" or person.program_area_name == selected_pa:
                people_options.append(f"{person.name} ({uid})")
        self.person_combo.configure(values=people_options)
        self.person_combo.set('')

    def refresh_enrollments(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for i, enroll in enumerate(self.section.enrollments):
            person = self.data_manager.people.get(enroll.user_id, Person("Unknown", enroll.user_id))
            self.tree.insert("", "end", iid=i, values=(enroll.user_id, person.name, enroll.role, enroll.status))

    def add_enrollment(self):
        person_str = self.person_combo.get()
        role = self.role_combo.get()
        if not person_str or not self.person_combo.is_valid():
            messagebox.showwarning("Input Error", "Please select a valid person from the list.", parent=self)
            return
        if not role:
            messagebox.showwarning("Input Error", "Please select a role.", parent=self)
            return
        user_id = person_str[person_str.rfind('(') + 1:-1]
        status = self.status_combo.get()
        enrollment = Enrollment(user_id, role, status)
        self.section.add_enrollment(enrollment)
        self.refresh_enrollments()
        self.data_manager.save_data()

    def delete_enrollment(self):
        selected = self.tree.selection()
        if not selected: return
        if messagebox.askyesno("Confirm Delete", "Delete selected enrollment(s)?"):
            indices_to_delete = sorted([int(s) for s in selected], reverse=True)
            for index in indices_to_delete: del self.section.enrollments[index]
            self.refresh_enrollments();
            self.data_manager.save_data()

    def buttonbox(self):
        box = ttk.Frame(self, style="TFrame");
        box.pack(pady=10)
        ttk.Button(box, text="Close", width=10, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok);
        self.bind("<Escape>", self.cancel)


class ExportDialog(simpledialog.Dialog):
    def __init__(self, parent, title, theme_colors):
        self.result = None
        self.theme_colors = theme_colors
        self.export_vars = {}
        self.directory_path = tk.StringVar(value="No directory selected")
        super().__init__(parent, title)

    def body(self, master):
        master.config(bg=self.theme_colors['dialog_bg'])
        options_frame = ttk.LabelFrame(master, text="Select Data to Export")
        options_frame.pack(padx=10, pady=10, fill="x")
        data_types = ['people', 'courses', 'program_areas', 'terms', 'accounts']
        for i, data_type in enumerate(data_types):
            var = tk.BooleanVar(value=True)
            self.export_vars[data_type] = var
            cb = ttk.Checkbutton(options_frame, text=data_type.replace('_', ' ').title(), variable=var)
            cb.pack(anchor="w", padx=10, pady=2)
        dir_frame = ttk.LabelFrame(master, text="Select Export Location")
        dir_frame.pack(padx=10, pady=5, fill="x")
        dir_label = ttk.Label(dir_frame, textvariable=self.directory_path, wraplength=300)
        dir_label.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        dir_button = ttk.Button(dir_frame, text="Choose...", command=self.choose_directory)
        dir_button.pack(side="right", padx=5, pady=5)

    def choose_directory(self):
        path = filedialog.askdirectory(title="Select Folder to Save CSV Files")
        if path: self.directory_path.set(path)

    def buttonbox(self):
        box = ttk.Frame(self);
        box.pack(pady=5)
        ttk.Button(box, text="Export Now", width=15, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5)
        ttk.Button(box, text="Cancel", width=10, command=self.cancel).pack(side=tk.LEFT, padx=5)
        self.bind("<Return>", self.ok);
        self.bind("<Escape>", self.cancel)

    def apply(self):
        selected_types = [dt for dt, var in self.export_vars.items() if var.get()]
        directory = self.directory_path.get()
        if not selected_types:
            messagebox.showwarning("Export Error", "Please select at least one data type to export.", parent=self)
            self.result = None;
            return
        if directory == "No directory selected":
            messagebox.showwarning("Export Error", "Please select a directory to save the files.", parent=self)
            self.result = None;
            return
        self.result = {'types': selected_types, 'directory': directory}


class ImportDialog(simpledialog.Dialog):
    def __init__(self, parent, title, theme_colors):
        self.result = None
        self.theme_colors = theme_colors;
        self.file_paths = {};
        self.path_vars = {}
        self.data_types = ['people', 'courses', 'program_areas', 'terms', 'accounts']
        super().__init__(parent, title)

    def body(self, master):
        master.config(bg=self.theme_colors['dialog_bg'])
        main_frame = ttk.Frame(master);
        main_frame.pack(padx=10, pady=10)
        for i, data_type in enumerate(self.data_types):
            label = ttk.Label(main_frame, text=f"{data_type.replace('_', ' ').title()}:");
            label.grid(row=i, column=0, sticky="w", padx=5, pady=5)
            var = tk.StringVar(value="No file selected");
            self.path_vars[data_type] = var
            entry = ttk.Entry(main_frame, textvariable=var, state="readonly", width=40);
            entry.grid(row=i, column=1, sticky="ew", padx=5)
            button = ttk.Button(main_frame, text="Choose File...", command=lambda dt=data_type: self.choose_file(dt));
            button.grid(row=i, column=2, sticky="ew", padx=5)

    def choose_file(self, data_type):
        path = filedialog.askopenfilename(title=f"Select {data_type.title()} CSV File",
                                          filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
        if path:
            self.file_paths[data_type] = path
            self.path_vars[data_type].set(os.path.basename(path))

    def buttonbox(self):
        box = ttk.Frame(self);
        box.pack(pady=5)
        ttk.Button(box, text="Import Now", width=15, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5)
        ttk.Button(box, text="Cancel", width=10, command=self.cancel).pack(side=tk.LEFT, padx=5)
        self.bind("<Return>", self.ok);
        self.bind("<Escape>", self.cancel)

    def apply(self):
        if not self.file_paths:
            messagebox.showwarning("Import Error", "No files were selected to import.", parent=self);
            self.result = None;
            return
        self.result = self.file_paths


class RoleEditDialog(simpledialog.Dialog):
    def __init__(self, parent, title, theme_colors, initial_data=None, is_new=False):
        self.canvas_role_entry = None
        self.display_name_entry = None
        self.result = None
        self.theme_colors = theme_colors
        self.initial_data = initial_data or {}
        self.is_new = is_new
        super().__init__(parent, title)

    def body(self, master):
        master.config(bg=self.theme_colors['dialog_bg'])
        ttk.Label(master, text="Display Name:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.display_name_entry = ttk.Entry(master, width=30)
        self.display_name_entry.insert(0, self.initial_data.get('display_name', ''))
        self.display_name_entry.grid(row=0, column=1, padx=5, pady=2)
        if not self.is_new:
            self.display_name_entry.config(state="readonly")

        ttk.Label(master, text="Canvas Role Value:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.canvas_role_entry = ttk.Entry(master, width=30)
        self.canvas_role_entry.insert(0, self.initial_data.get('canvas_role', ''))
        self.canvas_role_entry.grid(row=1, column=1, padx=5, pady=2)
        return self.display_name_entry

    def buttonbox(self):
        box = ttk.Frame(self, style="TFrame");
        box.pack(pady=10)
        ttk.Button(box, text="OK", width=10, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(box, text="Cancel", width=10, command=self.cancel).pack(side=tk.LEFT, padx=5, pady=5)

    def apply(self):
        self.result = {
            "display_name": self.display_name_entry.get(),
            "canvas_role": self.canvas_role_entry.get()
        }


# --- Main Execution ---
if __name__ == "__main__":
    dm = DataManager()
    app = App(dm)
    app.mainloop()