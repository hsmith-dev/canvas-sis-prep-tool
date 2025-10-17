import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import json
import csv
import os
import webbrowser

# --- Data File ---
DATA_FILE = 'course_data.json'


# --- Core Data Models ---

class Person:
    """Represents a person with a name and user_id."""

    def __init__(self, name, user_id):
        self.name = name
        self.user_id = user_id

    def to_dict(self):
        return {'name': self.name, 'user_id': self.user_id}

    @staticmethod
    def from_dict(data):
        return Person(data['name'], data['user_id'])


class Course:
    """Represents a course with a short name, long name, and ID portion."""

    def __init__(self, short_name, long_name, course_id_portion):
        self.short_name = short_name
        self.long_name = long_name
        self.course_id_portion = course_id_portion.upper()

    def to_dict(self):
        return {
            'short_name': self.short_name,
            'long_name': self.long_name,
            'course_id_portion': self.course_id_portion
        }

    @staticmethod
    def from_dict(data):
        return Course(data['short_name'], data['long_name'], data['course_id_portion'])


class Term:
    """Represents an academic term."""

    def __init__(self, term_id, name, short_code, start_date, end_date):
        self.term_id = term_id
        self.name = name
        self.short_code = short_code
        self.start_date = start_date
        self.end_date = end_date

    def to_dict(self):
        return {
            'term_id': self.term_id,
            'name': self.name,
            'short_code': self.short_code,
            'start_date': self.start_date,
            'end_date': self.end_date
        }

    @staticmethod
    def from_dict(data):
        return Term(data['term_id'], data.get('name', ''), data['short_code'], data['start_date'], data['end_date'])


class Account:
    """Represents an account ID."""

    def __init__(self, account_id):
        self.account_id = account_id

    def to_dict(self):
        return {'account_id': self.account_id}

    @staticmethod
    def from_dict(data):
        return Account(data['account_id'])


class Enrollment:
    """Represents a person's enrollment in a section."""

    def __init__(self, user_id, role, status='active'):
        self.user_id = user_id
        self.role = role
        self.status = status  # active, completed, inactive, deleted

    def to_dict(self):
        return {'user_id': self.user_id, 'role': self.role, 'status': self.status}

    @staticmethod
    def from_dict(data):
        return Enrollment(data['user_id'], data['role'], data.get('status', 'active'))


class Section:
    """Represents a specific offering of a course in a term."""

    def __init__(self, course_id_portion, term_name, account_id, section_number, status='active', start_date='',
                 end_date=''):
        self.course_id_portion = course_id_portion
        self.term_name = term_name
        self.account_id = account_id
        self.section_number = section_number
        self.status = status  # active, deleted, completed, published
        self.start_date = start_date
        self.end_date = end_date
        self.enrollments = []

    def add_enrollment(self, enrollment):
        self.enrollments.append(enrollment)

    def to_dict(self):
        return {
            'course_id_portion': self.course_id_portion,
            'term_name': self.term_name,
            'account_id': self.account_id,
            'section_number': self.section_number,
            'status': self.status,
            'start_date': self.start_date,
            'end_date': self.end_date,
            'enrollments': [e.to_dict() for e in self.enrollments]
        }

    @staticmethod
    def from_dict(data):
        section = Section(
            data['course_id_portion'],
            data.get('term_name'),  # Use .get() for safety with old files
            data['account_id'],
            data['section_number'],
            data.get('status', 'active'),
            data.get('start_date', ''),
            data.get('end_date', '')
        )
        section.enrollments = [Enrollment.from_dict(e) for e in data.get('enrollments', [])]
        return section


# --- DataManager (Handles data logic, no GUI code) ---
class DataManager:
    def __init__(self):
        self.people = {}
        self.courses = {}
        self.terms = {}
        self.accounts = {}
        self.sections = []
        self.load_data()

    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'r') as f:
                    data = json.load(f)
                    self.people = {uid: Person.from_dict(p) for uid, p in data.get('people', {}).items()}
                    self.courses = {cid: Course.from_dict(c) for cid, c in data.get('courses', {}).items()}

                    raw_terms = data.get('terms', {})
                    self.terms = {}
                    term_id_to_name_map = {}
                    for _, term_data in raw_terms.items():
                        term_obj = Term.from_dict(term_data)
                        if not term_obj.name:  # Handle old data that might not have a name
                            term_obj.name = term_obj.term_id
                        self.terms[term_obj.name] = term_obj
                        term_id_to_name_map[term_obj.term_id] = term_obj.name

                    self.accounts = {aid: Account.from_dict(a) for aid, a in data.get('accounts', {}).items()}
                    self.sections = [Section.from_dict(s) for s in data.get('sections', [])]

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

    def initialize_empty(self):
        self.people = {}
        self.courses = {}
        self.terms = {}
        self.accounts = {}
        self.sections = []

    def save_data(self):
        data = {
            'people': {uid: p.to_dict() for uid, p in self.people.items()},
            'courses': {cid: c.to_dict() for cid, c in self.courses.items()},
            'terms': {tname: t.to_dict() for tname, t in self.terms.items()},
            'accounts': {aid: a.to_dict() for aid, a in self.accounts.items()},
            'sections': [s.to_dict() for s in self.sections]
        }
        with open(DATA_FILE, 'w') as f:
            json.dump(data, f, indent=4)
        return True

    def _import_csv_data(self, file_path, required_headers, data_dict, key_field, constructor):
        added_count = 0
        skipped_count = 0
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                if not all(h in reader.fieldnames for h in required_headers):
                    return {'error': f"CSV is missing one of the required headers: {', '.join(required_headers)}"}

                for row in reader:
                    key = row.get(key_field)
                    if not key or key in data_dict:
                        skipped_count += 1
                        continue

                    args = {h: row.get(h, '') for h in required_headers}
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

    def import_terms_from_csv(self, file_path):
        return self._import_csv_data(file_path, ['name', 'term_id', 'short_code', 'start_date', 'end_date'], self.terms,
                                     'name', Term)

    def import_accounts_from_csv(self, file_path):
        return self._import_csv_data(file_path, ['account_id'], self.accounts, 'account_id', Account)

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

            effective_start_date = section.start_date if section.start_date else term_obj.start_date
            effective_end_date = section.end_date if section.end_date else term_obj.end_date

            courses_data.append({
                'course_id': course_id, 'short_name': course_obj.short_name,
                'long_name': long_name_with_section, 'account_id': section.account_id,
                'term_id': term_obj.term_id, 'status': section.status,
                'start_date': effective_start_date, 'end_date': effective_end_date
            })

            sections_data.append({
                'section_id': section_id, 'course_id': course_id,
                'name': long_name_with_section, 'status': section.status,
                'start_date': section.start_date, 'end_date': section.end_date
            })

            for enrollment in section.enrollments:
                enrollment_data.append({
                    'section_id': section_id, 'user_id': enrollment.user_id,
                    'role': enrollment.role, 'status': enrollment.status, 'course_id': course_id
                })

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


# --- GUI Application ---
class App(tk.Tk):
    def __init__(self, data_manager):
        super().__init__()
        self.data_manager = data_manager
        self.title("Canvas SIS Prep Tool")
        self.geometry("950x600")

        try:
            icon = tk.PhotoImage(file='app_icon.png')
            self.iconphoto(False, icon)
        except tk.TclError:
            print("app_icon.png not found, skipping icon.")

        self.current_theme = "light"

        self.setup_styles()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, expand=True, fill="both")

        self.create_sections_tab()
        self.create_actions_tab()
        self.create_people_tab()
        self.create_courses_tab()
        self.create_terms_accounts_tab()
        self.create_about_tab()

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def get_theme_colors(self):
        light_theme = {
            "accent": "#D94F4F", "secondary": "#F2F2F2", "contrast": "#222222",
            "highlight": "#4D9FE0", "surface": "#FFFFFF", "tree_heading": "#E0E0E0",
            "dialog_bg": "#F2F2F2", "button_bg": "#E1E1E1"
        }
        dark_theme = {
            "accent": "#D94F4F", "secondary": "#212121", "contrast": "#F5F5F5",
            "highlight": "#4D9FE0", "surface": "#2C2C2C", "tree_heading": "#383838",
            "dialog_bg": "#333333", "button_bg": "#4F4F4F"
        }
        return dark_theme if self.current_theme == "dark" else light_theme

    def toggle_theme(self):
        self.current_theme = "dark" if self.current_theme == "light" else "light"
        self.setup_styles()

    def setup_styles(self):
        colors = self.get_theme_colors()
        accent = colors["accent"]
        secondary = colors["secondary"]
        contrast = colors["contrast"]
        highlight = colors["highlight"]
        surface = colors["surface"]
        button_bg = colors["button_bg"]

        style = ttk.Style(self)
        style.theme_use("clam")

        self.configure(background=secondary)
        style.configure("Accent.TButton", background=accent, foreground="#FFFFFF", font=('TkDefaultFont', 10, 'bold'),
                        borderwidth=1, relief="raised")
        style.map("Accent.TButton", background=[('active', '#B84444')])

        style.configure("TButton", background=button_bg, foreground=contrast, borderwidth=1, relief="raised")
        style.map("TButton", background=[('active', highlight)])

        style.configure("TNotebook", background=secondary, borderwidth=0)
        style.configure("TNotebook.Tab", background=secondary, foreground=contrast, lightcolor=secondary,
                        padding=[10, 5])
        style.map("TNotebook.Tab", background=[("selected", surface)], foreground=[("selected", accent)])
        style.configure("TFrame", background=secondary)
        style.configure("TLabelFrame", background=secondary, borderwidth=1, relief="groove")
        style.configure("TLabelFrame.Label", background=secondary, foreground=contrast,
                        font=('TkDefaultFont', 10, 'bold'))
        style.configure("TLabel", background=secondary, foreground=contrast)
        style.configure("link.TLabel", foreground=highlight, background=secondary,
                        font=('TkDefaultFont', 10, 'underline'))
        style.configure("Treeview", background=surface, foreground=contrast, fieldbackground=surface, rowheight=25)
        style.map("Treeview", background=[('selected', highlight)], foreground=[('selected', "#FFFFFF")])
        style.configure("Treeview.Heading", background=colors["tree_heading"], foreground=contrast,
                        font=('TkDefaultFont', 10, 'bold'))

        style.map('TCombobox', fieldbackground=[('readonly', surface)],
                  selectbackground=[('readonly', surface)],
                  selectforeground=[('readonly', contrast)],
                  foreground=[('readonly', contrast)])

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to save your data and quit?"):
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
        ttk.Button(btn_frame, text="Import from CSV", command=lambda: self.import_from_csv(item_name)).pack(side="left",
                                                                                                            padx=5)

        return tree

    def create_people_tab(self):
        self.people_tree = self.create_management_tab(self.notebook, "People", ('user_id', 'name'), "Person",
                                                      lambda tree: self.edit_item("Person", tree))
        self.refresh_people_list()

    def refresh_people_list(self):
        for item in self.people_tree.get_children():
            self.people_tree.delete(item)
        for user_id, person in sorted(self.data_manager.people.items()):
            self.people_tree.insert("", "end", values=(user_id, person.name))

    def create_courses_tab(self):
        self.courses_tree = self.create_management_tab(self.notebook, "Courses",
                                                       ('course_id_portion', 'short_name', 'long_name'), "Course",
                                                       lambda tree: self.edit_item("Course", tree))
        self.refresh_courses_list()

    def refresh_courses_list(self):
        for item in self.courses_tree.get_children():
            self.courses_tree.delete(item)
        for cid, course in sorted(self.data_manager.courses.items()):
            self.courses_tree.insert("", "end", values=(cid, course.short_name, course.long_name))

    def create_terms_accounts_tab(self):
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Terms & Accounts")
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        terms_frame = ttk.LabelFrame(frame, text="Terms", padding="10")
        terms_frame.grid(row=0, column=0, sticky="nsew", padx=5)
        self.terms_tree = self.create_sub_management_view(terms_frame,
                                                          ('name', 'term_id', 'short_code', 'start_date', 'end_date'),
                                                          "Term")
        self.terms_tree.bind("<Double-1>", lambda event: self.edit_item("Term", self.terms_tree))

        accounts_frame = ttk.LabelFrame(frame, text="Accounts", padding="10")
        accounts_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        self.accounts_tree = self.create_sub_management_view(accounts_frame, ('account_id',), "Account")
        self.accounts_tree.bind("<Double-1>", lambda event: self.edit_item("Account", self.accounts_tree))

        self.refresh_terms_list()
        self.refresh_accounts_list()

    def create_sub_management_view(self, parent, columns, item_name):
        tree = ttk.Treeview(parent, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col.replace('_', ' ').title())
        tree.pack(expand=True, fill="both", pady=5)

        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x")

        ttk.Button(btn_frame, text="Add", command=lambda: self.add_item(item_name, tree)).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Edit", command=lambda: self.edit_item(item_name, tree)).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Delete", command=lambda: self.delete_item(item_name, tree)).pack(side="left",
                                                                                                     padx=5)
        ttk.Button(btn_frame, text="Import from CSV", command=lambda: self.import_from_csv(item_name)).pack(side="left",
                                                                                                            padx=5)

        return tree

    def refresh_terms_list(self):
        for item in self.terms_tree.get_children():
            self.terms_tree.delete(item)
        for tname, term in sorted(self.data_manager.terms.items()):
            self.terms_tree.insert("", "end",
                                   values=(tname, term.term_id, term.short_code, term.start_date, term.end_date))

    def refresh_accounts_list(self):
        for item in self.accounts_tree.get_children():
            self.accounts_tree.delete(item)
        for aid, acc in sorted(self.data_manager.accounts.items()):
            self.accounts_tree.insert("", "end", values=(aid,))

    def create_sections_tab(self):
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Sections & Enrollments")

        self.sections_tree = ttk.Treeview(frame,
                                          columns=('course', 'term', 'section', 'status', 'start_date', 'end_date'),
                                          show="headings")
        self.sections_tree.heading('course', text='Course')
        self.sections_tree.heading('term', text='Term')
        self.sections_tree.heading('section', text='Section #')
        self.sections_tree.heading('status', text='Status')
        self.sections_tree.heading('start_date', text='Start Date')
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
        for item in self.sections_tree.get_children():
            self.sections_tree.delete(item)
        for i, sec in enumerate(self.data_manager.sections):
            course_name = self.data_manager.courses.get(sec.course_id_portion, Course('N/A', 'N/A', 'N/A')).short_name
            term_obj = self.data_manager.terms.get(sec.term_name, Term('N/A', 'N/A', 'N/A', '', ''))
            self.sections_tree.insert("", "end", iid=i, values=(
                f"{course_name} ({sec.course_id_portion})",
                f"{term_obj.name} ({term_obj.term_id})",
                sec.section_number,
                sec.status,
                sec.start_date,
                sec.end_date
            ))

    def create_section(self):
        dialog = SectionDialog(self, "Create Section", self.data_manager, self.get_theme_colors())
        if dialog.result:
            new_section = Section(**dialog.result)
            self.data_manager.sections.append(new_section)
            self.refresh_sections_list()

    def edit_section(self, event=None):
        selected = self.sections_tree.selection()
        if not selected:
            if event:
                return
            messagebox.showwarning("Selection Error", "Please select a section to edit.")
            return

        section_index = int(selected[0])
        section_obj = self.data_manager.sections[section_index]

        dialog = SectionDialog(self, "Edit Section", self.data_manager, self.get_theme_colors(),
                               initial_data=section_obj.to_dict())
        if dialog.result:
            for key, value in dialog.result.items():
                setattr(section_obj, key, value)
            self.refresh_sections_list()

    def manage_enrollments(self):
        selected = self.sections_tree.selection()
        if not selected:
            messagebox.showwarning("Selection Error", "Please select a section to manage.")
            return

        section_index = int(selected[0])
        section = self.data_manager.sections[section_index]
        EnrollmentDialog(self, f"Enrollments for {section.course_id_portion}", section, self.data_manager,
                         self.get_theme_colors())

    def delete_section(self):
        selected = self.sections_tree.selection()
        if not selected:
            messagebox.showwarning("Selection Error", "Please select a section to delete.")
            return
        if messagebox.askyesno("Confirm Delete",
                               "Are you sure you want to delete this section and all its enrollments?"):
            indices_to_delete = sorted([int(s) for s in selected], reverse=True)
            for index in indices_to_delete:
                del self.data_manager.sections[index]
            self.refresh_sections_list()

    def create_actions_tab(self):
        frame = ttk.Frame(self.notebook, padding="50")
        self.notebook.add(frame, text="Generate Files")

        ttk.Button(frame, text="Save All Data", style="Accent.TButton", command=self.save_all_data).pack(pady=20,
                                                                                                         ipady=10,
                                                                                                         fill='x')
        ttk.Button(frame, text="Generate CSV Files", style="Accent.TButton", command=self.generate_csv).pack(pady=20,
                                                                                                             ipady=10,
                                                                                                             fill='x')

    def save_all_data(self):
        if self.data_manager.save_data():
            messagebox.showinfo("Success", f"Data saved successfully to {DATA_FILE}")
        else:
            messagebox.showerror("Error", "Could not save data.")

    def generate_csv(self):
        prefix = simpledialog.askstring("File Prefix", "Enter an optional prefix for the CSV files:", parent=self)
        if prefix is None:
            return

        directory = filedialog.askdirectory(title="Select Folder to Save CSV Files")
        if not directory:
            return

        result = self.data_manager.generate_csv_files(directory, prefix)
        messagebox.showinfo("CSV Generation", result)

    def add_item(self, item_name, tree):
        fields_map = {
            "Person": [('user_id', 'User ID'), ('name', 'Name')],
            "Course": [('course_id_portion', 'Course ID Portion'), ('short_name', 'Short Name'),
                       ('long_name', 'Long Name')],
            "Term": [('name', 'Name (Unique)'), ('term_id', 'Term ID'), ('short_code', 'Short Code'),
                     ('start_date', 'Start Date (YYYY-MM-DD)'), ('end_date', 'End Date (YYYY-MM-DD)')],
            "Account": [('account_id', 'Account ID')]
        }
        key_field = 'name' if item_name == "Term" else list(fields_map[item_name][0])[0]

        dialog = ManagementDialog(self, f"Add {item_name}", fields_map[item_name], theme_colors=self.get_theme_colors())
        if not dialog.result:
            return

        key = dialog.result[key_field]
        if not key:
            messagebox.showerror("Input Error", f"{key_field.replace('_', ' ').title()} cannot be empty.")
            return

        data_map = {"Person": self.data_manager.people, "Course": self.data_manager.courses,
                    "Term": self.data_manager.terms, "Account": self.data_manager.accounts}
        if key in data_map[item_name]:
            messagebox.showerror("Error", f"A {item_name.lower()} with this ID/Name already exists.")
            return

        constructors = {"Person": Person, "Course": Course, "Term": Term, "Account": Account}
        data_map[item_name][key] = constructors[item_name](**dialog.result)

        refresh_map = {"Person": self.refresh_people_list, "Course": self.refresh_courses_list,
                       "Term": self.refresh_terms_list, "Account": self.refresh_accounts_list}
        refresh_map[item_name]()

    def edit_item(self, item_name, tree):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Selection Error", f"Please select a {item_name.lower()} to edit.")
            return

        fields_map = {
            "Person": [('user_id', 'User ID'), ('name', 'Name')],
            "Course": [('course_id_portion', 'Course ID Portion'), ('short_name', 'Short Name'),
                       ('long_name', 'Long Name')],
            "Term": [('name', 'Name (Unique)'), ('term_id', 'Term ID'), ('short_code', 'Short Code'),
                     ('start_date', 'Start Date'), ('end_date', 'End Date')],
            "Account": [('account_id', 'Account ID')]
        }
        key_field = 'name' if item_name == "Term" else list(fields_map[item_name][0])[0]
        key = tree.item(selected[0])['values'][0]

        data_map = {"Person": self.data_manager.people, "Course": self.data_manager.courses,
                    "Term": self.data_manager.terms, "Account": self.data_manager.accounts}
        item_obj = data_map[item_name][key]

        dialog = ManagementDialog(self, f"Edit {item_name}", fields_map[item_name], initial_data=item_obj.to_dict(),
                                  readonly_key=key_field, theme_colors=self.get_theme_colors())
        if not dialog.result:
            return

        if item_name == "Term":
            new_name = dialog.result['name']
            old_name = key

            if new_name != old_name:
                if new_name in self.data_manager.terms:
                    messagebox.showerror("Error", "A term with this name already exists.")
                    return
                for section in self.data_manager.sections:
                    if section.term_name == old_name:
                        section.term_name = new_name
                del self.data_manager.terms[old_name]

            self.data_manager.terms[new_name] = Term(**dialog.result)
        else:
            constructors = {"Person": Person, "Course": Course, "Account": Account}
            data_map[item_name][key] = constructors[item_name](**dialog.result)

        refresh_map = {"Person": self.refresh_people_list, "Course": self.refresh_courses_list,
                       "Term": self.refresh_terms_list, "Account": self.refresh_accounts_list}
        refresh_map[item_name]()
        if item_name == "Term":
            self.refresh_sections_list()

    def delete_item(self, item_name, tree):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Selection Error", f"Please select a {item_name.lower()} to delete.")
            return

        key = tree.item(selected[0])['values'][0]

        in_use_msg = f"Cannot delete this {item_name.lower()}. It is in use by one or more sections."
        if item_name == "Person" and any(e.user_id == key for s in self.data_manager.sections for e in s.enrollments):
            messagebox.showerror("Deletion Error", in_use_msg)
            return

        if item_name in ["Course", "Term", "Account"]:
            check_key = {"Course": "course_id_portion", "Term": "term_name", "Account": "account_id"}[item_name]
            if any(getattr(s, check_key, None) == key for s in self.data_manager.sections):
                messagebox.showerror("Deletion Error", in_use_msg)
                return

        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete this {item_name.lower()}?"):
            data_map = {"Person": self.data_manager.people, "Course": self.data_manager.courses,
                        "Term": self.data_manager.terms, "Account": self.data_manager.accounts}
            if key in data_map[item_name]:
                del data_map[item_name][key]

            refresh_map = {"Person": self.refresh_people_list, "Course": self.refresh_courses_list,
                           "Term": self.refresh_terms_list, "Account": self.refresh_accounts_list}
            refresh_map[item_name]()

    def import_from_csv(self, item_name):
        file_path = filedialog.askopenfilename(
            title=f"Import {item_name}s from CSV",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        importer_map = {
            "Person": self.data_manager.import_people_from_csv,
            "Course": self.data_manager.import_courses_from_csv,
            "Term": self.data_manager.import_terms_from_csv,
            "Account": self.data_manager.import_accounts_from_csv
        }

        result = importer_map[item_name](file_path)

        if result.get('error'):
            messagebox.showerror("Import Error", result['error'])
        else:
            messagebox.showinfo("Import Complete",
                                f"Added: {result['added']}\nSkipped (duplicates): {result['skipped']}")

        refresh_map = {
            "Person": self.refresh_people_list,
            "Course": self.refresh_courses_list,
            "Term": self.refresh_terms_list,
            "Account": self.refresh_accounts_list
        }
        refresh_map[item_name]()

    def create_about_tab(self):
        frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(frame, text="About")

        content_frame = ttk.Frame(frame)
        content_frame.pack(anchor="center", expand=True)

        # Row 0
        ttk.Label(content_frame, text="App Name:", font=('TkDefaultFont', 10, 'bold')).grid(row=0, column=0,
                                                                                            sticky="ne", padx=5, pady=5)
        ttk.Label(content_frame, text="Canvas SIS Prep Tool").grid(row=0, column=1, sticky="nw", padx=5, pady=5)

        # Row 1
        ttk.Label(content_frame, text="Author:", font=('TkDefaultFont', 10, 'bold')).grid(row=1, column=0, sticky="ne",
                                                                                          padx=5, pady=5)
        ttk.Label(content_frame, text="Harrison Smith").grid(row=1, column=1, sticky="nw", padx=5, pady=5)

        # Row 2
        ttk.Label(content_frame, text="AI Assistant:", font=('TkDefaultFont', 10, 'bold')).grid(row=2, column=0,
                                                                                                sticky="ne", padx=5,
                                                                                                pady=5)
        ttk.Label(content_frame, text="Gemini 2.5 Pro").grid(row=2, column=1, sticky="nw", padx=5, pady=5)

        # Social links beside Author and AI Assistant
        links_frame = ttk.Frame(content_frame)
        links_frame.grid(row=1, column=2, rowspan=2, sticky="nsw", padx=20)

        github_link = ttk.Label(links_frame, text="GitHub", style="link.TLabel", cursor="hand2")
        github_link.pack(side="left", padx=5)
        github_link.bind("<Button-1>", lambda e: self.open_link("https://github.com/hsmith-dev"))

        linkedin_link = ttk.Label(links_frame, text="LinkedIn", style="link.TLabel", cursor="hand2")
        linkedin_link.pack(side="left", padx=5)
        linkedin_link.bind("<Button-1>", lambda e: self.open_link("https://linkedin.com/in/hsmith-dev"))

        email_link = ttk.Label(links_frame, text="Email", style="link.TLabel", cursor="hand2")
        email_link.pack(side="left", padx=5)
        email_link.bind("<Button-1>", lambda e: self.open_link("mailto:harrison@hsmith.dev"))

        # Row 3
        ttk.Label(content_frame, text="Last Update Date:", font=('TkDefaultFont', 10, 'bold')).grid(row=3, column=0,
                                                                                                    sticky="ne", padx=5,
                                                                                                    pady=5)
        ttk.Label(content_frame, text="October 18, 2025").grid(row=3, column=1, sticky="nw", padx=5, pady=5)

        # Row 4
        ttk.Label(content_frame, text="Description:", font=('TkDefaultFont', 10, 'bold')).grid(row=4, column=0,
                                                                                               sticky="ne", padx=5,
                                                                                               pady=5)
        description_text = "A desktop application designed to streamline the creation of CSV files for Canvas SIS imports. It allows for the management of people, courses, terms, and accounts, and facilitates the creation of course sections and enrollments. The tool helps reduce manual data entry errors by providing a structured interface for data management and generating Canvas-compliant `courses.csv`, `sections.csv`, and `enrollments.csv` files."
        ttk.Label(content_frame, text=description_text, wraplength=500, justify="left").grid(row=4, column=1,
                                                                                             columnspan=2, sticky="nw",
                                                                                             padx=5, pady=5)

        # Row 5
        ttk.Button(content_frame, text="Toggle Theme", command=self.toggle_theme).grid(row=5, column=0, columnspan=3,
                                                                                       pady=20)

    def open_link(self, url):
        webbrowser.open_new(url)


# --- Custom Dialogs ---
class ManagementDialog(simpledialog.Dialog):
    def __init__(self, parent, title, fields, theme_colors, initial_data=None, readonly_key=None):
        self.fields = fields
        self.theme_colors = theme_colors
        self.initial_data = initial_data or {}
        self.readonly_key = readonly_key
        self.entries = {}
        super().__init__(parent, title)

    def body(self, master):
        master.config(bg=self.theme_colors['dialog_bg'])
        for i, (key, label) in enumerate(self.fields):
            lbl = ttk.Label(master, text=f"{label}:", background=self.theme_colors['dialog_bg'],
                            foreground=self.theme_colors['contrast'])
            lbl.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            entry = ttk.Entry(master, width=30)
            entry.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
            entry.insert(0, self.initial_data.get(key, ""))
            if key == self.readonly_key:
                entry.config(state="readonly")
            self.entries[key] = entry
        return self.entries[self.fields[0][0]]

    def buttonbox(self):
        box = ttk.Frame(self)
        box.config(bg=self.theme_colors['dialog_bg'])
        ttk.Button(box, text="OK", width=10, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(box, text="Cancel", width=10, command=self.cancel).pack(side=tk.LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
        box.pack()

    def apply(self):
        self.result = {key: entry.get() for key, entry in self.entries.items()}


class SectionDialog(simpledialog.Dialog):
    def __init__(self, parent, title, data_manager, theme_colors, initial_data=None):
        self.data_manager = data_manager
        self.theme_colors = theme_colors
        self.initial_data = initial_data or {}
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        master.config(bg=self.theme_colors['dialog_bg'])

        ttk.Label(master, text="Course:", background=self.theme_colors['dialog_bg'],
                  foreground=self.theme_colors['contrast']).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.course_var = tk.StringVar()
        self.course_combo = ttk.Combobox(master, state="readonly", textvariable=self.course_var,
                                         values=list(self.data_manager.courses.keys()))
        self.course_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        self.course_combo.set(self.initial_data.get('course_id_portion', ''))

        ttk.Label(master, text="Term:", background=self.theme_colors['dialog_bg'],
                  foreground=self.theme_colors['contrast']).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.term_var = tk.StringVar()
        term_display = [f"{t.name} ({t.term_id})" for tname, t in self.data_manager.terms.items()]
        self.term_combo = ttk.Combobox(master, state="readonly", textvariable=self.term_var, values=term_display)
        self.term_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        if self.initial_data.get('term_name'):
            term = self.data_manager.terms.get(self.initial_data['term_name'])
            if term:
                self.term_combo.set(f"{term.name} ({term.term_id})")

        ttk.Label(master, text="Account:", background=self.theme_colors['dialog_bg'],
                  foreground=self.theme_colors['contrast']).grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.account_var = tk.StringVar()
        self.account_combo = ttk.Combobox(master, state="readonly", textvariable=self.account_var,
                                          values=list(self.data_manager.accounts.keys()))
        self.account_combo.grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        self.account_combo.set(self.initial_data.get('account_id', ''))

        ttk.Label(master, text="Section Number:", background=self.theme_colors['dialog_bg'],
                  foreground=self.theme_colors['contrast']).grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.section_num_entry = ttk.Entry(master)
        self.section_num_entry.grid(row=3, column=1, sticky="ew", padx=5, pady=2)
        self.section_num_entry.insert(0, self.initial_data.get('section_number', ''))

        ttk.Label(master, text="Status:", background=self.theme_colors['dialog_bg'],
                  foreground=self.theme_colors['contrast']).grid(row=4, column=0, sticky="w", padx=5, pady=2)
        self.status_var = tk.StringVar()
        self.status_combo = ttk.Combobox(master, state="readonly", textvariable=self.status_var,
                                         values=['active', 'deleted', 'completed', 'published'])
        self.status_combo.grid(row=4, column=1, sticky="ew", padx=5, pady=2)
        self.status_combo.set(self.initial_data.get('status', 'active'))

        ttk.Label(master, text="Start Date (YYYY-MM-DD):", background=self.theme_colors['dialog_bg'],
                  foreground=self.theme_colors['contrast']).grid(row=5, column=0, sticky="w", padx=5, pady=2)
        self.start_date_entry = ttk.Entry(master)
        self.start_date_entry.grid(row=5, column=1, sticky="ew", padx=5, pady=2)
        self.start_date_entry.insert(0, self.initial_data.get('start_date', ''))

        ttk.Label(master, text="End Date (YYYY-MM-DD):", background=self.theme_colors['dialog_bg'],
                  foreground=self.theme_colors['contrast']).grid(row=6, column=0, sticky="w", padx=5, pady=2)
        self.end_date_entry = ttk.Entry(master)
        self.end_date_entry.grid(row=6, column=1, sticky="ew", padx=5, pady=2)
        self.end_date_entry.insert(0, self.initial_data.get('end_date', ''))

        return self.course_combo

    def buttonbox(self):
        box = ttk.Frame(self)
        box.config(bg=self.theme_colors['dialog_bg'])
        ttk.Button(box, text="OK", width=10, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(box, text="Cancel", width=10, command=self.cancel).pack(side=tk.LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
        box.pack()

    def apply(self):
        term_str = self.term_var.get()
        term_name = term_str.rsplit(' (', 1)[0] if ' (' in term_str else ''

        self.result = {
            'course_id_portion': self.course_var.get(),
            'term_name': term_name,
            'account_id': self.account_var.get(),
            'section_number': self.section_num_entry.get(),
            'status': self.status_var.get(),
            'start_date': self.start_date_entry.get(),
            'end_date': self.end_date_entry.get()
        }


class EnrollmentDialog(simpledialog.Dialog):
    def __init__(self, parent, title, section, data_manager, theme_colors):
        self.section = section
        self.data_manager = data_manager
        self.theme_colors = theme_colors
        super().__init__(parent, title)

    def body(self, master):
        master.config(bg=self.theme_colors['dialog_bg'])
        self.tree = ttk.Treeview(master, columns=('user_id', 'name', 'role', 'status'), show="headings")
        self.tree.heading('user_id', text='User ID')
        self.tree.heading('name', text='Name')
        self.tree.heading('role', text='Role')
        self.tree.heading('status', text='Status')
        self.tree.grid(row=0, column=0, columnspan=2, sticky="nsew")
        self.refresh_enrollments()

        add_frame = ttk.LabelFrame(master, text="Add New Enrollment", padding=10)
        add_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=10)

        ttk.Label(add_frame, text="Person:", background=self.theme_colors['secondary'],
                  foreground=self.theme_colors['contrast']).grid(row=0, column=0, sticky="w")
        self.person_var = tk.StringVar()
        person_values = [f"{p.name} ({uid})" for uid, p in self.data_manager.people.items()]
        self.person_combo = ttk.Combobox(add_frame, state="readonly", textvariable=self.person_var,
                                         values=person_values)
        self.person_combo.grid(row=0, column=1, sticky="ew")

        ttk.Label(add_frame, text="Role:", background=self.theme_colors['secondary'],
                  foreground=self.theme_colors['contrast']).grid(row=1, column=0, sticky="w")
        self.role_var = tk.StringVar()
        self.role_combo = ttk.Combobox(add_frame, state="readonly", textvariable=self.role_var,
                                       values=["Student", "Teaching Assistant", "Instructor", "Program Manager"])
        self.role_combo.grid(row=1, column=1, sticky="ew")

        ttk.Label(add_frame, text="Status:", background=self.theme_colors['secondary'],
                  foreground=self.theme_colors['contrast']).grid(row=2, column=0, sticky="w")
        self.status_var = tk.StringVar()
        self.status_combo = ttk.Combobox(add_frame, state="readonly", textvariable=self.status_var,
                                         values=['active', 'completed', 'inactive', 'deleted'])
        self.status_combo.set('active')
        self.status_combo.grid(row=2, column=1, sticky="ew")

        ttk.Button(add_frame, text="Add", command=self.add_enrollment).grid(row=3, column=1, sticky="e", pady=5)
        ttk.Button(master, text="Delete Selected", command=self.delete_enrollment).grid(row=2, column=0)

    def refresh_enrollments(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, enroll in enumerate(self.section.enrollments):
            person = self.data_manager.people.get(enroll.user_id, Person("Unknown", enroll.user_id))
            self.tree.insert("", "end", iid=i, values=(enroll.user_id, person.name, enroll.role, enroll.status))

    def add_enrollment(self):
        person_str = self.person_var.get()
        if not person_str:
            messagebox.showwarning("Input Error", "Please select a person.")
            return
        user_id = person_str[person_str.rfind('(') + 1:-1]
        role = self.role_var.get()
        status = self.status_var.get()

        if not user_id or not role:
            messagebox.showwarning("Input Error", "Please select a person and a role.")
            return

        enrollment = Enrollment(user_id, role, status)
        self.section.add_enrollment(enrollment)
        self.refresh_enrollments()

    def delete_enrollment(self):
        selected = self.tree.selection()
        if not selected:
            return
        if messagebox.askyesno("Confirm Delete", "Delete selected enrollment(s)?"):
            indices_to_delete = sorted([int(s) for s in selected], reverse=True)
            for index in indices_to_delete:
                del self.section.enrollments[index]
            self.refresh_enrollments()

    def buttonbox(self):
        box = ttk.Frame(self)
        box.config(bg=self.theme_colors['dialog_bg'])
        ttk.Button(box, text="Close", width=10, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)


# --- Main Execution ---
if __name__ == "__main__":
    dm = DataManager()
    app = App(dm)
    app.mainloop()

