# kindergarten_manager_v5.4_fixed_layout.py
# سامانه مدیریت مهدکودک - نسخه با رفع مشکل چیدمان دکمه‌ها

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import jdatetime
import pandas as pd
import atexit
import re

class KindergartenApp(tk.Tk):
    """
    کلاس اصلی اپلیکیشن مدیریت مهدکودک با بهبودهای چیدمان و تغییر اندازه.
    """
    # --- تعریف پالت رنگی (تم گرم و دوستانه بر پایه سبزآبی) ---
    BG_COLOR = "#F4F6F7"
    FRAME_COLOR = "#FFFFFF"
    PRIMARY_COLOR = "#1ABC9C"
    ACCENT_COLOR = "#16A085"
    NAV_BG_COLOR = "#E8F6F3"
    NAV_SELECTED_BG = "#FFFFFF"
    TEXT_COLOR = "#34495E"
    MUTED_TEXT_COLOR = "#7F8C8D"
    HEADER_TEXT_COLOR = "#FFFFFF"
    SUCCESS_COLOR = "#2ECC71"
    ERROR_COLOR = "#E74C3C"
    ROW_ALT_COLOR = "#F8F9F9"

    # --- تعریف فونت‌های مدرن ---
    FONT_NAME = "Vazirmatn"
    HEADER_FONT = (FONT_NAME, 15, 'bold')
    BODY_FONT = (FONT_NAME, 11)
    BOLD_FONT = (FONT_NAME, 11, 'bold')
    TAB_FONT = (FONT_NAME, 12, 'bold')

    def __init__(self):
        super().__init__()
        self.title("سامانه مدیریت مهدکودک")
        self.geometry("1400x950")

        # --- بهبود ۱: تعیین حداقل اندازه برای پنجره ---
        self.minsize(1100, 750)

        self.configure(bg=self.BG_COLOR)

        # --- اعمال استایل‌ها ---
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure('.', background=self.BG_COLOR, foreground=self.TEXT_COLOR, font=self.BODY_FONT)
        self.style.configure('TFrame', background=self.FRAME_COLOR)
        self.style.configure('Nav.TFrame', background=self.NAV_BG_COLOR)
        self.style.configure('TLabelframe', font=self.BOLD_FONT, background=self.FRAME_COLOR, borderwidth=1, relief="solid", bordercolor="#DEE2E6")
        self.style.configure('TLabelframe.Label', font=self.HEADER_FONT, background=self.FRAME_COLOR, foreground=self.PRIMARY_COLOR, padding=(0, 0, 0, 10))
        self.style.configure('TLabel', font=self.BOLD_FONT, background=self.FRAME_COLOR, foreground=self.TEXT_COLOR)
        self.style.configure('Header.TLabel', font=self.HEADER_FONT, background=self.FRAME_COLOR, foreground=self.PRIMARY_COLOR)
        self.style.configure('TButton', font=self.BOLD_FONT, padding=12, background=self.PRIMARY_COLOR, foreground=self.HEADER_TEXT_COLOR, borderwidth=0, relief='flat', focusthickness=0)
        self.style.map('TButton', background=[('active', self.ACCENT_COLOR), ('!disabled', self.PRIMARY_COLOR)])
        self.style.configure('Nav.TButton', font=self.TAB_FONT, padding=(20, 15), background=self.NAV_BG_COLOR, foreground=self.MUTED_TEXT_COLOR, borderwidth=0, relief='flat', anchor='e')
        self.style.map('Nav.TButton', background=[('active', '#D4E6E1'), ('selected', self.NAV_SELECTED_BG)], foreground=[('active', self.TEXT_COLOR), ('selected', self.PRIMARY_COLOR)])
        self.style.configure('TEntry', font=self.BODY_FONT, padding=8, relief='flat')
        self.style.configure('TCombobox', font=self.BODY_FONT, padding=8, relief='flat')
        self.style.map('TEntry', fieldbackground=[('focus', self.FRAME_COLOR)], relief=[('focus', 'solid')], bordercolor=[('focus', self.PRIMARY_COLOR)])
        self.style.map('TCombobox', fieldbackground=[('focus', self.FRAME_COLOR)], relief=[('focus', 'solid')], bordercolor=[('focus', self.PRIMARY_COLOR)])
        self.style.configure('Treeview', font=self.BODY_FONT, rowheight=32, background=self.FRAME_COLOR, fieldbackground=self.FRAME_COLOR, relief='flat', borderwidth=0)
        self.style.configure('Treeview.Heading', font=self.BOLD_FONT, padding=12, background=self.PRIMARY_COLOR, foreground=self.HEADER_TEXT_COLOR, relief="flat")
        self.style.map('Treeview.Heading', background=[('active', self.ACCENT_COLOR)])
        self.style.configure("oddrow.Treeview", background=self.ROW_ALT_COLOR)
        self.style.configure("evenrow.Treeview", background=self.FRAME_COLOR)

        # --- اتصال به دیتابیس ---
        self.db_path = "kindergarten.db"
        self.conn = None
        self.cursor = None
        self.student_name_to_id = {}
        self.connect_to_database()

        atexit.register(self.close_database_connection)

        # --- ساختار اصلی برنامه با نویگیشن عمودی ---
        main_frame = ttk.Frame(self, style='TFrame')
        main_frame.pack(expand=True, fill="both")

        nav_frame = ttk.Frame(main_frame, width=220, style='Nav.TFrame')
        nav_frame.pack(side="right", fill="y")
        nav_frame.pack_propagate(False)

        self.content_frame = ttk.Frame(main_frame, style='TFrame')
        self.content_frame.pack(side="right", expand=True, fill="both")
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        # --- ایجاد فریم‌ها ---
        self.tabs = {}
        page_creators = {
            "دانش‌آموزان": self.create_student_tab,
            "پرسنل": self.create_staff_tab,
            "حضور و غیاب": self.create_attendance_tab,
            "امور مالی": self.create_finance_tab,
            "ارتباطات": self.create_communication_tab,
            "گزارشات": self.create_report_tab
        }

        # --- ایجاد دکمه‌های نویگیشن ---
        title_label = ttk.Label(nav_frame, text="مدیریت مهدکودک", font=self.HEADER_FONT, background=self.NAV_BG_COLOR, foreground=self.PRIMARY_COLOR, padding=20)
        title_label.pack(pady=(10, 20))

        self.nav_buttons = {}
        for name, creator in page_creators.items():
            frame = creator()
            self.tabs[name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

            btn = ttk.Button(nav_frame, text=name, command=lambda n=name: self.show_frame(n), style='Nav.TButton')
            btn.pack(fill="x", pady=2, padx=10)
            self.nav_buttons[name] = btn

        # --- بارگذاری اولیه داده‌ها پس از ساخت کامل UI ---
        self.update_all_student_comboboxes()
        self.load_all_data()
        self.show_frame("دانش‌آموزان")

    def show_frame(self, page_name):
        frame = self.tabs[page_name]
        frame.tkraise()
        for name, button in self.nav_buttons.items():
            if name == page_name:
                button.state(['selected'])
            else:
                button.state(['!selected'])

    def connect_to_database(self):
        try:
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()
            self.create_tables()
        except sqlite3.Error as e:
            messagebox.showerror("خطای دیتابیس", f"خطا در اتصال به پایگاه داده: {e}")
            self.destroy()

    def create_tables(self):
        self.cursor.execute("CREATE TABLE IF NOT EXISTS students (id INTEGER PRIMARY KEY, name TEXT NOT NULL, family_name TEXT NOT NULL, national_id TEXT UNIQUE, birth_date TEXT, entry_date TEXT NOT NULL, father_name TEXT, mother_name TEXT, parent_phone TEXT, address TEXT, class_name TEXT, status TEXT NOT NULL, allergies TEXT, notes TEXT)")
        self.cursor.execute("CREATE TABLE IF NOT EXISTS staff (id INTEGER PRIMARY KEY, name TEXT NOT NULL, family_name TEXT NOT NULL, national_id TEXT UNIQUE, birth_date TEXT, hire_date TEXT NOT NULL, phone TEXT, address TEXT, role TEXT NOT NULL, salary REAL, status TEXT NOT NULL, notes TEXT)")
        self.cursor.execute("CREATE TABLE IF NOT EXISTS attendance (id INTEGER PRIMARY KEY, person_id INTEGER NOT NULL, person_type TEXT NOT NULL, date TEXT NOT NULL, entry_time TEXT, exit_time TEXT, status TEXT, UNIQUE(person_id, person_type, date))")
        self.cursor.execute("CREATE TABLE IF NOT EXISTS student_finance (id INTEGER PRIMARY KEY, student_id INTEGER NOT NULL, payment_date TEXT NOT NULL, payment_amount REAL NOT NULL, payment_type TEXT, description TEXT, FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE)")
        self.cursor.execute("CREATE TABLE IF NOT EXISTS student_communications (id INTEGER PRIMARY KEY, student_id INTEGER NOT NULL, date TEXT NOT NULL, subject TEXT, message TEXT, FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE)")
        self.cursor.execute("CREATE TABLE IF NOT EXISTS kindergarten_expenses (id INTEGER PRIMARY KEY, date TEXT NOT NULL, description TEXT, amount REAL NOT NULL, category TEXT)")
        self.conn.commit()

    def close_database_connection(self):
        if self.conn:
            self.conn.close()

    def load_all_data(self):
        self.load_students()
        self.load_staff()
        self.load_finances()
        self.load_communications()
        self.load_expenses()
        self.load_attendance_list()
        self.generate_financial_report()

    def _add_items_to_treeview(self, tree, items, has_iid=False):
        for i, item in enumerate(items):
            tag = 'oddrow' if i % 2 != 0 else 'evenrow'
            if has_iid:
                iid, values = item
                tree.insert("", "end", iid=iid, values=values, tags=(tag,))
            else:
                tree.insert("", "end", values=item, tags=(tag,))

    # --- تب دانش‌آموزان (Students) ---
    def create_student_tab(self):
        student_tab = ttk.Frame(self.content_frame, style='TFrame', padding=25)

        # --- بهبود ۲: استفاده از PanedWindow برای تقسیم‌بندی قابل تنظیم ---
        main_pane = ttk.PanedWindow(student_tab, orient=tk.HORIZONTAL)
        main_pane.pack(fill="both", expand=True)

        # --- بخش فرم در یک طرف PanedWindow ---
        form_frame = ttk.LabelFrame(main_pane, text="فرم ثبت/ویرایش دانش‌آموز", padding=25)

        self.student_labels = {
            "name": "نام:", "family_name": "نام خانوادگی:", "national_id": "کد ملی:",
            "birth_date": "تاریخ تولد (13yy/mm/dd):", "entry_date": "تاریخ ثبت‌نام (14yy/mm/dd):",
            "father_name": "نام پدر:", "mother_name": "نام مادر:", "parent_phone": "شماره تماس والدین:",
            "address": "آدرس:", "class_name": "کلاس:", "status": "وضعیت:",
            "allergies": "آلرژی‌ها:", "notes": "توضیحات:"
        }
        self.student_entries = {}
        for i, (key, text) in enumerate(self.student_labels.items()):
            ttk.Label(form_frame, text=text).grid(row=i, column=0, padx=10, pady=10, sticky="w")
            if key == "status":
                self.student_status_combobox = ttk.Combobox(form_frame, values=["فعال", "غیرفعال"], state="readonly")
                self.student_status_combobox.grid(row=i, column=1, padx=10, pady=10, sticky="ew")
                self.student_status_combobox.set("فعال")
            else:
                entry = ttk.Entry(form_frame, width=35)
                entry.grid(row=i, column=1, padx=10, pady=10, sticky="ew")
                self.student_entries[key] = entry

        form_frame.grid_columnconfigure(1, weight=1) # اجازه می‌دهد ورودی‌ها عریض شوند

        # --- بخش دکمه‌ها (اصلاح شده با grid) ---
        button_frame = ttk.Frame(form_frame, style='TFrame')
        button_frame.grid(row=len(self.student_labels), column=0, columnspan=2, pady=25, sticky="ew")

        # پیکربندی ستون‌های داخل فریم دکمه‌ها برای توزیع مساوی فضا
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)
        button_frame.grid_columnconfigure(3, weight=1)

        # استفاده از grid برای چیدمان دکمه‌ها برای پایداری بیشتر
        btn_config = {'padx': 4, 'sticky': 'ew'}
        ttk.Button(button_frame, text="💾 ذخیره", command=self.save_student).grid(row=0, column=0, **btn_config)
        ttk.Button(button_frame, text="✏️ ویرایش", command=self.update_student).grid(row=0, column=1, **btn_config)
        ttk.Button(button_frame, text="🗑️ حذف", command=self.delete_student).grid(row=0, column=2, **btn_config)
        ttk.Button(button_frame, text="✨ پاک کردن", command=self.clear_student_form).grid(row=0, column=3, **btn_config)

        # --- بخش لیست در طرف دیگر PanedWindow ---
        list_frame = ttk.LabelFrame(main_pane, text="لیست دانش‌آموزان", padding=25)

        search_frame = ttk.Frame(list_frame)
        search_frame.pack(fill="x", pady=10, padx=5)
        ttk.Label(search_frame, text="جستجو:").pack(side="right", padx=10)
        self.student_search_entry = ttk.Entry(search_frame)
        self.student_search_entry.pack(side="right", expand=True, fill="x", padx=5)
        self.student_search_entry.bind("<KeyRelease>", self.search_students)

        columns = ("id", "نام", "نام خانوادگی", "کد ملی", "شماره تماس", "کلاس", "وضعیت")
        self.student_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        self.student_tree.pack(fill="both", expand=True, pady=(10, 0))

        headings = {"id": "ID", "نام": "نام", "نام خانوادگی": "نام خانوادگی", "کد ملی": "کد ملی", "شماره تماس": "شماره تماس", "کلاس": "کلاس", "وضعیت": "وضعیت"}
        widths = {"id": 50, "نام": 130, "نام خانوادگی": 170, "کد ملی": 130, "شماره تماس": 130, "کلاس": 110, "وضعیت": 90}
        for col, text in headings.items():
            self.student_tree.heading(col, text=text, anchor="center")
            self.student_tree.column(col, width=widths.get(col, 100), anchor="center")

        scrollbar = ttk.Scrollbar(self.student_tree, orient="vertical", command=self.student_tree.yview)
        self.student_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.student_tree.bind("<Double-1>", self.load_student_to_form)

        # اضافه کردن فریم‌ها به PanedWindow
        # weight مشخص می‌کند چه نسبتی از فضا به هر بخش اختصاص یابد
        main_pane.add(list_frame, weight=3) # list_frame is typically the left/first pane
        # Set a minimum size for the form_frame pane to ensure buttons and fields remain visible
        # Estimated minimum content width: 320px (for buttons)
        # form_frame padding is 25 on each side, so 320 + 2*25 = 370px
        main_pane.add(form_frame, weight=1, minsize=370)

        return student_tab

    def load_students(self, query=None):
        for item in self.student_tree.get_children(): self.student_tree.delete(item)
        if query:
             sql_query = "SELECT id, name, family_name, national_id, parent_phone, class_name, status FROM students WHERE lower(name) LIKE ? OR lower(family_name) LIKE ? OR national_id LIKE ? ORDER BY family_name"
             params = (f"%{query}%", f"%{query}%", f"%{query}%")
        else:
            sql_query = "SELECT id, name, family_name, national_id, parent_phone, class_name, status FROM students ORDER BY family_name"
            params = ()

        self.cursor.execute(sql_query, params)
        students = []
        for student in self.cursor.fetchall():
            students.append((*student[:6], "فعال" if student[6] == "active" else "غیرفعال"))
        self._add_items_to_treeview(self.student_tree, students)

    def search_students(self, event=None):
        query = self.student_search_entry.get().strip().lower()
        self.load_students(query)

    def get_student_form_data(self):
        data = {key: entry.get().strip() for key, entry in self.student_entries.items()}
        data['status'] = "active" if self.student_status_combobox.get() == "فعال" else "inactive"

        if not all([data['name'], data['family_name'], data['national_id'], data['entry_date']]):
            messagebox.showerror("خطا", "لطفاً فیلدهای الزامی (نام، نام خانوادگی، کد ملی و تاریخ ثبت‌نام) را وارد کنید.")
            return None
        if data['national_id'] and (not data['national_id'].isdigit() or len(data['national_id']) != 10):
             messagebox.showerror("خطا", "کد ملی باید یک عدد ۱۰ رقمی باشد.")
             return None

        ordered_data = (
            data['name'], data['family_name'], data['national_id'], data['birth_date'],
            data['entry_date'], data['father_name'], data['mother_name'], data['parent_phone'],
            data['address'], data['class_name'], data['status'], data['allergies'], data['notes']
        )
        return ordered_data

    def save_student(self):
        data = self.get_student_form_data()
        if not data: return
        try:
            self.cursor.execute("INSERT INTO students (name, family_name, national_id, birth_date, entry_date, father_name, mother_name, parent_phone, address, class_name, status, allergies, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)
            self.conn.commit()
            messagebox.showinfo("موفقیت", "دانش‌آموز با موفقیت ذخیره شد.")
            self.clear_student_form(); self.load_students(); self.update_all_student_comboboxes()
        except sqlite3.IntegrityError:
            messagebox.showerror("خطا", "کد ملی وارد شده تکراری است.")
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ذخیره دانش‌آموز: {e}")

    def update_student(self):
        selected_item = self.student_tree.focus()
        if not selected_item: return messagebox.showwarning("انتخاب نشده", "لطفا یک دانش‌آموز را برای ویرایش انتخاب کنید.")
        student_id = self.student_tree.item(selected_item)['values'][0]
        data = self.get_student_form_data()
        if not data: return
        try:
            self.cursor.execute("UPDATE students SET name=?, family_name=?, national_id=?, birth_date=?, entry_date=?, father_name=?, mother_name=?, parent_phone=?, address=?, class_name=?, status=?, allergies=?, notes=? WHERE id=?", (*data, student_id))
            self.conn.commit()
            messagebox.showinfo("موفقیت", "اطلاعات دانش‌آموز به‌روزرسانی شد.")
            self.clear_student_form(); self.load_students(); self.update_all_student_comboboxes()
        except sqlite3.IntegrityError:
            messagebox.showerror("خطا", "کد ملی وارد شده تکراری است.")
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در به‌روزرسانی دانش‌آموز: {e}")

    def delete_student(self):
        selected_item = self.student_tree.focus()
        if not selected_item: return messagebox.showwarning("انتخاب نشده", "لطفا یک دانش‌آموز را برای حذف انتخاب کنید.")
        student_id = self.student_tree.item(selected_item)['values'][0]
        if messagebox.askyesno("تأیید حذف", "آیا مطمئن هستید؟ (تمام اطلاعات مالی و ارتباطی این دانش‌آموز نیز حذف خواهد شد)"):
            try:
                self.cursor.execute("DELETE FROM students WHERE id = ?", (student_id,))
                self.conn.commit()
                messagebox.showinfo("موفقیت", "دانش‌آموز با موفقیت حذف شد.")
                self.clear_student_form(); self.load_students(); self.update_all_student_comboboxes()
            except Exception as e:
                messagebox.showerror("خطا", f"خطا در حذف دانش‌آموز: {e}")

    def load_student_to_form(self, event):
        selected_item = self.student_tree.focus()
        if not selected_item: return
        student_id = self.student_tree.item(selected_item)['values'][0]
        self.cursor.execute("SELECT * FROM students WHERE id = ?", (student_id,))
        student_data = self.cursor.fetchone()
        if student_data:
            self.clear_student_form()
            entry_map = {
                "name": student_data[1], "family_name": student_data[2], "national_id": student_data[3],
                "birth_date": student_data[4], "entry_date": student_data[5], "father_name": student_data[6],
                "mother_name": student_data[7], "parent_phone": student_data[8], "address": student_data[9],
                "class_name": student_data[10], "allergies": student_data[12], "notes": student_data[13]
            }
            for key, widget in self.student_entries.items():
                widget.insert(0, str(entry_map.get(key) or ""))
            self.student_status_combobox.set("فعال" if student_data[11] == "active" else "غیرفعال")

    def clear_student_form(self):
        for widget in self.student_entries.values():
            if isinstance(widget, ttk.Entry): widget.delete(0, tk.END)
        self.student_status_combobox.set("فعال")

    # --- تب پرسنل (Staff) ---
    def create_staff_tab(self):
        staff_tab = ttk.Frame(self.content_frame, style='TFrame', padding=25)

        main_pane = ttk.PanedWindow(staff_tab, orient=tk.HORIZONTAL)
        main_pane.pack(fill="both", expand=True)

        form_frame = ttk.LabelFrame(main_pane, text="فرم ثبت/ویرایش پرسنل", padding=25)

        self.staff_labels = {
            "name": "نام:", "family_name": "نام خانوادگی:", "national_id": "کد ملی:",
            "birth_date": "تاریخ تولد:", "hire_date": "تاریخ استخدام:", "phone": "شماره تماس:",
            "address": "آدرس:", "role": "سمت:", "salary": "حقوق (ریال):",
            "status": "وضعیت:", "notes": "توضیحات:"
        }
        self.staff_entries = {}
        for i, (key, text) in enumerate(self.staff_labels.items()):
            ttk.Label(form_frame, text=text).grid(row=i, column=0, padx=10, pady=10, sticky="w")
            if key == "status":
                self.staff_status_combobox = ttk.Combobox(form_frame, values=["فعال", "غیرفعال"], state="readonly")
                self.staff_status_combobox.grid(row=i, column=1, padx=10, pady=10, sticky="ew")
                self.staff_status_combobox.set("فعال")
            else:
                entry = ttk.Entry(form_frame, width=35)
                entry.grid(row=i, column=1, padx=10, pady=10, sticky="ew")
                self.staff_entries[key] = entry

        form_frame.grid_columnconfigure(1, weight=1)

        button_frame = ttk.Frame(form_frame, style='TFrame')
        button_frame.grid(row=len(self.staff_labels), column=0, columnspan=2, pady=25, sticky="ew")

        # پیکربندی ستون‌ها برای توزیع مساوی فضا
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)
        button_frame.grid_columnconfigure(3, weight=1)

        btn_config = {'padx': 4, 'sticky': 'ew'}
        ttk.Button(button_frame, text="💾 ذخیره", command=self.save_staff).grid(row=0, column=0, **btn_config)
        ttk.Button(button_frame, text="✏️ ویرایش", command=self.update_staff).grid(row=0, column=1, **btn_config)
        ttk.Button(button_frame, text="🗑️ حذف", command=self.delete_staff).grid(row=0, column=2, **btn_config)
        ttk.Button(button_frame, text="✨ پاک کردن", command=self.clear_staff_form).grid(row=0, column=3, **btn_config)

        list_frame = ttk.LabelFrame(main_pane, text="لیست پرسنل", padding=25)

        search_frame = ttk.Frame(list_frame); search_frame.pack(fill="x", pady=10, padx=5)
        ttk.Label(search_frame, text="جستجو:").pack(side="right", padx=10)
        self.staff_search_entry = ttk.Entry(search_frame); self.staff_search_entry.pack(side="right", expand=True, fill="x", padx=5); self.staff_search_entry.bind("<KeyRelease>", self.search_staff)

        columns = ("id", "نام", "نام خانوادگی", "کد ملی", "شماره تماس", "سمت", "حقوق", "وضعیت")
        self.staff_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse"); self.staff_tree.pack(fill="both", expand=True, pady=(10,0))
        headings = {"id": "ID", "نام": "نام", "نام خانوادگی": "نام خانوادگی", "کد ملی": "کد ملی", "شماره تماس": "شماره تماس", "سمت": "سمت", "حقوق": "حقوق", "وضعیت": "وضعیت"}
        widths = {"id": 50, "نام": 130, "نام خانوادگی": 170, "کد ملی": 130, "شماره تماس": 130, "سمت": 120, "حقوق": 130, "وضعیت": 90}
        for col, text in headings.items(): self.staff_tree.heading(col, text=text, anchor="center"); self.staff_tree.column(col, width=widths.get(col, 100), anchor="center")

        scrollbar = ttk.Scrollbar(self.staff_tree, orient="vertical", command=self.staff_tree.yview); self.staff_tree.configure(yscrollcommand=scrollbar.set); scrollbar.pack(side="right", fill="y"); self.staff_tree.bind("<Double-1>", self.load_staff_to_form)

        main_pane.add(list_frame, weight=3)
        main_pane.add(form_frame, weight=1)

        return staff_tab

    def load_staff(self, query=None):
        for item in self.staff_tree.get_children(): self.staff_tree.delete(item)
        if query:
             sql_query = "SELECT id, name, family_name, national_id, phone, role, salary, status FROM staff WHERE lower(name) LIKE ? OR lower(family_name) LIKE ? OR national_id LIKE ? ORDER BY family_name"
             params = (f"%{query}%", f"%{query}%", f"%{query}%")
        else:
            sql_query = "SELECT id, name, family_name, national_id, phone, role, salary, status FROM staff ORDER BY family_name"
            params = ()

        self.cursor.execute(sql_query, params)
        staff_members = []
        for staff_member in self.cursor.fetchall(): # Renamed variable to avoid conflict
            formatted_salary = f"{staff_member[6]:,.0f}" if staff_member[6] else "0"
            staff_members.append((*staff_member[:6], formatted_salary, "فعال" if staff_member[7] == "active" else "غیرفعال"))
        self._add_items_to_treeview(self.staff_tree, staff_members)

    def search_staff(self, event=None):
        query = self.staff_search_entry.get().strip().lower()
        self.load_staff(query)

    def get_staff_form_data(self):
        data = {key: entry.get().strip() for key, entry in self.staff_entries.items()}
        data['status'] = "active" if self.staff_status_combobox.get() == "فعال" else "inactive"

        try:
            salary_str = data['salary']
            data['salary'] = float(salary_str) if salary_str else 0.0
        except ValueError:
            messagebox.showerror("خطا", "لطفا حقوق را به صورت عددی معتبر وارد کنید.")
            return None

        if not all([data['name'], data['family_name'], data['national_id'], data['hire_date']]):
            messagebox.showerror("خطا", "لطفاً فیلدهای الزامی (نام، نام خانوادگی، کد ملی و تاریخ استخدام) را وارد کنید.")
            return None
        if data['national_id'] and (not data['national_id'].isdigit() or len(data['national_id']) != 10):
             messagebox.showerror("خطا", "کد ملی باید یک عدد ۱۰ رقمی باشد.")
             return None

        ordered_data = (
            data['name'], data['family_name'], data['national_id'], data['birth_date'],
            data['hire_date'], data['phone'], data['address'], data['role'],
            data['salary'], data['status'], data['notes']
        )
        return ordered_data

    def save_staff(self):
        data = self.get_staff_form_data()
        if not data: return
        try:
            self.cursor.execute("INSERT INTO staff (name, family_name, national_id, birth_date, hire_date, phone, address, role, salary, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)
            self.conn.commit()
            messagebox.showinfo("موفقیت", "پرسنل با موفقیت ذخیره شد.")
            self.clear_staff_form(); self.load_staff()
        except sqlite3.IntegrityError:
            messagebox.showerror("خطا", "کد ملی وارد شده تکراری است.")
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ذخیره پرسنل: {e}")

    def update_staff(self):
        selected_item = self.staff_tree.focus()
        if not selected_item: return messagebox.showwarning("انتخاب نشده", "لطفا یک پرسنل را برای ویرایش انتخاب کنید.")
        staff_id = self.staff_tree.item(selected_item)['values'][0]
        data = self.get_staff_form_data()
        if not data: return
        try:
            self.cursor.execute("UPDATE staff SET name=?, family_name=?, national_id=?, birth_date=?, hire_date=?, phone=?, address=?, role=?, salary=?, status=?, notes=? WHERE id=?", (*data, staff_id))
            self.conn.commit()
            messagebox.showinfo("موفقیت", "اطلاعات پرسنل به‌روزرسانی شد.")
            self.clear_staff_form(); self.load_staff()
        except sqlite3.IntegrityError: messagebox.showerror("خطا", "کد ملی وارد شده تکراری است.")
        except Exception as e: messagebox.showerror("خطا", f"خطا در به‌روزرسانی پرسنل: {e}")

    def delete_staff(self):
        selected_item = self.staff_tree.focus()
        if not selected_item: return messagebox.showwarning("انتخاب نشده", "لطفا یک پرسنل را برای حذف انتخاب کنید.")
        staff_id = self.staff_tree.item(selected_item)['values'][0]
        if messagebox.askyesno("تأیید حذف", "آیا مطمئن هستید؟"):
            try:
                self.cursor.execute("DELETE FROM staff WHERE id = ?", (staff_id,))
                self.conn.commit()
                messagebox.showinfo("موفقیت", "پرسنل با موفقیت حذف شد.")
                self.clear_staff_form(); self.load_staff()
            except Exception as e: messagebox.showerror("خطا", f"خطا در حذف پرسنل: {e}")

    def load_staff_to_form(self, event):
        selected_item = self.staff_tree.focus()
        if not selected_item: return
        staff_id = self.staff_tree.item(selected_item)['values'][0]
        self.cursor.execute("SELECT * FROM staff WHERE id = ?", (staff_id,))
        staff_data = self.cursor.fetchone()
        if staff_data:
            self.clear_staff_form()
            entry_map = {
                "name": staff_data[1], "family_name": staff_data[2], "national_id": staff_data[3],
                "birth_date": staff_data[4], "hire_date": staff_data[5], "phone": staff_data[6],
                "address": staff_data[7], "role": staff_data[8], "salary": staff_data[9],
                "notes": staff_data[11]
            }
            for key, widget in self.staff_entries.items():
                widget.insert(0, str(entry_map.get(key) or ""))
            self.staff_status_combobox.set("فعال" if staff_data[10] == "active" else "غیرفعال")

    def clear_staff_form(self):
        for widget in self.staff_entries.values(): widget.delete(0, tk.END)
        self.staff_status_combobox.set("فعال")

    def create_attendance_tab(self):
        attendance_tab = ttk.Frame(self.content_frame, style='TFrame', padding=25)
        control_frame = ttk.LabelFrame(attendance_tab, text="کنترل حضور و غیاب", padding=20)
        control_frame.pack(side="top", fill="x", padx=5, pady=5)
        ttk.Button(control_frame, text="🔄 بارگذاری لیست", command=self.load_attendance_list).pack(side="left", padx=15)
        self.attendance_type_combobox = ttk.Combobox(control_frame, values=["دانش‌آموزان", "پرسنل"], state="readonly", width=15, font=self.BODY_FONT)
        self.attendance_type_combobox.pack(side="right", padx=5)
        self.attendance_type_combobox.set("دانش‌آموزان")
        self.attendance_type_combobox.bind("<<ComboboxSelected>>", self.load_attendance_list)
        ttk.Label(control_frame, text="نوع:").pack(side="right", padx=10)
        self.attendance_date_entry = ttk.Entry(control_frame, width=15, font=self.BODY_FONT)
        self.attendance_date_entry.pack(side="right", padx=5)
        self.attendance_date_entry.insert(0, jdatetime.date.today().strftime('%Y/%m/%d'))
        ttk.Label(control_frame, text="تاریخ:").pack(side="right", padx=10)
        tree_frame = ttk.Frame(attendance_tab, padding=5)
        tree_frame.pack(fill="both", expand=True, pady=(15, 0))
        columns = ("نام", "نام خانوادگی", "وضعیت", "زمان ورود", "زمان خروج")
        self.attendance_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        self.attendance_tree.pack(fill="both", expand=True)
        for col in columns: self.attendance_tree.heading(col, text=col, anchor="center"); self.attendance_tree.column(col, anchor="center")
        self.attendance_tree.bind("<Double-1>", self.show_attendance_update_dialog)
        return attendance_tab

    def load_attendance_list(self, event=None):
        for item in self.attendance_tree.get_children(): self.attendance_tree.delete(item)
        date = self.attendance_date_entry.get().strip()
        if not re.match(r'^\d{4}/\d{2}/\d{2}$', date):
            return messagebox.showerror("خطای فرمت", "فرمت تاریخ باید YYYY/MM/DD باشد.")
        person_type_fa = self.attendance_type_combobox.get()
        person_type = "student" if person_type_fa == "دانش‌آموزان" else "staff"
        table_name = "students" if person_type == "student" else "staff"
        try:
            self.cursor.execute(f"SELECT id, name, family_name FROM {table_name} WHERE status = 'active'")
            active_people = self.cursor.fetchall()
            self.cursor.execute("SELECT person_id, entry_time, exit_time, status FROM attendance WHERE date = ? AND person_type = ?", (date, person_type))
            existing_attendance = {row[0]: {'entry': row[1], 'exit': row[2], 'status': row[3]} for row in self.cursor.fetchall()}
            missing_people = []
            for person_id, _, _ in active_people:
                if person_id not in existing_attendance:
                    missing_people.append((person_id, person_type, date, "absent"))
            if missing_people:
                self.cursor.executemany("INSERT OR IGNORE INTO attendance (person_id, person_type, date, status) VALUES (?, ?, ?, ?)", missing_people)
                self.conn.commit()
                self.cursor.execute("SELECT person_id, entry_time, exit_time, status FROM attendance WHERE date = ? AND person_type = ?", (date, person_type))
                existing_attendance = {row[0]: {'entry': row[1], 'exit': row[2], 'status': row[3]} for row in self.cursor.fetchall()}
            items_to_add = []
            status_map = {"present": "حاضر", "absent": "غایب", "leave": "مرخصی"}
            for person_id, name, family_name in active_people:
                iid = f"{person_type}_{person_id}"
                record = existing_attendance.get(person_id)
                if record:
                    status = status_map.get(record['status'], "نامشخص")
                    entry = record['entry'] or "---"
                    exit_time = record['exit'] or "---"
                    item_values = (name, family_name, status, entry, exit_time)
                    items_to_add.append((iid, item_values))
            self._add_items_to_treeview(self.attendance_tree, items_to_add, has_iid=True)
        except sqlite3.Error as e:
            messagebox.showerror("خطای دیتابیس", f"خطا در بارگذاری لیست حضور و غیاب: {e}")

    def show_attendance_update_dialog(self, event):
        selected_item_iid = self.attendance_tree.focus()
        if not selected_item_iid: return
        item_data = self.attendance_tree.item(selected_item_iid)
        try:
            person_type, person_id_str = selected_item_iid.split('_')
            person_id = int(person_id_str)
            name, family_name, current_status, current_entry, current_exit = item_data['values']
        except (ValueError, IndexError) as e:
            messagebox.showerror("خطا", f"خطا در پردازش اطلاعات ردیف: {e}")
            return
        dialog = tk.Toplevel(self)
        dialog.title("به‌روزرسانی وضعیت")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg=self.BG_COLOR)
        dialog_frame = ttk.Frame(dialog, padding=20)
        dialog_frame.pack(expand=True, fill="both")
        dialog_frame.columnconfigure(1, weight=1)
        ttk.Label(dialog_frame, text=f"{name} {family_name}", font=self.HEADER_FONT, anchor='center').grid(row=0, column=0, columnspan=2, pady=(0, 15))
        ttk.Label(dialog_frame, text="وضعیت جدید:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        status_combobox = ttk.Combobox(dialog_frame, values=["حاضر", "غایب", "مرخصی"], state="readonly", font=self.BODY_FONT)
        status_combobox.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        status_combobox.set(current_status)
        ttk.Label(dialog_frame, text="زمان ورود (HH:MM):").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        entry_time_entry = ttk.Entry(dialog_frame)
        entry_time_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        if current_entry != "---": entry_time_entry.insert(0, current_entry)
        ttk.Label(dialog_frame, text="زمان خروج (HH:MM):").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        exit_time_entry = ttk.Entry(dialog_frame)
        exit_time_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        if current_exit != "---": exit_time_entry.insert(0, current_exit)
        def update_time_entry_state(event=None):
            state = "normal" if status_combobox.get() == "حاضر" else "disabled"
            entry_time_entry.config(state=state)
            exit_time_entry.config(state=state)
            if state == "disabled":
                entry_time_entry.delete(0, tk.END); exit_time_entry.delete(0, tk.END)
        status_combobox.bind("<<ComboboxSelected>>", update_time_entry_state)
        update_time_entry_state()
        def save_update():
            new_status_fa = status_combobox.get()
            new_status_en = {"حاضر": "present", "غایب": "absent", "مرخصی": "leave"}[new_status_fa]
            new_entry = entry_time_entry.get().strip() or None
            new_exit = exit_time_entry.get().strip() or None
            time_regex = r'^\d{2}:\d{2}$'
            if new_status_en == "present":
                if not new_entry:
                    return messagebox.showerror("خطا", "برای وضعیت 'حاضر'، زمان ورود الزامی است.", parent=dialog)
                if not re.match(time_regex, new_entry):
                    return messagebox.showerror("خطا", "فرمت زمان ورود باید HH:MM باشد.", parent=dialog)
                if new_exit and not re.match(time_regex, new_exit):
                    return messagebox.showerror("خطا", "فرمت زمان خروج باید HH:MM باشد.", parent=dialog)
            else:
                new_entry, new_exit = None, None
            self.update_attendance(person_id, person_type, new_status_en, new_entry, new_exit)
            dialog.destroy()
            self.load_attendance_list()
        button_frame = ttk.Frame(dialog_frame, style='TFrame')
        button_frame.grid(row=4, column=0, columnspan=2, pady=(20, 0))
        ttk.Button(button_frame, text="💾 ذخیره تغییرات", command=save_update).pack(expand=True)
        self.center_toplevel(dialog)

    def update_attendance(self, person_id, person_type, status, entry_time, exit_time):
        date = self.attendance_date_entry.get().strip()
        try:
            self.cursor.execute("""
                UPDATE attendance SET entry_time = ?, exit_time = ?, status = ?
                WHERE person_id = ? AND person_type = ? AND date = ?
            """, (entry_time, exit_time, status, person_id, person_type, date))
            self.conn.commit()
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در به‌روزرسانی حضور و غیاب: {e}", parent=self)

    def create_finance_tab(self):
        self.finance_tab = ttk.Frame(self.content_frame, style='TFrame', padding=10)
        main_pane = ttk.PanedWindow(self.finance_tab, orient=tk.VERTICAL)
        main_pane.pack(fill="both", expand=True)
        student_finance_frame = ttk.LabelFrame(main_pane, text="پرداختی دانش‌آموزان", padding=15)
        main_pane.add(student_finance_frame, weight=3)
        sf_form_frame = ttk.Frame(student_finance_frame); sf_form_frame.pack(fill="x", pady=5)
        self.finance_student_combobox = ttk.Combobox(sf_form_frame, state="readonly", font=self.BODY_FONT); self.finance_student_combobox.pack(side="right", padx=5, expand=True, fill="x")
        self.finance_student_combobox.bind("<<ComboboxSelected>>", self.load_finances)
        ttk.Label(sf_form_frame, text="دانش‌آموز:").pack(side="right", padx=5)
        sf_details_frame = ttk.Frame(student_finance_frame); sf_details_frame.pack(fill="x", pady=10)
        sf_grid = ttk.Frame(sf_details_frame); sf_grid.pack()
        ttk.Label(sf_grid, text="تاریخ پرداخت:").grid(row=0, column=3, padx=5, pady=5, sticky='w')
        self.finance_payment_date_entry = ttk.Entry(sf_grid); self.finance_payment_date_entry.grid(row=0, column=2, padx=5, pady=5, sticky='ew')
        self.finance_payment_date_entry.insert(0, jdatetime.date.today().strftime('%Y/%m/%d'))
        ttk.Label(sf_grid, text="مبلغ (ریال):").grid(row=0, column=1, padx=5, pady=5, sticky='w')
        self.finance_amount_entry = ttk.Entry(sf_grid); self.finance_amount_entry.grid(row=0, column=0, padx=5, pady=5, sticky='ew')
        ttk.Label(sf_grid, text="نوع پرداخت:").grid(row=1, column=3, padx=5, pady=5, sticky='w')
        self.finance_payment_type_combobox = ttk.Combobox(sf_grid, values=["شهریه", "غذا", "سرویس", "متفرقه"], state="readonly", font=self.BODY_FONT)
        self.finance_payment_type_combobox.grid(row=1, column=2, padx=5, pady=5, sticky='ew'); self.finance_payment_type_combobox.set("شهریه")
        ttk.Label(sf_grid, text="توضیحات:").grid(row=1, column=1, padx=5, pady=5, sticky='w')
        self.finance_description_entry = ttk.Entry(sf_grid); self.finance_description_entry.grid(row=1, column=0, padx=5, pady=5, sticky='ew')
        sf_button_frame = ttk.Frame(student_finance_frame); sf_button_frame.pack(fill="x", pady=10)
        ttk.Button(sf_button_frame, text="➕ ثبت پرداخت", command=self.save_finance).pack(side="right", padx=5)
        ttk.Button(sf_button_frame, text="🗑️ حذف پرداخت", command=self.delete_finance).pack(side="right", padx=5)
        ttk.Button(sf_button_frame, text="✨ پاک کردن فرم", command=self.clear_finance_form).pack(side="right", padx=5)
        finance_columns = ("id", "تاریخ", "مبلغ", "نوع", "توضیحات")
        self.finance_tree = ttk.Treeview(student_finance_frame, columns=finance_columns, show="headings", selectmode="browse")
        self.finance_tree.pack(fill="both", expand=True, pady=(5,0))
        for col in finance_columns: self.finance_tree.heading(col, text=col, anchor="center")
        self.finance_tree.column("id", width=40, anchor="center"); self.finance_tree.column("توضیحات", anchor="w", width=300); self.finance_tree.column("مبلغ", anchor="center"); self.finance_tree.column("نوع", anchor="center"); self.finance_tree.column("تاریخ", anchor="center")
        expenses_frame = ttk.LabelFrame(main_pane, text="هزینه‌های مهدکودک", padding=15)
        main_pane.add(expenses_frame, weight=2)
        exp_grid = ttk.Frame(expenses_frame); exp_grid.pack(fill='x', pady=5)
        ttk.Label(exp_grid, text="تاریخ:").grid(row=0, column=5, sticky='w', padx=5, pady=5)
        self.expense_date_entry = ttk.Entry(exp_grid); self.expense_date_entry.grid(row=0, column=4, sticky='ew', padx=5)
        self.expense_date_entry.insert(0, jdatetime.date.today().strftime('%Y/%m/%d'))
        ttk.Label(exp_grid, text="مبلغ (ریال):").grid(row=0, column=3, sticky='w', padx=5)
        self.expense_amount_entry = ttk.Entry(exp_grid); self.expense_amount_entry.grid(row=0, column=2, sticky='ew', padx=5)
        ttk.Label(exp_grid, text="دسته:").grid(row=0, column=1, sticky='w', padx=5)
        self.expense_category_combobox = ttk.Combobox(exp_grid, values=["اجاره", "حقوق", "قبوض", "خرید", "تعمیرات", "متفرقه"], state="readonly", font=self.BODY_FONT)
        self.expense_category_combobox.grid(row=0, column=0, sticky='ew', padx=5); self.expense_category_combobox.set("خرید")
        ttk.Label(exp_grid, text="توضیحات:").grid(row=1, column=5, sticky='w', padx=5, pady=5)
        self.expense_description_entry = ttk.Entry(exp_grid); self.expense_description_entry.grid(row=1, column=0, columnspan=5, sticky='ew', padx=5)
        exp_grid.columnconfigure(4, weight=1); exp_grid.columnconfigure(2, weight=1); exp_grid.columnconfigure(0, weight=1); exp_grid.columnconfigure(1, weight=1)
        exp_btn_frame = ttk.Frame(expenses_frame); exp_btn_frame.pack(fill='x', pady=10)
        ttk.Button(exp_btn_frame, text="➕ ثبت هزینه", command=self.save_expense).pack(side="right", padx=5)
        ttk.Button(exp_btn_frame, text="🗑️ حذف هزینه", command=self.delete_expense).pack(side="right", padx=5)
        ttk.Button(exp_btn_frame, text="✨ پاک کردن فرم", command=self.clear_expense_form).pack(side="right", padx=5)
        expenses_columns = ("id", "تاریخ", "مبلغ", "دسته", "توضیحات")
        self.expenses_tree = ttk.Treeview(expenses_frame, columns=expenses_columns, show="headings", selectmode="browse")
        self.expenses_tree.pack(fill="both", expand=True, pady=(5,0))
        for col in expenses_columns: self.expenses_tree.heading(col, text=col, anchor="center")
        self.expenses_tree.column("id", width=40, anchor="center"); self.expenses_tree.column("توضیحات", anchor="w", width=300); self.expenses_tree.column("مبلغ", anchor="center"); self.expenses_tree.column("دسته", anchor="center"); self.expenses_tree.column("تاریخ", anchor="center")
        return self.finance_tab

    def update_all_student_comboboxes(self):
        self.load_student_names_to_combobox(self.finance_student_combobox, self.load_finances)
        self.load_student_names_to_combobox(self.communication_student_combobox, self.load_communications)

    def load_student_names_to_combobox(self, combobox_widget, callback_func=None):
        current_selection_id = self.get_selected_student_id(combobox_widget)
        self.cursor.execute("SELECT id, name, family_name FROM students WHERE status = 'active' ORDER BY family_name")
        students = self.cursor.fetchall()
        self.student_name_to_id = {f"{s[1]} {s[2]} (ID: {s[0]})": s[0] for s in students}
        if combobox_widget is self.communication_student_combobox:
            all_students_key = "--- نمایش همه ---"
            all_students_dict = {all_students_key: None}
            all_students_dict.update(self.student_name_to_id)
            self.student_name_to_id = all_students_dict
        combobox_widget['values'] = list(self.student_name_to_id.keys())
        restored = False
        if current_selection_id is not None:
            for display_name, s_id in self.student_name_to_id.items():
                if s_id == current_selection_id:
                    combobox_widget.set(display_name)
                    restored = True
                    break
        if not restored and combobox_widget['values']:
            combobox_widget.set(combobox_widget['values'][0])
        if callback_func:
            callback_func()

    def get_selected_student_id(self, combobox_widget):
        selected_display = combobox_widget.get()
        return self.student_name_to_id.get(selected_display)

    def save_finance(self):
        student_id = self.get_selected_student_id(self.finance_student_combobox)
        if not student_id: return messagebox.showerror("خطا", "لطفا یک دانش‌آموز را انتخاب کنید.")
        date = self.finance_payment_date_entry.get().strip()
        amount_str = self.finance_amount_entry.get().strip()
        if not all([date, amount_str]): return messagebox.showerror("خطا", "لطفا تاریخ و مبلغ را وارد کنید.")
        if not re.match(r'^\d{4}/\d{2}/\d{2}$', date): return messagebox.showerror("خطای فرمت", "فرمت تاریخ پرداخت باید YYYY/MM/DD باشد.")
        try: amount = float(amount_str)
        except ValueError: return messagebox.showerror("خطا", "مبلغ باید عددی باشد.")
        try:
            self.cursor.execute("INSERT INTO student_finance (student_id, payment_date, payment_amount, payment_type, description) VALUES (?, ?, ?, ?, ?)",
                                (student_id, date, amount, self.finance_payment_type_combobox.get(), self.finance_description_entry.get().strip()))
            self.conn.commit()
            messagebox.showinfo("موفقیت", "پرداخت ثبت شد.")
            self.clear_finance_form(); self.load_finances(); self.generate_financial_report()
        except Exception as e: messagebox.showerror("خطا", f"خطا در ثبت پرداخت: {e}")

    def load_finances(self, event=None):
        for item in self.finance_tree.get_children(): self.finance_tree.delete(item)
        student_id = self.get_selected_student_id(self.finance_student_combobox)
        if not student_id: return
        self.cursor.execute("SELECT id, payment_date, payment_amount, payment_type, description FROM student_finance WHERE student_id = ? ORDER BY payment_date DESC", (student_id,))
        items = []
        for row in self.cursor.fetchall():
            formatted_amount = f"{row[2]:,.0f}"
            items.append((row[0], row[1], formatted_amount, row[3], row[4]))
        self._add_items_to_treeview(self.finance_tree, items)

    def delete_finance(self):
        selected_item = self.finance_tree.focus()
        if not selected_item: return messagebox.showwarning("انتخاب نشده", "لطفا یک پرداخت را برای حذف انتخاب کنید.")
        finance_id = self.finance_tree.item(selected_item)['values'][0]
        if messagebox.askyesno("تأیید حذف", "آیا مطمئن هستید؟"):
            self.cursor.execute("DELETE FROM student_finance WHERE id = ?", (finance_id,))
            self.conn.commit()
            messagebox.showinfo("موفقیت", "پرداخت حذف شد.")
            self.load_finances(); self.generate_financial_report()

    def clear_finance_form(self):
        self.finance_amount_entry.delete(0, tk.END)
        self.finance_description_entry.delete(0, tk.END)
        self.finance_payment_date_entry.delete(0, tk.END)
        self.finance_payment_date_entry.insert(0, jdatetime.date.today().strftime('%Y/%m/%d'))

    def save_expense(self):
        date = self.expense_date_entry.get().strip()
        amount_str = self.expense_amount_entry.get().strip()
        desc = self.expense_description_entry.get().strip()
        if not all([date, amount_str, desc]): return messagebox.showerror("خطا", "لطفا تمام فیلدها را پر کنید.")
        if not re.match(r'^\d{4}/\d{2}/\d{2}$', date): return messagebox.showerror("خطای فرمت", "فرمت تاریخ هزینه باید YYYY/MM/DD باشد.")
        try: amount = float(amount_str)
        except ValueError: return messagebox.showerror("خطا", "مبلغ باید عددی باشد.")
        try:
            self.cursor.execute("INSERT INTO kindergarten_expenses (date, description, amount, category) VALUES (?, ?, ?, ?)",
                                (date, desc, amount, self.expense_category_combobox.get()))
            self.conn.commit()
            messagebox.showinfo("موفقیت", "هزینه ثبت شد.")
            self.clear_expense_form(); self.load_expenses(); self.generate_financial_report()
        except Exception as e: messagebox.showerror("خطا", f"خطا در ثبت هزینه: {e}")

    def load_expenses(self):
        for item in self.expenses_tree.get_children(): self.expenses_tree.delete(item)
        self.cursor.execute("SELECT id, date, amount, category, description FROM kindergarten_expenses ORDER BY date DESC")
        items = []
        for row in self.cursor.fetchall():
            formatted_amount = f"{row[2]:,.0f}"
            items.append((row[0], row[1], formatted_amount, row[3], row[4]))
        self._add_items_to_treeview(self.expenses_tree, items)

    def delete_expense(self):
        selected_item = self.expenses_tree.focus()
        if not selected_item: return messagebox.showwarning("انتخاب نشده", "لطفا یک هزینه را برای حذف انتخاب کنید.")
        expense_id = self.expenses_tree.item(selected_item)['values'][0]
        if messagebox.askyesno("تأیید حذف", "آیا مطمئن هستید؟"):
            self.cursor.execute("DELETE FROM kindergarten_expenses WHERE id = ?", (expense_id,))
            self.conn.commit()
            messagebox.showinfo("موفقیت", "هزینه حذف شد.")
            self.load_expenses(); self.generate_financial_report()

    def clear_expense_form(self):
        self.expense_description_entry.delete(0, tk.END)
        self.expense_amount_entry.delete(0, tk.END)
        self.expense_date_entry.delete(0, tk.END)
        self.expense_date_entry.insert(0, jdatetime.date.today().strftime('%Y/%m/%d'))

    def create_communication_tab(self):
        self.communication_tab = ttk.Frame(self.content_frame, style='TFrame', padding=25)
        form_frame = ttk.LabelFrame(self.communication_tab, text="ثبت ارتباط با والدین", padding=25)
        form_frame.pack(side="right", fill="y", padx=(15, 0), pady=5)
        form_frame.grid_columnconfigure(1, weight=1)
        ttk.Label(form_frame, text="دانش‌آموز:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.communication_student_combobox = ttk.Combobox(form_frame, state="readonly", font=self.BODY_FONT)
        self.communication_student_combobox.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.communication_student_combobox.bind("<<ComboboxSelected>>", self.load_communications)
        ttk.Label(form_frame, text="تاریخ:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.comm_date_entry = ttk.Entry(form_frame); self.comm_date_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.comm_date_entry.insert(0, jdatetime.date.today().strftime('%Y/%m/%d'))
        ttk.Label(form_frame, text="موضوع:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.comm_subject_entry = ttk.Entry(form_frame); self.comm_subject_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        ttk.Label(form_frame, text="پیام:").grid(row=3, column=0, padx=10, pady=10, sticky="nw")
        self.comm_message_text = tk.Text(form_frame, height=10, width=40, font=self.BODY_FONT, relief='solid', borderwidth=1, highlightbackground="#DDDDDD", highlightthickness=1)
        self.comm_message_text.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        button_frame = ttk.Frame(form_frame, style='TFrame')
        button_frame.grid(row=4, column=0, columnspan=2, pady=25, sticky="ew")
        btn_config = {'padx': 8}
        ttk.Button(button_frame, text="💾 ذخیره", command=self.save_communication).pack(side="left", fill="x", expand=True, **btn_config)
        ttk.Button(button_frame, text="🗑️ حذف", command=self.delete_communication).pack(side="left", fill="x", expand=True, **btn_config)
        ttk.Button(button_frame, text="✨ پاک کردن", command=self.clear_communication_form).pack(side="left", fill="x", expand=True, **btn_config)
        list_frame = ttk.LabelFrame(self.communication_tab, text="لیست ارتباطات ثبت‌شده", padding=25)
        list_frame.pack(side="left", fill="both", expand=True, padx=(0, 15), pady=5)
        columns = ("id", "دانش‌آموز", "تاریخ", "موضوع", "پیام")
        self.communication_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        self.communication_tree.pack(fill="both", expand=True)
        for col in columns: self.communication_tree.heading(col, text=col, anchor='center')
        self.communication_tree.column("id", width=40, anchor="center"); self.communication_tree.column("دانش‌آموز", width=150, anchor="center")
        self.communication_tree.column("تاریخ", width=100, anchor="center"); self.communication_tree.column("موضوع", width=200, anchor='w')
        self.communication_tree.column("پیام", width=400, anchor='w')
        self.communication_tree.bind("<Double-1>", self.load_communication_to_form)
        return self.communication_tab

    def save_communication(self):
        student_id = self.get_selected_student_id(self.communication_student_combobox)
        if student_id is None: return messagebox.showerror("خطا", "لطفا یک دانش‌آموز را انتخاب کنید (گزینه 'نمایش همه' برای ثبت نیست).")
        date = self.comm_date_entry.get().strip()
        subject = self.comm_subject_entry.get().strip()
        message = self.comm_message_text.get("1.0", tk.END).strip()
        if not all([date, subject, message]): return messagebox.showerror("خطا", "لطفا تمام فیلدها را پر کنید.")
        if not re.match(r'^\d{4}/\d{2}/\d{2}$', date): return messagebox.showerror("خطای فرمت", "فرمت تاریخ باید YYYY/MM/DD باشد.")
        try:
            self.cursor.execute("INSERT INTO student_communications (student_id, date, subject, message) VALUES (?, ?, ?, ?)", (student_id, date, subject, message))
            self.conn.commit()
            messagebox.showinfo("موفقیت", "ارتباط ثبت شد.")
            self.clear_communication_form(); self.load_communications()
        except Exception as e: messagebox.showerror("خطا", f"خطا در ثبت ارتباط: {e}")

    def delete_communication(self):
        selected_item = self.communication_tree.focus()
        if not selected_item: return messagebox.showwarning("انتخاب نشده", "لطفا یک مورد را برای حذف انتخاب کنید.")
        comm_id = self.communication_tree.item(selected_item)['values'][0]
        if messagebox.askyesno("تأیید حذف", "آیا مطمئن هستید؟"):
            self.cursor.execute("DELETE FROM student_communications WHERE id = ?", (comm_id,))
            self.conn.commit()
            messagebox.showinfo("موفقیت", "ارتباط حذف شد."); self.load_communications()

    def load_communications(self, event=None):
        for item in self.communication_tree.get_children(): self.communication_tree.delete(item)
        student_id = self.get_selected_student_id(self.communication_student_combobox)
        query = "SELECT sc.id, s.name, s.family_name, sc.date, sc.subject, sc.message FROM student_communications sc JOIN students s ON sc.student_id = s.id"
        params = []
        if student_id is not None:
            query += " WHERE sc.student_id = ?"
            params.append(student_id)
        query += " ORDER BY sc.date DESC"
        self.cursor.execute(query, params)
        items = []
        for row in self.cursor.fetchall():
            student_full_name = f"{row[1]} {row[2]}"
            items.append((row[0], student_full_name, row[3], row[4], row[5]))
        self._add_items_to_treeview(self.communication_tree, items)

    def load_communication_to_form(self, event):
        selected_item = self.communication_tree.focus()
        if not selected_item: return
        comm_id = self.communication_tree.item(selected_item)['values'][0]
        self.cursor.execute("SELECT student_id, date, subject, message FROM student_communications WHERE id = ?", (comm_id,))
        comm_data = self.cursor.fetchone()
        if comm_data:
            self.clear_communication_form()
            student_id, date, subject, message = comm_data
            for display_name, s_id in self.student_name_to_id.items():
                if s_id == student_id:
                    self.communication_student_combobox.set(display_name)
                    break
            self.comm_date_entry.insert(0, date)
            self.comm_subject_entry.insert(0, subject)
            self.comm_message_text.insert("1.0", message)

    def clear_communication_form(self):
        self.comm_subject_entry.delete(0, tk.END)
        self.comm_message_text.delete("1.0", tk.END)
        self.comm_date_entry.delete(0, tk.END)
        self.comm_date_entry.insert(0, jdatetime.date.today().strftime('%Y/%m/%d'))

    def create_report_tab(self):
        report_tab = ttk.Frame(self.content_frame, style='TFrame', padding=25)
        finance_report_frame = ttk.LabelFrame(report_tab, text="گزارش مالی", padding=25)
        finance_report_frame.pack(side="top", fill="x", padx=5, pady=10)
        date_range_frame = ttk.Frame(finance_report_frame)
        date_range_frame.pack(pady=10)
        ttk.Button(date_range_frame, text="📊 نمایش گزارش", command=self.generate_financial_report).pack(side="left", padx=20)
        self.report_end_date = ttk.Entry(date_range_frame); self.report_end_date.pack(side="right", padx=5)
        self.report_end_date.insert(0, jdatetime.date.today().strftime('%Y/%m/%d'))
        ttk.Label(date_range_frame, text="تا تاریخ:").pack(side="right", padx=5)
        self.report_start_date = ttk.Entry(date_range_frame); self.report_start_date.pack(side="right", padx=5)
        self.report_start_date.insert(0, jdatetime.date.today().replace(day=1).strftime('%Y/%m/%d'))
        ttk.Label(date_range_frame, text="از تاریخ:").pack(side="right", padx=5)
        summary_frame = ttk.Frame(finance_report_frame)
        summary_frame.pack(fill="x", pady=20, expand=True)
        summary_font = (self.FONT_NAME, 16, 'bold')
        self.total_income_label = ttk.Label(summary_frame, text="درآمد کل: ۰", font=summary_font, foreground=self.SUCCESS_COLOR)
        self.total_income_label.pack(side="right", padx=20, expand=True)
        self.total_expense_label = ttk.Label(summary_frame, text="هزینه کل: ۰", font=summary_font, foreground=self.ERROR_COLOR)
        self.total_expense_label.pack(side="right", padx=20, expand=True)
        self.net_profit_label = ttk.Label(summary_frame, text="سود/زیان خالص: ۰", font=summary_font, foreground=self.PRIMARY_COLOR)
        self.net_profit_label.pack(side="right", padx=20, expand=True)
        export_frame = ttk.LabelFrame(report_tab, text="خروجی اطلاعات (اکسل)", padding=25)
        export_frame.pack(side="top", fill="x", padx=5, pady=10)
        btn_padding = {'padx': 10, 'pady': 10}
        ttk.Button(export_frame, text="خروجی دانش‌آموزان", command=lambda: self.export_data_to_excel("students")).grid(row=0, column=0, **btn_padding)
        ttk.Button(export_frame, text="خروجی پرسنل", command=lambda: self.export_data_to_excel("staff")).grid(row=0, column=1, **btn_padding)
        ttk.Button(export_frame, text="خروجی پرداختی‌ها", command=lambda: self.export_data_to_excel("student_finance")).grid(row=0, column=2, **btn_padding)
        ttk.Button(export_frame, text="خروجی هزینه‌ها", command=lambda: self.export_data_to_excel("kindergarten_expenses")).grid(row=1, column=0, **btn_padding)
        ttk.Button(export_frame, text="خروجی ارتباطات", command=lambda: self.export_data_to_excel("student_communications")).grid(row=1, column=1, **btn_padding)
        ttk.Button(export_frame, text="خروجی حضور و غیاب", command=lambda: self.export_data_to_excel("attendance")).grid(row=1, column=2, **btn_padding)
        export_frame.grid_columnconfigure((0,1,2), weight=1)
        return report_tab

    def generate_financial_report(self):
        start, end = self.report_start_date.get(), self.report_end_date.get()
        date_regex = r'^\d{4}/\d{2}/\d{2}$'
        if not re.match(date_regex, start) or not re.match(date_regex, end):
            return messagebox.showerror("خطای فرمت", "فرمت تاریخ باید YYYY/MM/DD باشد.")
        try:
            income = self.cursor.execute("SELECT SUM(payment_amount) FROM student_finance WHERE payment_date BETWEEN ? AND ?", (start, end)).fetchone()[0] or 0.0
            expense = self.cursor.execute("SELECT SUM(amount) FROM kindergarten_expenses WHERE date BETWEEN ? AND ?", (start, end)).fetchone()[0] or 0.0
            net = income - expense
            self.total_income_label.config(text=f"درآمد کل: {income:,.0f} ریال")
            self.total_expense_label.config(text=f"هزینه کل: {expense:,.0f} ریال")
            net_profit_color = self.SUCCESS_COLOR if net >= 0 else self.ERROR_COLOR
            self.net_profit_label.config(text=f"سود/زیان خالص: {net:,.0f} ریال", foreground=net_profit_color)
        except sqlite3.Error as e:
            messagebox.showerror("خطای دیتابیس", f"خطا در تولید گزارش مالی: {e}")

    def _get_person_name_map(self, table_name):
        self.cursor.execute(f"SELECT id, name, family_name FROM {table_name}")
        return {p_id: f"{name} {family}" for p_id, name, family in self.cursor.fetchall()}

    def export_data_to_excel(self, table_name):
        try:
            df = pd.read_sql_query(f"SELECT * FROM {table_name}", self.conn)
            if df.empty: return messagebox.showinfo("خالی", "داده‌ای برای خروجی وجود ندارد.")
            if table_name in ["student_finance", "student_communications"]:
                student_map = self._get_person_name_map("students")
                df['نام دانش‌آموز'] = df['student_id'].map(student_map)
            elif table_name == "attendance":
                student_map = self._get_person_name_map("students")
                staff_map = self._get_person_name_map("staff")
                df['نام'] = df.apply(lambda row: student_map.get(row['person_id']) if row['person_type'] == 'student' else staff_map.get(row['person_id']), axis=1)
            column_map = {
                "students": {'id': 'شناسه', 'name': 'نام', 'family_name': 'نام خانوادگی', 'national_id': 'کد ملی', 'birth_date': 'تاریخ تولد', 'entry_date': 'تاریخ ثبت‌نام', 'father_name': 'نام پدر', 'mother_name': 'نام مادر', 'parent_phone': 'شماره تماس', 'address': 'آدرس', 'class_name': 'کلاس', 'status': 'وضعیت', 'allergies': 'آلرژی‌ها', 'notes': 'توضیحات'},
                "staff": {'id': 'شناسه', 'name': 'نام', 'family_name': 'نام خانوادگی', 'national_id': 'کد ملی', 'birth_date': 'تاریخ تولد', 'hire_date': 'تاریخ استخدام', 'phone': 'شماره تماس', 'address': 'آدرس', 'role': 'سمت', 'salary': 'حقوق', 'status': 'وضعیت', 'notes': 'توضیحات'},
                "attendance": {'person_type': 'نوع شخص', 'date': 'تاریخ', 'entry_time': 'زمان ورود', 'exit_time': 'زمان خروج', 'status': 'وضعیت'},
                "student_finance": {'payment_date': 'تاریخ پرداخت', 'payment_amount': 'مبلغ', 'payment_type': 'نوع پرداخت', 'description': 'توضیحات'},
                "kindergarten_expenses": {'date': 'تاریخ', 'description': 'توضیحات', 'amount': 'مبلغ', 'category': 'دسته'},
                "student_communications": {'date': 'تاریخ', 'subject': 'موضوع', 'message': 'پیام'}
            }
            df.rename(columns=column_map.get(table_name, {}), inplace=True)
            if 'وضعیت' in df.columns: df['وضعیت'] = df['وضعیت'].map({'active': 'فعال', 'inactive': 'غیرفعال', 'present': 'حاضر', 'absent': 'غایب', 'leave': 'مرخصی'})
            if 'نوع شخص' in df.columns: df['نوع شخص'] = df['نوع شخص'].map({'student': 'دانش‌آموز', 'staff': 'پرسنل'})
            final_columns = {
                "students": ['شناسه', 'نام', 'نام خانوادگی', 'کد ملی', 'تاریخ ثبت‌نام', 'شماره تماس', 'کلاس', 'وضعیت', 'نام پدر', 'نام مادر', 'تاریخ تولد', 'آدرس', 'آلرژی‌ها', 'توضیحات'],
                "staff": ['شناسه', 'نام', 'نام خانوادگی', 'کد ملی', 'تاریخ استخدام', 'شماره تماس', 'سمت', 'حقوق', 'وضعیت', 'تاریخ تولد', 'آدرس', 'توضیحات'],
                "attendance": ['نام', 'نوع شخص', 'تاریخ', 'وضعیت', 'زمان ورود', 'زمان خروج'],
                "student_finance": ['نام دانش‌آموز', 'تاریخ پرداخت', 'مبلغ', 'نوع پرداخت', 'توضیحات'],
                "kindergarten_expenses": ['تاریخ', 'مبلغ', 'دسته', 'توضیحات'],
                "student_communications": ['نام دانش‌آموز', 'تاریخ', 'موضوع', 'پیام']
            }
            df = df[[col for col in final_columns.get(table_name, df.columns) if col in df.columns]]
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="ذخیره فایل اکسل")
            if file_path:
                df.to_excel(file_path, index=False, engine='openpyxl')
                messagebox.showinfo("موفقیت", f"اطلاعات با موفقیت به اکسل خروجی گرفته شد.")
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در خروجی گرفتن به اکسل: {e}")

    def center_toplevel(self, toplevel):
        toplevel.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (toplevel.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (toplevel.winfo_height() // 2)
        toplevel.geometry(f"+{x}+{y}")


if __name__ == "__main__":
    app = KindergartenApp()
    app.mainloop()
