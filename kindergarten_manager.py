# kindergarten_manager_with_login.py

# ====================================================================================
# توضیحات و پیش‌نیازها
# ====================================================================================
# این نسخه از برنامه دارای صفحه لاگین است.
# نام کاربری و کلمه عبور پیشفرض برای ورود: admin
# برای اجرای صحیح برنامه، لطفاً پیش‌نیازهای زیر را نصب کنید:
# pip install jdatetime tkcalendar pandas openpyxl matplotlib
# ====================================================================================

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import jdatetime
from datetime import datetime
import atexit
import re
import logging
import hashlib
import shutil
import threading
from tkcalendar import DateEntry
import os
import traceback
import sys

# ماژول‌ها با بارگذاری تنبل (Lazy Loading)
pd = None
openpyxl = None
Figure = None
FigureCanvasTkAgg = None

# تنظیمات اولیه لاگ‌گیری
logging.basicConfig(filename='kindergarten_app.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s', encoding='utf-8')

# یک استثنای سفارشی برای خطاهای بحرانی دیتابیس
class DatabaseCriticalError(Exception):
    pass

# ====================================================================================
# کلاس مدیریت دیتابیس (بخش Model)
# ====================================================================================
class DatabaseManager:
    """کلاس مدیریت تمام تعاملات با پایگاه داده SQLite."""
    def __init__(self, db_path="kindergarten.db"):
        self.db_path = db_path
        self.conn = None
        self.cursor = None

    def connect(self):
        try:
            # استفاده از check_same_thread=False برای پشتیبانی از Threading در عملیات پشتیبان‌گیری
            self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
            self.cursor = self.conn.cursor()
            logging.info("اتصال به پایگاه داده با موفقیت برقرار شد.")
            self.create_tables()
        except sqlite3.Error as e:
            logging.critical(f"CRITICAL DB ERROR: خطا در اتصال به پایگاه داده: {e}")
            raise DatabaseCriticalError(f"امکان اتصال به پایگاه داده وجود ندارد: {e}\nبرنامه بسته خواهد شد.")

    def close(self):
        if self.conn:
            self.conn.close()
            logging.info("اتصال به پایگاه داده بسته شد.")

    def commit(self):
        if self.conn: self.conn.commit()

    def rollback(self):
        if self.conn: self.conn.rollback()
        logging.warning("تراکنش دیتابیس به دلیل خطا Rollback شد.")

    def execute_query(self, query, params=(), fetch=None):
        try:
            self.cursor.execute(query, params)
            if fetch == "one": return self.cursor.fetchone()
            if fetch == "all": return self.cursor.fetchall()
            return True
        except sqlite3.Error as e:
            logging.error(f"خطا در اجرای کوئری '{query[:50]}...': {e}")
            if "UNIQUE constraint failed" in str(e):
                 messagebox.showerror("خطای داده تکراری", "کد ملی وارد شده قبلاً در سیستم ثبت شده است.")
            else:
                 messagebox.showerror("خطای دیتابیس", f"خطا در عملیات پایگاه داده: {e}")
            return None if fetch else False

    def create_tables(self):
        try:
            with self.conn:
                queries = [
                    "CREATE TABLE IF NOT EXISTS students (id INTEGER PRIMARY KEY, name TEXT NOT NULL, family_name TEXT NOT NULL, national_id TEXT UNIQUE, birth_date TEXT, entry_date TEXT NOT NULL, father_name TEXT, mother_name TEXT, parent_phone TEXT, address TEXT, class_name TEXT, status TEXT NOT NULL, allergies TEXT, notes TEXT)",
                    "CREATE TABLE IF NOT EXISTS staff (id INTEGER PRIMARY KEY, name TEXT NOT NULL, family_name TEXT NOT NULL, national_id TEXT UNIQUE, birth_date TEXT, hire_date TEXT NOT NULL, phone TEXT, address TEXT, role TEXT NOT NULL, salary REAL, status TEXT NOT NULL, notes TEXT)",
                    "CREATE TABLE IF NOT EXISTS attendance (id INTEGER PRIMARY KEY, person_id INTEGER NOT NULL, person_type TEXT NOT NULL, date TEXT NOT NULL, status TEXT, entry_time TEXT, exit_time TEXT, UNIQUE(person_id, person_type, date))",
                    "CREATE TABLE IF NOT EXISTS student_finance (id INTEGER PRIMARY KEY, student_id INTEGER NOT NULL, payment_date TEXT NOT NULL, payment_amount REAL NOT NULL, payment_type TEXT, description TEXT, FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE)",
                    "CREATE TABLE IF NOT EXISTS student_communications (id INTEGER PRIMARY KEY, student_id INTEGER NOT NULL, date TEXT NOT NULL, subject TEXT, message TEXT, FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE)",
                    "CREATE TABLE IF NOT EXISTS kindergarten_expenses (id INTEGER PRIMARY KEY, date TEXT NOT NULL, description TEXT, amount REAL NOT NULL, category TEXT)",
                    "CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)"
                ]
                for query in queries: self.cursor.execute(query)
                default_settings = {'kindergarten_name': 'مهدکودک من', 'kindergarten_address': 'آدرس پیش‌فرض', 'kindergarten_phone': 'تلفن پیش‌فرض'}
                for key, value in default_settings.items():
                    if not self.execute_query("SELECT value FROM settings WHERE key=?", (key,), fetch="one"):
                        self.execute_query("INSERT INTO settings (key, value) VALUES (?, ?)", (key, value))
        except sqlite3.Error as e:
            logging.critical(f"CRITICAL DB ERROR: خطا در ایجاد جداول: {e}")
            raise DatabaseCriticalError(f"خطا در ایجاد ساختار پایگاه داده: {e}\nبرنامه بسته خواهد شد.")

    def get_setting(self, key, default=""):
        result = self.execute_query("SELECT value FROM settings WHERE key=?", (key,), fetch="one")
        return result[0] if result else default

    def get_dashboard_stats(self):
        stats = {}
        stats['students'] = (self.execute_query("SELECT COUNT(id) FROM students WHERE status='active'", fetch="one") or [0])[0]
        stats['staff'] = (self.execute_query("SELECT COUNT(id) FROM staff WHERE status='active'", fetch="one") or [0])[0]
        stats['income'] = (self.execute_query("SELECT SUM(payment_amount) FROM student_finance", fetch="one") or [0])[0] or 0.0
        stats['expense'] = (self.execute_query("SELECT SUM(amount) FROM kindergarten_expenses", fetch="one") or [0])[0] or 0.0
        return stats
        
    def get_financial_report_data(self):
        data = {}
        data['total_income'] = (self.execute_query("SELECT SUM(payment_amount) FROM student_finance", fetch="one") or [0.0])[0] or 0.0
        data['total_expense'] = (self.execute_query("SELECT SUM(amount) FROM kindergarten_expenses", fetch="one") or [0.0])[0] or 0.0
        data['income_by_type'] = self.execute_query("SELECT payment_type, SUM(payment_amount) FROM student_finance GROUP BY payment_type", fetch="all") or []
        data['expense_by_category'] = self.execute_query("SELECT category, SUM(amount) FROM kindergarten_expenses GROUP BY category", fetch="all") or []
        return data

# ====================================================================================
# کلاس اصلی برنامه (بخش View و Controller)
# ====================================================================================
class KindergartenApp(ttk.Frame):
    # تعریف ثابت‌ها برای خوانایی بهتر
    STATUS_ACTIVE, STATUS_INACTIVE = 'active', 'inactive'
    PERSON_TYPE_STUDENT, PERSON_TYPE_STAFF = 'student', 'staff'
    ATTENDANCE_PRESENT, ATTENDANCE_ABSENT, ATTENDANCE_LEAVE = 'حاضر', 'غایب', 'مرخصی'
    
    # پالت رنگی و فونت‌ها
    BG_COLOR, FRAME_COLOR, PRIMARY_COLOR, ACCENT_COLOR = "#F4F6F7", "#FFFFFF", "#1ABC9C", "#16A085"
    NAV_BG_COLOR, NAV_SELECTED_BG = "#E8F6F3", "#FFFFFF"
    TEXT_COLOR, MUTED_TEXT_COLOR, HEADER_TEXT_COLOR = "#34495E", "#7F8C8D", "#FFFFFF"
    SUCCESS_COLOR, ERROR_COLOR, ROW_ALT_COLOR = "#2ECC71", "#E74C3C", "#F8F9F9"
    FONT_FAMILY = ("Vazirmatn", "Tahoma", "Arial")
    HEADER_FONT = (FONT_FAMILY[0], 15, 'bold')
    BODY_FONT = (FONT_FAMILY[0], 11)
    BOLD_FONT = (FONT_FAMILY[0], 11, 'bold')
    TAB_FONT = (FONT_FAMILY[0], 12, 'bold')
    DASHBOARD_LARGE_FONT = (FONT_FAMILY[0], 28, 'bold')
    TEXT_WIDGET_FONT = (FONT_FAMILY[0], 12)

    def __init__(self, parent, db_manager, user_info):
        super().__init__(parent)
        self.parent = parent
        self.db_manager = db_manager
        self.user_info = user_info
        self.student_map = {}

        self._lazy_load_modules()
        self.setup_styles()
        self.icons = self.load_icons()
        self.create_main_layout()
        self.create_pages()
        self.update_all_student_comboboxes()
        self.load_all_data()
        self.show_frame("داشبورد")

        self.pack(fill="both", expand=True)

    def _lazy_load_modules(self):
        """ماژول‌های سنگین را فقط در صورت نیاز بارگذاری می‌کند."""
        global pd, openpyxl, Figure, FigureCanvasTkAgg
        try:
            import pandas as pd
        except ImportError:
            logging.warning("کتابخانه 'pandas' نصب نیست. برخی قابلیت‌ها محدود خواهد بود.")
        try:
            import openpyxl
        except ImportError:
            logging.warning("کتابخانه 'openpyxl' نصب نیست. قابلیت خروجی اکسل غیرفعال خواهد بود.")
        try:
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        except ImportError:
            logging.warning("کتابخانه 'matplotlib' نصب نیست. قابلیت نمایش نمودار غیرفعال خواهد بود.")

    def setup_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure('.', background=self.BG_COLOR, foreground=self.TEXT_COLOR, font=self.BODY_FONT)
        self.style.configure('TFrame', background=self.FRAME_COLOR)
        self.style.configure('Nav.TFrame', background=self.NAV_BG_COLOR)
        self.style.configure('TLabelframe', font=self.BOLD_FONT, background=self.FRAME_COLOR, borderwidth=1, relief="solid")
        self.style.configure('TLabelframe.Label', font=self.HEADER_FONT, background=self.FRAME_COLOR, foreground=self.PRIMARY_COLOR, padding=(0, 0, 0, 10))
        self.style.configure('TLabel', font=self.BOLD_FONT, background=self.FRAME_COLOR, foreground=self.TEXT_COLOR)
        self.style.configure('Header.TLabel', font=self.HEADER_FONT, background=self.FRAME_COLOR, foreground=self.PRIMARY_COLOR)
        self.style.configure('TButton', font=self.BOLD_FONT, padding=12, background=self.PRIMARY_COLOR, foreground=self.HEADER_TEXT_COLOR, borderwidth=0, relief='flat', focusthickness=0)
        self.style.map('TButton', background=[('active', self.ACCENT_COLOR), ('!disabled', self.PRIMARY_COLOR)])
        self.style.configure('Nav.TButton', font=self.TAB_FONT, padding=(20, 15), background=self.NAV_BG_COLOR, foreground=self.MUTED_TEXT_COLOR, borderwidth=0, relief='flat', anchor='e', compound='right')
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

    def load_icons(self):
        return {"dashboard": "📊", "students": "👶", "staff": "👥", "attendance": "📅", "finance": "💰", "communication": "💬", "reports": "📈", "settings": "⚙️", "expense": "💸", "success": "✅", "error": "❌", "export": "📄"}

    def create_main_layout(self):
        main_frame = ttk.Frame(self, style='TFrame')
        main_frame.pack(expand=True, fill="both")
        nav_frame = ttk.Frame(main_frame, width=250, style='Nav.TFrame')
        nav_frame.pack(side="right", fill="y")
        nav_frame.pack_propagate(False)
        self.content_frame = ttk.Frame(main_frame, style='TFrame')
        self.content_frame.pack(side="right", expand=True, fill="both")
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.nav_frame = nav_frame

    def create_pages(self):
        self.tabs = {}
        page_creators = {"داشبورد": (self.create_dashboard_tab, self.icons.get('dashboard')), "دانش‌آموزان": (self.create_student_tab, self.icons.get('students')), "پرسنل": (self.create_staff_tab, self.icons.get('staff')), "حضور و غیاب": (self.create_attendance_tab, self.icons.get('attendance')), "امور مالی": (self.create_finance_tab, self.icons.get('finance')), "ارتباطات": (self.create_communication_tab, self.icons.get('communication')), "گزارشات": (self.create_report_tab, self.icons.get('reports'))}
        if self.user_info['role'] == 'admin': page_creators["تنظیمات"] = (self.create_settings_tab, self.icons.get('settings'))
        kindergarten_name = self.db_manager.get_setting('kindergarten_name', 'مهدکودک')
        self.title_label = ttk.Label(self.nav_frame, text=kindergarten_name, font=self.HEADER_FONT, background=self.NAV_BG_COLOR, foreground=self.PRIMARY_COLOR, padding=20)
        self.title_label.pack(pady=(10, 20))
        self.nav_buttons = {}
        for name, (creator_func, icon) in page_creators.items():
            frame = creator_func()
            self.tabs[name] = frame
            frame.grid(row=0, column=0, sticky="nsew")
            btn = ttk.Button(self.nav_frame, text=f"{icon}  {name}", command=lambda n=name: self.show_frame(n), style='Nav.TButton')
            btn.pack(fill="x", pady=2, padx=10)
            self.nav_buttons[name] = btn

    def show_frame(self, page_name):
        frame = self.tabs[page_name]
        frame.tkraise()
        for name, button in self.nav_buttons.items(): button.state(['!selected'])
        if page_name in self.nav_buttons: self.nav_buttons[page_name].state(['selected'])
        # Refresh data when switching to certain tabs
        if page_name == "داشبورد": self.refresh_dashboard()
        elif page_name == "گزارشات": self.generate_financial_report()
        elif page_name == "حضور و غیاب": self.load_attendance_list()

    def load_all_data(self):
        self.load_students(); self.load_staff(); self.load_student_finances(); self.load_expenses(); self.load_communications();

    def _add_items_to_treeview(self, tree, items, has_iid=False):
        tree.delete(*tree.get_children())
        for i, item in enumerate(items):
            if not item: 
                logging.warning(f"داده نامعتبر در آیتم {i} رد شد."); continue
            tag = 'oddrow' if i % 2 != 0 else 'evenrow'
            if has_iid: 
                iid, values = item
                tree.insert("", "end", iid=iid, values=values, tags=(tag,))
            else: 
                tree.insert("", "end", values=item, tags=(tag,))

    def create_date_entry(self, parent, **kwargs):
        # Using standard datetime because jdatetime dependency can be problematic
        return DateEntry(parent, width=12, background=self.PRIMARY_COLOR, foreground='white', borderwidth=2, calendar_kwargs={'font': self.BODY_FONT}, date_pattern='yyyy-mm-dd', font=self.BODY_FONT)

    # ====================================================================================
    # بخش داشبورد
    # ====================================================================================
    def create_dashboard_tab(self):
        dashboard_tab = ttk.Frame(self.content_frame, style='TFrame', padding=40)
        stats_frame = ttk.Frame(dashboard_tab); stats_frame.pack(fill='x', expand=True, pady=20)
        stats_frame.columnconfigure((0,1,2,3), weight=1)
        self.student_count_label = self._create_stat_card(stats_frame, 0, "دانش‌آموزان فعال", self.icons['students'])
        self.staff_count_label = self._create_stat_card(stats_frame, 1, "پرسنل فعال", self.icons['staff'])
        self.income_total_label = self._create_stat_card(stats_frame, 2, "مجموع درآمد (ریال)", self.icons['finance'])
        self.expense_total_label = self._create_stat_card(stats_frame, 3, "مجموع هزینه (ریال)", self.icons['expense'])
        info_frame = ttk.LabelFrame(dashboard_tab, text="خوش آمدید", padding=20)
        info_frame.pack(fill='both', expand=True, pady=20)
        # Display a welcome message with the username
        ttk.Label(info_frame, text=f"کاربر '{self.user_info['username']}' به سامانه مدیریت مهدکودک خوش آمدید!\n\nاز منوی سمت راست برای دسترسی به بخش‌های مختلف استفاده کنید.", wraplength=800, justify='right', font=self.HEADER_FONT, foreground=self.TEXT_COLOR).pack(anchor='center', expand=True)
        return dashboard_tab

    def _create_stat_card(self, parent, col, title, icon):
        card_frame = ttk.Frame(parent, style='TFrame', relief='solid', borderwidth=1)
        card_frame.grid(row=0, column=col, padx=15, pady=10, sticky='nsew')
        card_frame.columnconfigure(0, weight=1)
        icon_label = ttk.Label(card_frame, text=icon, font=(self.FONT_FAMILY[0], 30), anchor='center'); icon_label.pack(pady=(15, 5))
        title_label = ttk.Label(card_frame, text=title, font=self.BOLD_FONT, anchor='center', foreground=self.PRIMARY_COLOR); title_label.pack(pady=5)
        value_label = ttk.Label(card_frame, text="0", font=self.DASHBOARD_LARGE_FONT, anchor='center'); value_label.pack(pady=(5, 15))
        return value_label

    def refresh_dashboard(self):
        try:
            stats = self.db_manager.get_dashboard_stats()
            self.student_count_label.config(text=str(stats.get('students', 0)))
            self.staff_count_label.config(text=str(stats.get('staff', 0)))
            self.income_total_label.config(text=f"{stats.get('income', 0):,.0f}")
            self.expense_total_label.config(text=f"{stats.get('expense', 0):,.0f}")
            logging.info("داشبورد به‌روزرسانی شد")
        except Exception as e:
            logging.error(f"خطا در به‌روزرسانی داشبورد: {e}")
            messagebox.showerror("خطا", f"خطا در بارگذاری اطلاعات داشبورد: {e}")

    # ====================================================================================
    # بخش گزارشات
    # ====================================================================================
    def create_report_tab(self):
        report_tab = ttk.Frame(self.content_frame, style='TFrame', padding=25)
        report_tab.rowconfigure(1, weight=1)
        report_tab.columnconfigure((0, 1), weight=1)

        control_frame = ttk.Frame(report_tab)
        control_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 15))
        ttk.Button(control_frame, text="به‌روزرسانی گزارش", command=self.generate_financial_report).pack(side='right', padx=5)

        self.export_button = ttk.Button(control_frame, text=f"{self.icons['export']} خروجی اکسل", command=self.export_financial_report)
        self.export_button.pack(side='right', padx=5)
        if not (pd and openpyxl):
            self.export_button.config(state="disabled")

        summary_frame = ttk.LabelFrame(report_tab, text="خلاصه مالی", padding=15)
        summary_frame.grid(row=1, column=0, sticky='nsew', padx=(0, 10))
        self.report_text = tk.Text(summary_frame, font=self.TEXT_WIDGET_FONT, wrap='word', state='disabled', relief='flat', bg=self.FRAME_COLOR)
        self.report_text.pack(expand=True, fill='both')

        self.chart_frame = ttk.LabelFrame(report_tab, text="نمودار درآمد و هزینه", padding=15)
        self.chart_frame.grid(row=1, column=1, sticky='nsew')
        self.chart_canvas = None
        
        return report_tab

    def generate_financial_report(self):
        if not self.chart_canvas:
            if Figure and FigureCanvasTkAgg:
                self.report_fig = Figure(figsize=(5, 4), dpi=100)
                self.report_ax = self.report_fig.add_subplot(111)
                self.chart_canvas = FigureCanvasTkAgg(self.report_fig, master=self.chart_frame)
                self.chart_canvas.get_tk_widget().pack(expand=True, fill='both')
                self.report_fig.set_facecolor(self.FRAME_COLOR)
            else:
                ttk.Label(self.chart_frame, text="کتابخانه Matplotlib برای نمایش نمودار نصب نیست.").pack()
                logging.error("Matplotlib is not installed, cannot create chart.")
                self._update_report_text_only()
                return

        data = self.db_manager.get_financial_report_data()
        
        total_income = data.get('total_income', 0.0)
        total_expense = data.get('total_expense', 0.0)
        net_profit = total_income - total_expense
        
        report_str = f"خلاصه وضعیت مالی مهدکودک\n{'='*30}\n"
        report_str += f"مجموع درآمد کل: {total_income:,.0f} ریال\n"
        report_str += f"مجموع هزینه‌های کل: {total_expense:,.0f} ریال\n"
        report_str += f"{'-'*30}\n"
        report_str += f"سود/زیان خالص: {net_profit:,.0f} ریال\n\n{'='*30}\n"
        report_str += "درآمدها به تفکیک نوع:\n"
        for ptype, amount in data.get('income_by_type', []): report_str += f"  - {ptype or 'نامشخص'}: {amount:,.0f} ریال\n"
        report_str += f"\n{'='*30}\n"
        report_str += "هزینه‌ها به تفکیک دسته‌بندی:\n"
        for cat, amount in data.get('expense_by_category', []): report_str += f"  - {cat or 'نامشخص'}: {amount:,.0f} ریال\n"

        self.report_text.config(state='normal')
        self.report_text.delete('1.0', tk.END)
        self.report_text.insert('1.0', report_str)
        self.report_text.config(state='disabled')
        
        self.report_ax.clear()
        labels = ['درآمد', 'هزینه']
        sizes = [total_income, total_expense]
        colors = [self.SUCCESS_COLOR, self.ERROR_COLOR]
        explode = (0.1, 0) if total_income > total_expense else (0, 0.1)
        
        if any(s > 0 for s in sizes):
            self.report_ax.pie(sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%', shadow=True, startangle=90, textprops={'fontname': self.FONT_FAMILY[0], 'fontsize': 10})
        else:
            self.report_ax.text(0.5, 0.5, 'داده‌ای برای نمایش وجود ندارد', ha='center', va='center', fontdict={'fontname': self.FONT_FAMILY[0]})
        
        self.report_ax.axis('equal')
        self.report_fig.tight_layout()
        self.chart_canvas.draw()
        logging.info("گزارشات مالی به‌روزرسانی شد.")
        
    def _update_report_text_only(self):
        data = self.db_manager.get_financial_report_data()
        total_income = data.get('total_income', 0.0)
        total_expense = data.get('total_expense', 0.0)
        report_str = f"خلاصه وضعیت مالی (نمودار غیرفعال است)\n{'='*30}\n"
        report_str += f"مجموع درآمد کل: {total_income:,.0f} ریال\n"
        report_str += f"مجموع هزینه‌های کل: {total_expense:,.0f} ریال\n"
        self.report_text.config(state='normal')
        self.report_text.delete('1.0', tk.END)
        self.report_text.insert('1.0', report_str)
        self.report_text.config(state='disabled')

    def export_financial_report(self):
        if not (pd and openpyxl):
            messagebox.showerror("خطا", "کتابخانه‌های 'pandas' و 'openpyxl' برای این عملیات مورد نیاز هستند.\nلطفاً آنها را نصب کنید.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="ذخیره گزارش اکسل", initialfile=f"financial_report_{datetime.now().strftime('%Y-%m-%d')}.xlsx")
        if not file_path: return

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                income_sql = "SELECT s.name || ' ' || s.family_name as 'دانش‌آموز', sf.payment_date as 'تاریخ', sf.payment_amount as 'مبلغ', sf.payment_type as 'نوع', sf.description as 'توضیحات' FROM student_finance sf JOIN students s ON sf.student_id = s.id"
                income_df = pd.read_sql_query(income_sql, self.db_manager.conn)
                income_df.to_excel(writer, sheet_name='درآمدها', index=False)
                
                expense_sql = "SELECT date as 'تاریخ', description as 'شرح', amount as 'مبلغ', category as 'دسته‌بندی' FROM kindergarten_expenses"
                expense_df = pd.read_sql_query(expense_sql, self.db_manager.conn)
                expense_df.to_excel(writer, sheet_name='هزینه‌ها', index=False)
            
            messagebox.showinfo("موفقیت", f"گزارش با موفقیت در فایل زیر ذخیره شد:\n{file_path}")
            logging.info(f"گزارش اکسل در {file_path} ذخیره شد.")

        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ایجاد فایل اکسل: {e}")
            logging.error(f"خطا در خروجی اکسل: {e}")
            
    def _create_person_form(self, parent, labels_dict, date_keys):
        entries = {}
        # Reduce vertical padding to make form more compact
        for i, (key, text) in enumerate(labels_dict.items()):
            ttk.Label(parent, text=text).grid(row=i, column=0, padx=5, pady=5, sticky="w")
            if key == "status":
                widget = ttk.Combobox(parent, values=["فعال", "غیرفعال"], state="readonly")
                widget.set("فعال")
            elif key in date_keys:
                widget = self.create_date_entry(parent)
            else:
                widget = ttk.Entry(parent)
            widget.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
            entries[key] = widget
        return entries

    def _create_person_list_frame(self, parent, title, search_handler):
        list_frame = ttk.LabelFrame(parent, text=title, padding=15)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        list_frame.rowconfigure(1, weight=1)
        list_frame.columnconfigure(0, weight=1)

        search_frame = ttk.Frame(list_frame)
        search_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 10))
        search_frame.columnconfigure(1, weight=1)
        ttk.Label(search_frame, text="جستجو:").grid(row=0, column=0, padx=(0, 10))
        search_entry = ttk.Entry(search_frame)
        search_entry.grid(row=0, column=1, sticky='ew')
        search_entry.bind("<KeyRelease>", search_handler)

        return list_frame, search_entry

    def _create_treeview(self, parent, columns, headings, widths):
        tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        tree.grid(row=1, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=1, column=1, sticky='ns')

        for col, text in headings.items():
            tree.heading(col, text=text, anchor="center")
            tree.column(col, width=widths.get(col, 100), anchor="center")
        tree.column("id", stretch=tk.NO, width=40)
        return tree
    
    def create_student_tab(self):
        student_tab = ttk.Frame(self.content_frame, style='TFrame', padding=15)
        student_tab.columnconfigure(0, weight=3); student_tab.columnconfigure(1, weight=1); student_tab.rowconfigure(0, weight=1)
        list_frame, self.student_search_entry = self._create_person_list_frame(student_tab, "لیست دانش‌آموزان", self.search_students)
        columns = ("id", "نام", "نام خانوادگی", "کد ملی", "شماره تماس", "کلاس", "وضعیت")
        headings = {"id": "ID", "نام": "نام", "نام خانوادگی": "نام خانوادگی", "کد ملی": "کد ملی", "شماره تماس": "شماره تماس", "کلاس": "کلاس", "وضعیت": "وضعیت"}
        widths = {"نام": 120, "نام خانوادگی": 150, "کد ملی": 110, "شماره تماس": 120, "کلاس": 100, "وضعیت": 80}
        self.student_tree = self._create_treeview(list_frame, columns, headings, widths)
        self.student_tree.bind("<<TreeviewSelect>>", self.load_student_to_form)
        form_frame = ttk.LabelFrame(student_tab, text="فرم ثبت/ویرایش دانش‌آموز", padding=15)
        form_frame.grid(row=0, column=1, sticky="nsew"); form_frame.columnconfigure(1, weight=1)
        self.student_labels = { "name": "نام:", "family_name": "نام خانوادگی:", "national_id": "کد ملی:", "birth_date": "تاریخ تولد:", "entry_date": "تاریخ ثبت‌نام:", "father_name": "نام پدر:", "mother_name": "نام مادر:", "parent_phone": "شماره تماس والدین:", "address": "آدرس:", "class_name": "کلاس:", "status": "وضعیت:", "allergies": "آلرژی‌ها:", "notes": "توضیحات:" }
        self.student_entries = self._create_person_form(form_frame, self.student_labels, ["birth_date", "entry_date"])
        
        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=len(self.student_labels), column=0, columnspan=2, pady=(15, 5), sticky="ew")
        button_frame.columnconfigure((0, 1), weight=1)
        btn_config = {'padx': 5, 'pady': 5, 'sticky': 'ew'}
        
        ttk.Button(button_frame, text="ذخیره", command=self.save_student).grid(row=0, column=0, **btn_config)
        ttk.Button(button_frame, text="ویرایش", command=self.update_student).grid(row=0, column=1, **btn_config)
        ttk.Button(button_frame, text="حذف", command=self.delete_student).grid(row=1, column=0, **btn_config)
        ttk.Button(button_frame, text="پاک کردن", command=self.clear_student_form).grid(row=1, column=1, **btn_config)

        self.clear_student_form()
        return student_tab
        
    def clear_student_form(self):
        self.student_tree.selection_set(())
        for key, widget in self.student_entries.items():
            if isinstance(widget, ttk.Combobox): widget.set("فعال")
            elif isinstance(widget, DateEntry): widget.set_date(datetime.now())
            else: widget.delete(0, tk.END)

    def load_students(self, query=None):
        search_query = f"%{query.lower()}%" if query else None
        if search_query:
            sql = "SELECT id, name, family_name, national_id, parent_phone, class_name, status FROM students WHERE lower(name) LIKE ? OR lower(family_name) LIKE ? OR national_id LIKE ? ORDER BY family_name, name"
            params = (search_query, search_query, search_query)
        else:
            sql = "SELECT id, name, family_name, national_id, parent_phone, class_name, status FROM students ORDER BY family_name, name"
            params = ()
        results = self.db_manager.execute_query(sql, params, fetch="all")
        if results is not None: 
            students = [(s[0], s[1], s[2], s[3], s[4], s[5], "فعال" if s[6] == self.STATUS_ACTIVE else "غیرفعال") for s in results]
            self._add_items_to_treeview(self.student_tree, students)

    def search_students(self, event=None): 
        self.load_students(self.student_search_entry.get().strip())

    def load_student_to_form(self, event=None):
        selected_items = self.student_tree.selection()
        if not selected_items: return
        student_id = self.student_tree.item(selected_items[0])['values'][0]
        student_data = self.db_manager.execute_query("SELECT * FROM students WHERE id=?", (student_id,), fetch="one")
        if not student_data: return
        col_keys = list(self.student_labels.keys())
        data_map = {col_keys[i]: student_data[i+1] for i in range(len(col_keys))}
        for key, widget in self.student_entries.items():
            value = data_map.get(key, "")
            if isinstance(widget, DateEntry):
                try:
                    if value: widget.set_date(value)
                    else: widget.set_date(datetime.now())
                except (ValueError, TypeError): widget.set_date(datetime.now())
            elif isinstance(widget, ttk.Combobox): widget.set("فعال" if value == self.STATUS_ACTIVE else "غیرفعال")
            else:
                widget.delete(0, tk.END)
                if value is not None: widget.insert(0, value)

    def get_student_form_data(self):
        data = {key: widget.get().strip() for key, widget in self.student_entries.items()}
        data['status'] = self.STATUS_ACTIVE if self.student_entries['status'].get() == "فعال" else self.STATUS_INACTIVE
        if not all([data['name'], data['family_name'], data['national_id'], data['entry_date']]):
            messagebox.showerror("خطای اعتبارسنجی", "فیلدهای نام، نام خانوادگی، کد ملی و تاریخ ثبت‌نام الزامی هستند.")
            return None
        if not (data['national_id'].isdigit() and len(data['national_id']) == 10):
            messagebox.showerror("خطای اعتبارسنجی", "کد ملی باید یک عدد ۱۰ رقمی معتبر باشد.")
            return None
        if data['parent_phone'] and not re.match(r'^[\d\s\-()]+$', data['parent_phone']):
            messagebox.showerror("خطای اعتبارسنجی", "فرمت شماره تماس والدین معتبر نیست.")
            return None
        return tuple(data.values())

    def save_student(self):
        data = self.get_student_form_data()
        if not data: return
        sql = "INSERT INTO students (name, family_name, national_id, birth_date, entry_date, father_name, mother_name, parent_phone, address, class_name, status, allergies, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        try:
            if self.db_manager.execute_query(sql, data):
                self.db_manager.commit()
                messagebox.showinfo("موفقیت", "دانش‌آموز با موفقیت ذخیره شد.")
                logging.info(f"دانش‌آموز جدید ذخیره شد: {data[0]} {data[1]}")
                self.clear_student_form(); self.load_students(); self.update_all_student_comboboxes()
            else: self.db_manager.rollback()
        except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در ذخیره دانش‌آموز: {e}")

    def update_student(self):
        selected_items = self.student_tree.selection()
        if not selected_items: messagebox.showwarning("انتخاب نشده", "لطفا یک دانش‌آموز را برای ویرایش انتخاب کنید."); return
        student_id = self.student_tree.item(selected_items[0])['values'][0]
        data = self.get_student_form_data()
        if not data: return
        sql = "UPDATE students SET name=?, family_name=?, national_id=?, birth_date=?, entry_date=?, father_name=?, mother_name=?, parent_phone=?, address=?, class_name=?, status=?, allergies=?, notes=? WHERE id=?"
        try:
            if self.db_manager.execute_query(sql, (*data, student_id)):
                self.db_manager.commit()
                messagebox.showinfo("موفقیت", "اطلاعات دانش‌آموز به‌روزرسانی شد.")
                logging.info(f"دانش‌آموز {student_id} به‌روزرسانی شد")
                self.clear_student_form(); self.load_students(); self.update_all_student_comboboxes()
            else: self.db_manager.rollback()
        except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در ویرایش دانش‌آموز: {e}")

    def delete_student(self):
        selected_items = self.student_tree.selection()
        if not selected_items: messagebox.showwarning("انتخاب نشده", "لطفا یک دانش‌آموز را برای حذف انتخاب کنید."); return
        student_id = self.student_tree.item(selected_items[0])['values'][0]
        student_name = f"{self.student_tree.item(selected_items[0])['values'][1]} {self.student_tree.item(selected_items[0])['values'][2]}"
        if messagebox.askyesno("تأیید حذف", f"آیا از حذف دانش‌آموز '{student_name}' مطمئن هستید؟\n(تمام اطلاعات مالی و ارتباطی این دانش‌آموز نیز حذف خواهد شد)"):
            try:
                if self.db_manager.execute_query("DELETE FROM students WHERE id = ?", (student_id,)):
                    self.db_manager.commit()
                    messagebox.showinfo("موفقیت", "دانش‌آموز با موفقیت حذف شد.")
                    logging.info(f"دانش‌آموز {student_id} حذف شد")
                    self.clear_student_form(); self.load_students(); self.update_all_student_comboboxes()
                else: self.db_manager.rollback()
            except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در حذف دانش‌آموز: {e}")

    def update_all_student_comboboxes(self):
        try:
            results = self.db_manager.execute_query("SELECT id, name, family_name FROM students WHERE status = 'active' ORDER BY family_name, name", fetch="all")
            if results is None: self.student_map, student_names = {}, []
            else: self.student_map, student_names = {f"{row[1]} {row[2]}": row[0] for row in results}, list(self.student_map.keys())
            if hasattr(self, 'finance_student_combo'): self.finance_student_combo['values'] = student_names; self.finance_student_combo.set('')
            if hasattr(self, 'comm_student_combo'): self.comm_student_combo['values'] = student_names; self.comm_student_combo.set('')
        except Exception as e: logging.error(f"خطا در به‌روزرسانی لیست دانش‌آموزان در کامبوباکس‌ها: {e}")

    def create_staff_tab(self):
        staff_tab = ttk.Frame(self.content_frame, style='TFrame', padding=15)
        staff_tab.columnconfigure(0, weight=3); staff_tab.columnconfigure(1, weight=1); staff_tab.rowconfigure(0, weight=1)
        list_frame, self.staff_search_entry = self._create_person_list_frame(staff_tab, "لیست پرسنل", self.search_staff)
        columns = ("id", "نام", "نام خانوادگی", "کد ملی", "سمت", "حقوق", "وضعیت")
        headings = {"id": "ID", "نام": "نام", "نام خانوادگی": "نام خانوادگی", "کد ملی": "کد ملی", "سمت": "سمت", "حقوق": "حقوق (ریال)", "وضعیت": "وضعیت"}
        widths = {"نام": 120, "نام خانوادگی": 150, "کد ملی": 110, "سمت": 120, "حقوق": 120, "وضعیت": 80}
        self.staff_tree = self._create_treeview(list_frame, columns, headings, widths)
        self.staff_tree.bind("<<TreeviewSelect>>", self.load_staff_to_form)
        form_frame = ttk.LabelFrame(staff_tab, text="فرم ثبت/ویرایش پرسنل", padding=15)
        form_frame.grid(row=0, column=1, sticky="nsew"); form_frame.columnconfigure(1, weight=1)
        self.staff_labels = { "name": "نام:", "family_name": "نام خانوادگی:", "national_id": "کد ملی:", "birth_date": "تاریخ تولد:", "hire_date": "تاریخ استخدام:", "phone": "شماره تماس:", "address": "آدرس:", "role": "سمت:", "salary": "حقوق (ریال):", "status": "وضعیت:", "notes": "توضیحات:" }
        self.staff_entries = self._create_person_form(form_frame, self.staff_labels, ["birth_date", "hire_date"])

        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=len(self.staff_labels), column=0, columnspan=2, pady=(15, 5), sticky="ew")
        button_frame.columnconfigure((0, 1), weight=1)
        btn_config = {'padx': 5, 'pady': 5, 'sticky': 'ew'}
        
        ttk.Button(button_frame, text="ذخیره", command=self.save_staff).grid(row=0, column=0, **btn_config)
        ttk.Button(button_frame, text="ویرایش", command=self.update_staff).grid(row=0, column=1, **btn_config)
        ttk.Button(button_frame, text="حذف", command=self.delete_staff).grid(row=1, column=0, **btn_config)
        ttk.Button(button_frame, text="پاک کردن", command=self.clear_staff_form).grid(row=1, column=1, **btn_config)

        self.clear_staff_form()
        return staff_tab
        
    def clear_staff_form(self):
        self.staff_tree.selection_set(())
        for key, widget in self.staff_entries.items():
            if isinstance(widget, ttk.Combobox): widget.set("فعال")
            elif isinstance(widget, DateEntry): widget.set_date(datetime.now())
            else: widget.delete(0, tk.END)

    def load_staff(self, query=None):
        search_query = f"%{query.lower()}%" if query else None
        if search_query:
            sql, params = "SELECT id, name, family_name, national_id, role, salary, status FROM staff WHERE lower(name) LIKE ? OR lower(family_name) LIKE ? OR national_id LIKE ? ORDER BY family_name, name", (search_query, search_query, search_query)
        else:
            sql, params = "SELECT id, name, family_name, national_id, role, salary, status FROM staff ORDER BY family_name, name", ()
        results = self.db_manager.execute_query(sql, params, fetch="all")
        if results is not None:
            staff_list = [ (s[0], s[1], s[2], s[3], s[4], f"{s[5]:,.0f}" if s[5] is not None else "0", "فعال" if s[6] == self.STATUS_ACTIVE else "غیرفعال") for s in results ]
            self._add_items_to_treeview(self.staff_tree, staff_list)

    def search_staff(self, event=None): self.load_staff(self.staff_search_entry.get().strip())

    def load_staff_to_form(self, event=None):
        selected_items = self.staff_tree.selection()
        if not selected_items: return
        staff_id = self.staff_tree.item(selected_items[0])['values'][0]
        staff_data = self.db_manager.execute_query("SELECT * FROM staff WHERE id=?", (staff_id,), fetch="one")
        if not staff_data: return
        col_keys = list(self.staff_labels.keys())
        data_map = {col_keys[i]: staff_data[i+1] for i in range(len(col_keys))}
        for key, widget in self.staff_entries.items():
            value = data_map.get(key, "")
            if isinstance(widget, DateEntry):
                try:
                    if value: widget.set_date(value)
                    else: widget.set_date(datetime.now())
                except (ValueError, TypeError): widget.set_date(datetime.now())
            elif isinstance(widget, ttk.Combobox): widget.set("فعال" if value == self.STATUS_ACTIVE else "غیرفعال")
            else:
                widget.delete(0, tk.END)
                if value is not None: widget.insert(0, str(value))

    def get_staff_form_data(self):
        data = {key: widget.get().strip() for key, widget in self.staff_entries.items()}
        data['status'] = self.STATUS_ACTIVE if data['status'] == "فعال" else self.STATUS_INACTIVE
        if not all([data['name'], data['family_name'], data['national_id'], data['hire_date'], data['role']]):
            messagebox.showerror("خطای اعتبارسنجی", "فیلدهای نام، نام خانوادگی، کد ملی، تاریخ استخدام و سمت الزامی هستند."); return None
        if not (data['national_id'].isdigit() and len(data['national_id']) == 10):
            messagebox.showerror("خطای اعتبارسنجی", "کد ملی باید یک عدد ۱۰ رقمی معتبر باشد."); return None
        try:
            data['salary'] = float(data['salary'].replace(',', '')) if data['salary'] else 0.0
        except ValueError: messagebox.showerror("خطای اعتبارسنجی", "مقدار حقوق باید یک عدد معتبر باشد."); return None
        return (data['name'], data['family_name'], data['national_id'], data['birth_date'], data['hire_date'], data['phone'], data['address'], data['role'], data['salary'], data['status'], data['notes'])

    def save_staff(self):
        data = self.get_staff_form_data()
        if not data: return
        sql = "INSERT INTO staff (name, family_name, national_id, birth_date, hire_date, phone, address, role, salary, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        try:
            if self.db_manager.execute_query(sql, data):
                self.db_manager.commit()
                messagebox.showinfo("موفقیت", "پرسنل با موفقیت ذخیره شد."); logging.info(f"پرسنل جدید ذخیره شد: {data[0]} {data[1]}"); self.clear_staff_form(); self.load_staff()
            else: self.db_manager.rollback()
        except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در ذخیره پرسنل: {e}")

    def update_staff(self):
        selected_items = self.staff_tree.selection()
        if not selected_items: messagebox.showwarning("انتخاب نشده", "لطفا یک شخص را برای ویرایش انتخاب کنید."); return
        staff_id = self.staff_tree.item(selected_items[0])['values'][0]
        data = self.get_staff_form_data()
        if not data: return
        sql = "UPDATE staff SET name=?, family_name=?, national_id=?, birth_date=?, hire_date=?, phone=?, address=?, role=?, salary=?, status=?, notes=? WHERE id=?"
        try:
            if self.db_manager.execute_query(sql, (*data, staff_id)):
                self.db_manager.commit()
                messagebox.showinfo("موفقیت", "اطلاعات پرسنل به‌روزرسانی شد."); logging.info(f"پرسنل {staff_id} به‌روزرسانی شد"); self.clear_staff_form(); self.load_staff()
            else: self.db_manager.rollback()
        except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در ویرایش پرسنل: {e}")

    def delete_staff(self):
        selected_items = self.staff_tree.selection()
        if not selected_items: messagebox.showwarning("انتخاب نشده", "لطفا یک شخص را برای حذف انتخاب کنید."); return
        staff_id = self.staff_tree.item(selected_items[0])['values'][0]
        staff_name = f"{self.staff_tree.item(selected_items[0])['values'][1]} {self.staff_tree.item(selected_items[0])['values'][2]}"
        if messagebox.askyesno("تأیید حذف", f"آیا از حذف '{staff_name}' مطمئن هستید؟"):
            try:
                if self.db_manager.execute_query("DELETE FROM staff WHERE id = ?", (staff_id,)):
                    self.db_manager.commit()
                    messagebox.showinfo("موفقیت", "پرسنل با موفقیت حذف شد."); logging.info(f"پرسنل {staff_id} حذف شد"); self.clear_staff_form(); self.load_staff()
                else: self.db_manager.rollback()
            except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در حذف پرسنل: {e}")

    def create_attendance_tab(self):
        attendance_tab = ttk.Frame(self.content_frame, style='TFrame', padding=15); attendance_tab.rowconfigure(1, weight=1); attendance_tab.columnconfigure(0, weight=1)
        top_frame = ttk.Frame(attendance_tab); top_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15)); top_frame.columnconfigure(2, weight=1)
        ttk.Label(top_frame, text="تاریخ:").grid(row=0, column=0, padx=5, pady=5)
        self.attendance_date_entry = self.create_date_entry(top_frame); self.attendance_date_entry.grid(row=0, column=1, padx=5, pady=5); self.attendance_date_entry.bind("<<DateEntrySelected>>", self.load_attendance_list)
        ttk.Button(top_frame, text="ذخیره حضور و غیاب", command=self.save_attendance).grid(row=0, column=3, padx=10, pady=5, sticky='e')
        notebook = ttk.Notebook(attendance_tab); notebook.grid(row=1, column=0, sticky="nsew")
        self.student_attendance_tree = self._create_attendance_person_frame(notebook, "دانش‌آموزان", self.PERSON_TYPE_STUDENT)
        self.staff_attendance_tree = self._create_attendance_person_frame(notebook, "پرسنل", self.PERSON_TYPE_STAFF)
        return attendance_tab

    def _create_attendance_person_frame(self, notebook, text, person_type):
        frame = ttk.Frame(notebook, padding=10); notebook.add(frame, text=text); frame.columnconfigure(0, weight=1); frame.rowconfigure(1, weight=1)
        controls_frame = ttk.Frame(frame); controls_frame.grid(row=0, column=0, sticky='ew', pady=(0, 10))
        ttk.Button(controls_frame, text="حاضر", command=lambda: self._set_attendance_status(person_type, self.ATTENDANCE_PRESENT)).pack(side='right', padx=2)
        ttk.Button(controls_frame, text="غایب", command=lambda: self._set_attendance_status(person_type, self.ATTENDANCE_ABSENT)).pack(side='right', padx=2)
        ttk.Button(controls_frame, text="مرخصی", command=lambda: self._set_attendance_status(person_type, self.ATTENDANCE_LEAVE)).pack(side='right', padx=2)
        columns = ("id", "نام", "نام خانوادگی", "وضعیت", "ساعت ورود", "ساعت خروج")
        headings = {"id": "ID", "نام": "نام", "نام خانوادگی": "نام خانوادگی", "وضعیت": "وضعیت", "ساعت ورود": "ورود", "ساعت خروج": "خروج"}
        widths = {"id": 40, "نام": 150, "نام خانوادگی": 180, "وضعیت": 100, "ساعت ورود": 100, "ساعت خروج": 100}
        tree = self._create_treeview(frame, columns, headings, widths); tree.grid(row=1, column=0, sticky='nsew')
        tree.bind("<Double-1>", lambda event, pt=person_type: self._edit_time_popup(event, pt))
        return tree

    def load_attendance_list(self, event=None):
        selected_date = self.attendance_date_entry.get()
        self._load_person_attendance(self.PERSON_TYPE_STUDENT, selected_date, self.student_attendance_tree)
        self._load_person_attendance(self.PERSON_TYPE_STAFF, selected_date, self.staff_attendance_tree)

    def _load_person_attendance(self, person_type, date, tree):
        table_name = "students" if person_type == self.PERSON_TYPE_STUDENT else "staff"
        sql = f"SELECT p.id, p.name, p.family_name, a.status, a.entry_time, a.exit_time FROM {table_name} p LEFT JOIN attendance a ON p.id = a.person_id AND a.person_type = ? AND a.date = ? WHERE p.status = ? ORDER BY p.family_name, p.name"
        results = self.db_manager.execute_query(sql, (person_type, date, self.STATUS_ACTIVE), fetch="all")
        person_list = [ (p[0], p[1], p[2], p[3] or self.ATTENDANCE_ABSENT, p[4] or "", p[5] or "") for p in results ] if results else []
        self._add_items_to_treeview(tree, person_list)

    def _set_attendance_status(self, person_type, status):
        tree = self.student_attendance_tree if person_type == self.PERSON_TYPE_STUDENT else self.staff_attendance_tree
        selected_items = tree.selection()
        if not selected_items: messagebox.showwarning("انتخاب نشده", "لطفا حداقل یک نفر را از لیست انتخاب کنید."); return
        for item_id in selected_items:
            current_values = list(tree.item(item_id, 'values'))
            current_values[3] = status
            if status == self.ATTENDANCE_PRESENT and not current_values[4]: current_values[4] = datetime.now().strftime('%H:%M')
            tree.item(item_id, values=tuple(current_values))

    def _edit_time_popup(self, event, person_type):
        tree = self.student_attendance_tree if person_type == self.PERSON_TYPE_STUDENT else self.staff_attendance_tree
        item_id = tree.identify_row(event.y)
        if not item_id: return
        col_index = int(tree.identify_column(event.x).replace('#', '')) - 1
        if col_index not in [4, 5]: return
        current_values = list(tree.item(item_id, 'values'))
        popup = tk.Toplevel(self); popup.transient(self); popup.grab_set(); popup.title("ویرایش زمان")
        label_text = "ساعت ورود (HH:MM):" if col_index == 4 else "ساعت خروج (HH:MM):"
        ttk.Label(popup, text=label_text).pack(padx=10, pady=5)
        entry = ttk.Entry(popup); entry.pack(padx=10, pady=5); entry.insert(0, current_values[col_index]); entry.focus()
        def on_ok():
            new_time = entry.get().strip()
            if re.match(r'^\d{2}:\d{2}$', new_time): current_values[col_index] = new_time; tree.item(item_id, values=tuple(current_values)); popup.destroy()
            else: messagebox.showerror("فرمت نامعتبر", "لطفا زمان را در فرمت HH:MM وارد کنید.", parent=popup)
        ttk.Button(popup, text="تایید", command=on_ok).pack(pady=10)
        entry.bind("<Return>", lambda e: on_ok())

    def save_attendance(self):
        date = self.attendance_date_entry.get()
        attendance_data = [ (*tree.item(item_id, 'values')[:1], person_type, date, *tree.item(item_id, 'values')[3:]) for tree, person_type in [(self.student_attendance_tree, self.PERSON_TYPE_STUDENT), (self.staff_attendance_tree, self.PERSON_TYPE_STAFF)] for item_id in tree.get_children() ]
        sql = "INSERT INTO attendance (person_id, person_type, date, status, entry_time, exit_time) VALUES (?, ?, ?, ?, ?, ?) ON CONFLICT(person_id, person_type, date) DO UPDATE SET status=excluded.status, entry_time=excluded.entry_time, exit_time=excluded.exit_time"
        try:
            for record in attendance_data: self.db_manager.execute_query(sql, record)
            self.db_manager.commit()
            messagebox.showinfo("موفقیت", "حضور و غیاب برای تاریخ انتخاب شده با موفقیت ذخیره شد."); logging.info(f"حضور و غیاب برای تاریخ {date} ذخیره شد.")
        except Exception as e: self.db_manager.rollback(); messagebox.showerror("خطای دیتابیس", f"خطا در ذخیره اطلاعات حضور و غیاب: {e}"); logging.error(f"خطا در ذخیره حضور و غیاب: {e}")

    def create_finance_tab(self):
        finance_tab = ttk.Frame(self.content_frame, style='TFrame', padding=15)
        notebook = ttk.Notebook(finance_tab)
        notebook.pack(expand=True, fill="both")
        notebook.add(self.create_student_finance_frame(notebook), text=f"{self.icons['finance']} درآمدها (شهریه)")
        notebook.add(self.create_expense_frame(notebook), text=f"{self.icons['expense']} هزینه‌ها")
        return finance_tab

    def create_student_finance_frame(self, parent):
        frame = ttk.Frame(parent, style='TFrame', padding=15); frame.columnconfigure(0, weight=3); frame.columnconfigure(1, weight=1); frame.rowconfigure(0, weight=1)
        list_frame = ttk.LabelFrame(frame, text="لیست پرداخت‌های دانش‌آموزان", padding=15); list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10)); list_frame.rowconfigure(0, weight=1); list_frame.columnconfigure(0, weight=1)
        self.student_finance_tree = self._create_treeview(list_frame, ("id", "دانش‌آموز", "تاریخ پرداخت", "مبلغ", "نوع پرداخت", "توضیحات"), {"id": "ID", "دانش‌آموز": "دانش‌آموز", "تاریخ پرداخت": "تاریخ", "مبلغ": "مبلغ (ریال)", "نوع پرداخت": "نوع", "توضیحات": "توضیحات"}, {})
        form_frame = ttk.LabelFrame(frame, text="ثبت پرداخت جدید", padding=15); form_frame.grid(row=0, column=1, sticky="nsew"); form_frame.columnconfigure(1, weight=1)
        ttk.Label(form_frame, text="دانش‌آموز:").grid(row=0, column=0, sticky='w', padx=5, pady=5); self.finance_student_combo = ttk.Combobox(form_frame, state='readonly'); self.finance_student_combo.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(form_frame, text="تاریخ پرداخت:").grid(row=1, column=0, sticky='w', padx=5, pady=5); self.finance_date_entry = self.create_date_entry(form_frame); self.finance_date_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(form_frame, text="مبلغ (ریال):").grid(row=2, column=0, sticky='w', padx=5, pady=5); self.finance_amount_entry = ttk.Entry(form_frame); self.finance_amount_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(form_frame, text="نوع پرداخت:").grid(row=3, column=0, sticky='w', padx=5, pady=5); self.finance_type_entry = ttk.Combobox(form_frame, values=["شهریه", "کلاس فوق‌العاده", "سرویس", "غیره"]); self.finance_type_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(form_frame, text="توضیحات:").grid(row=4, column=0, sticky='w', padx=5, pady=5); self.finance_desc_entry = ttk.Entry(form_frame); self.finance_desc_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=5)
        btn_frame = ttk.Frame(form_frame); btn_frame.grid(row=5, column=0, columnspan=2, pady=15, sticky='ew'); btn_frame.columnconfigure((0,1), weight=1)
        ttk.Button(btn_frame, text="ذخیره", command=self.save_student_finance).pack(side='right', fill='x', padx=2)
        ttk.Button(btn_frame, text="حذف", command=self.delete_student_finance).pack(side='left', fill='x', padx=2)
        return frame
        
    def load_student_finances(self):
        sql = "SELECT sf.id, s.name, s.family_name, sf.payment_date, sf.payment_amount, sf.payment_type, sf.description FROM student_finance sf JOIN students s ON sf.student_id = s.id ORDER BY sf.payment_date DESC"
        results = self.db_manager.execute_query(sql, fetch="all")
        finance_list = [ (r[0], f"{r[1]} {r[2]}", r[3], f"{r[4]:,.0f}", r[5], r[6]) for r in results ] if results else []
        self._add_items_to_treeview(self.student_finance_tree, finance_list)

    def save_student_finance(self):
        student_id = self.student_map.get(self.finance_student_combo.get())
        date, amount_str, ptype, desc = self.finance_date_entry.get(), self.finance_amount_entry.get().replace(',', ''), self.finance_type_entry.get(), self.finance_desc_entry.get().strip()
        if not all([student_id, date, amount_str, ptype]): messagebox.showerror("خطا", "لطفا تمام فیلدهای ستاره‌دار را پر کنید."); return
        try: amount = float(amount_str)
        except ValueError: messagebox.showerror("خطا", "مبلغ باید یک عدد معتبر باشد."); return
        sql = "INSERT INTO student_finance (student_id, payment_date, payment_amount, payment_type, description) VALUES (?, ?, ?, ?, ?)"
        try:
            if self.db_manager.execute_query(sql, (student_id, date, amount, ptype, desc)):
                self.db_manager.commit(); messagebox.showinfo("موفقیت", "پرداخت با موفقیت ثبت شد."); self.load_student_finances()
                self.finance_student_combo.set(''); self.finance_amount_entry.delete(0, tk.END); self.finance_type_entry.set(''); self.finance_desc_entry.delete(0, tk.END)
            else: self.db_manager.rollback()
        except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در ثبت پرداخت: {e}")

    def delete_student_finance(self):
        selected_items = self.student_finance_tree.selection()
        if not selected_items: messagebox.showwarning("انتخاب نشده", "لطفا یک رکورد پرداخت را برای حذف انتخاب کنید."); return
        finance_id = self.student_finance_tree.item(selected_items[0])['values'][0]
        if messagebox.askyesno("تایید حذف", "آیا از حذف این رکورد پرداخت مطمئن هستید؟"):
            try:
                if self.db_manager.execute_query("DELETE FROM student_finance WHERE id=?", (finance_id,)):
                    self.db_manager.commit(); messagebox.showinfo("موفقیت", "رکورد پرداخت حذف شد."); self.load_student_finances()
                else: self.db_manager.rollback()
            except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در حذف رکورد پرداخت: {e}")

    def create_expense_frame(self, parent):
        frame = ttk.Frame(parent, style='TFrame', padding=15); frame.columnconfigure(0, weight=3); frame.columnconfigure(1, weight=1); frame.rowconfigure(0, weight=1)
        list_frame = ttk.LabelFrame(frame, text="لیست هزینه‌ها", padding=15); list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10)); list_frame.rowconfigure(0, weight=1); list_frame.columnconfigure(0, weight=1)
        self.expense_tree = self._create_treeview(list_frame, ("id", "تاریخ", "شرح هزینه", "مبلغ", "دسته‌بندی"), {"id": "ID", "تاریخ": "تاریخ", "شرح هزینه": "شرح", "مبلغ": "مبلغ (ریال)", "دسته‌بندی": "دسته‌بندی"}, {})
        form_frame = ttk.LabelFrame(frame, text="ثبت هزینه جدید", padding=15); form_frame.grid(row=0, column=1, sticky="nsew"); form_frame.columnconfigure(1, weight=1)
        ttk.Label(form_frame, text="تاریخ:").grid(row=0, column=0, sticky='w', padx=5, pady=5); self.expense_date_entry = self.create_date_entry(form_frame); self.expense_date_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(form_frame, text="شرح هزینه:").grid(row=1, column=0, sticky='w', padx=5, pady=5); self.expense_desc_entry = ttk.Entry(form_frame); self.expense_desc_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(form_frame, text="مبلغ (ریال):").grid(row=2, column=0, sticky='w', padx=5, pady=5); self.expense_amount_entry = ttk.Entry(form_frame); self.expense_amount_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(form_frame, text="دسته‌بندی:").grid(row=3, column=0, sticky='w', padx=5, pady=5); self.expense_category_entry = ttk.Combobox(form_frame, values=["حقوق پرسنل", "اجاره", "قبوض", "مواد غذایی", "لوازم التحریر", "تعمیرات", "غیره"]); self.expense_category_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=5)
        btn_frame = ttk.Frame(form_frame); btn_frame.grid(row=4, column=0, columnspan=2, pady=15, sticky='ew'); btn_frame.columnconfigure((0,1), weight=1)
        ttk.Button(btn_frame, text="ذخیره", command=self.save_expense).pack(side='right', fill='x', padx=2)
        ttk.Button(btn_frame, text="حذف", command=self.delete_expense).pack(side='left', fill='x', padx=2)
        return frame

    def load_expenses(self):
        sql = "SELECT id, date, description, amount, category FROM kindergarten_expenses ORDER BY date DESC"
        results = self.db_manager.execute_query(sql, fetch="all")
        expense_list = [ (r[0], r[1], r[2], f"{r[3]:,.0f}", r[4]) for r in results ] if results else []
        self._add_items_to_treeview(self.expense_tree, expense_list)

    def save_expense(self):
        date, desc, amount_str, category = self.expense_date_entry.get(), self.expense_desc_entry.get().strip(), self.expense_amount_entry.get().replace(',', ''), self.expense_category_entry.get()
        if not all([date, desc, amount_str, category]): messagebox.showerror("خطا", "لطفا تمام فیلدها را پر کنید."); return
        try: amount = float(amount_str)
        except ValueError: messagebox.showerror("خطا", "مبلغ باید یک عدد معتبر باشد."); return
        sql = "INSERT INTO kindergarten_expenses (date, description, amount, category) VALUES (?, ?, ?, ?)"
        try:
            if self.db_manager.execute_query(sql, (date, desc, amount, category)):
                self.db_manager.commit(); messagebox.showinfo("موفقیت", "هزینه با موفقیت ثبت شد."); self.load_expenses()
                self.expense_desc_entry.delete(0, tk.END); self.expense_amount_entry.delete(0, tk.END); self.expense_category_entry.set('')
            else: self.db_manager.rollback()
        except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در ثبت هزینه: {e}")
            
    def delete_expense(self):
        selected_items = self.expense_tree.selection()
        if not selected_items: messagebox.showwarning("انتخاب نشده", "لطفا یک هزینه را برای حذف انتخاب کنید."); return
        expense_id = self.expense_tree.item(selected_items[0])['values'][0]
        if messagebox.askyesno("تایید حذف", "آیا از حذف این هزینه مطمئن هستید؟"):
            try:
                if self.db_manager.execute_query("DELETE FROM kindergarten_expenses WHERE id=?", (expense_id,)):
                    self.db_manager.commit(); messagebox.showinfo("موفقیت", "هزینه حذف شد."); self.load_expenses()
                else: self.db_manager.rollback()
            except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در حذف هزینه: {e}")

    def create_communication_tab(self):
        comm_tab = ttk.Frame(self.content_frame, style='TFrame', padding=15); comm_tab.columnconfigure(0, weight=1); comm_tab.rowconfigure(1, weight=1)
        form_frame = ttk.LabelFrame(comm_tab, text="ثبت ارتباط جدید", padding=15); form_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10)); form_frame.columnconfigure(1, weight=1)
        ttk.Label(form_frame, text="دانش‌آموز:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.comm_student_combo = ttk.Combobox(form_frame, state='readonly'); self.comm_student_combo.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky='ew'); self.comm_student_combo.bind("<<ComboboxSelected>>", self.load_communications)
        ttk.Label(form_frame, text="تاریخ:").grid(row=1, column=0, padx=5, pady=5, sticky='w'); self.comm_date_entry = self.create_date_entry(form_frame); self.comm_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        ttk.Label(form_frame, text="موضوع:").grid(row=1, column=2, padx=5, pady=5, sticky='w'); self.comm_subject_entry = ttk.Entry(form_frame); self.comm_subject_entry.grid(row=1, column=3, padx=5, pady=5, sticky='ew')
        ttk.Label(form_frame, text="متن پیام:").grid(row=2, column=0, padx=5, pady=5, sticky='nw'); self.comm_message_text = tk.Text(form_frame, height=5, font=self.TEXT_WIDGET_FONT, relief='solid', borderwidth=1, highlightthickness=1); self.comm_message_text.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
        ttk.Button(form_frame, text="ذخیره", command=self.save_communication).grid(row=3, column=3, padx=5, pady=10, sticky='e')
        list_frame = ttk.LabelFrame(comm_tab, text="تاریخچه ارتباطات", padding=15); list_frame.grid(row=1, column=0, sticky="nsew"); list_frame.rowconfigure(0, weight=1); list_frame.columnconfigure(0, weight=1)
        self.comm_tree = self._create_treeview(list_frame, ("id", "تاریخ", "موضوع", "متن پیام"), {"id": "ID", "تاریخ": "تاریخ", "موضوع": "موضوع", "متن پیام": "متن پیام"}, {"تاریخ": 100, "موضوع": 200})
        return comm_tab
        
    def load_communications(self, event=None):
        student_id = self.student_map.get(self.comm_student_combo.get())
        if not student_id: self._add_items_to_treeview(self.comm_tree, []); return
        sql = "SELECT id, date, subject, message FROM student_communications WHERE student_id=? ORDER BY date DESC"
        results = self.db_manager.execute_query(sql, (student_id,), fetch="all")
        self._add_items_to_treeview(self.comm_tree, results if results else [])

    def save_communication(self):
        student_id, date, subject, message = self.student_map.get(self.comm_student_combo.get()), self.comm_date_entry.get(), self.comm_subject_entry.get().strip(), self.comm_message_text.get("1.0", tk.END).strip()
        if not all([student_id, date, subject, message]): messagebox.showerror("خطا", "لطفا دانش‌آموز را انتخاب کرده و تمام فیلدها را پر کنید."); return
        sql = "INSERT INTO student_communications (student_id, date, subject, message) VALUES (?, ?, ?, ?)"
        try:
            if self.db_manager.execute_query(sql, (student_id, date, subject, message)):
                self.db_manager.commit(); messagebox.showinfo("موفقیت", "ارتباط با موفقیت ثبت شد."); self.load_communications()
                self.comm_subject_entry.delete(0, tk.END); self.comm_message_text.delete("1.0", tk.END)
            else: self.db_manager.rollback()
        except Exception as e: self.db_manager.rollback(); logging.error(f"خطا در ثبت ارتباط: {e}")
        
    def create_settings_tab(self):
        settings_tab = ttk.Frame(self.content_frame, style='TFrame', padding=25)
        general_frame = ttk.LabelFrame(settings_tab, text="تنظیمات عمومی", padding=20); general_frame.pack(fill="x", pady=10, anchor='n'); general_frame.columnconfigure(1, weight=1)
        self.settings_entries = {}
        settings_labels = {"kindergarten_name": "نام مهدکودک:", "kindergarten_address": "آدرس:", "kindergarten_phone": "تلفن:"}
        for i, (key, text) in enumerate(settings_labels.items()):
            ttk.Label(general_frame, text=text).grid(row=i, column=0, padx=10, pady=10, sticky='w')
            entry = ttk.Entry(general_frame); entry.grid(row=i, column=1, padx=10, pady=10, sticky='ew'); entry.insert(0, self.db_manager.get_setting(key)); self.settings_entries[key] = entry
        ttk.Button(general_frame, text="ذخیره تنظیمات", command=self.save_settings).grid(row=len(settings_labels), column=0, columnspan=2, pady=20)
        backup_frame = ttk.LabelFrame(settings_tab, text="پشتیبان‌گیری و بازیابی", padding=20); backup_frame.pack(fill="x", pady=20, anchor='n')
        self.backup_status_label = ttk.Label(backup_frame, text="", font=self.BODY_FONT); self.backup_status_label.pack(pady=5)
        btn_frame = ttk.Frame(backup_frame); btn_frame.pack()
        self.backup_button = ttk.Button(btn_frame, text="تهیه فایل پشتیبان", command=self.backup_database); self.backup_button.pack(side='right', padx=10)
        self.restore_button = ttk.Button(btn_frame, text="بازیابی از فایل", command=self.restore_database); self.restore_button.pack(side='right', padx=10)
        return settings_tab
        
    def save_settings(self):
        try:
            for key, entry in self.settings_entries.items():
                self.db_manager.execute_query("REPLACE INTO settings (key, value) VALUES (?, ?)", (key, entry.get().strip()))
            self.db_manager.commit(); messagebox.showinfo("موفقیت", "تنظیمات با موفقیت ذخیره شد.")
            new_name = self.settings_entries['kindergarten_name'].get()
            self.winfo_toplevel().title(f"سامانه مدیریت {new_name}"); self.title_label.config(text=new_name)
        except Exception as e: self.db_manager.rollback(); messagebox.showerror("خطا", f"خطا در ذخیره تنظیمات: {e}")

    def backup_database(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".db", filetypes=[("Database files", "*.db"), ("All files", "*.*")], title="ذخیره فایل پشتیبان", initialfile=f"backup_{datetime.now().strftime('%Y-%m-%d')}.db")
        if not file_path: return
        self.backup_button.config(state="disabled"); self.restore_button.config(state="disabled"); self.backup_status_label.config(text="در حال تهیه فایل پشتیبان...", foreground="blue")
        threading.Thread(target=self._perform_backup, args=(file_path,), daemon=True).start()

    def _perform_backup(self, file_path):
        try:
            backup_conn = sqlite3.connect(file_path)
            with backup_conn: self.db_manager.conn.backup(backup_conn, pages=1, progress=None)
            backup_conn.close(); logging.info(f"پشتیبان دیتابیس در {file_path} ایجاد شد.")
            self.parent.after(100, lambda: self._backup_restore_finished(True, f"پشتیبان‌گیری با موفقیت در مسیر زیر انجام شد:\n{file_path}"))
        except Exception as e:
            logging.error(f"خطا در پشتیبان‌گیری: {e}"); self.parent.after(100, lambda: self._backup_restore_finished(False, f"خطا در هنگام پشتیبان‌گیری: {e}"))

    def restore_database(self):
        if not messagebox.askyesno("تایید بازیابی", "هشدار! با این کار تمام اطلاعات فعلی با اطلاعات فایل پشتیبان جایگزین خواهد شد. آیا مطمئن هستید؟"): return
        file_path = filedialog.askopenfilename(filetypes=[("Database files", "*.db"), ("All files", "*.*")], title="انتخاب فایل پشتیبان برای بازیابی")
        if not file_path: return
        self.backup_button.config(state="disabled"); self.restore_button.config(state="disabled"); self.backup_status_label.config(text="در حال بازیابی اطلاعات...", foreground="blue")
        threading.Thread(target=self._perform_restore, args=(file_path,), daemon=True).start()

    def _perform_restore(self, file_path):
        try:
            self.db_manager.close()
            shutil.copyfile(file_path, self.db_manager.db_path)
            logging.info(f"دیتابیس از فایل {file_path} بازیابی شد.")
            self.parent.after(100, self._restart_after_restore)
        except Exception as e:
            self.db_manager.connect(); logging.error(f"خطا در بازیابی: {e}"); self.parent.after(100, lambda: self._backup_restore_finished(False, f"خطا در هنگام بازیابی اطلاعات: {e}"))

    def _backup_restore_finished(self, success, message):
        self.backup_button.config(state="normal"); self.restore_button.config(state="normal")
        if success: self.backup_status_label.config(text=f"{self.icons['success']} عملیات موفق بود.", foreground="green"); messagebox.showinfo("موفقیت", message)
        else: self.backup_status_label.config(text=f"{self.icons['error']} عملیات ناموفق بود.", foreground="red"); messagebox.showerror("خطا", message)
        self.after(5000, lambda: self.backup_status_label.config(text=""))

    def _restart_after_restore(self):
        messagebox.showinfo("موفقیت", "بازیابی با موفقیت انجام شد. برنامه برای اعمال تغییرات بسته خواهد شد.")
        self.winfo_toplevel().destroy()


# ====================================================================================
# کلاس صفحه لاگین
# ====================================================================================
class LoginFrame(ttk.Frame):
    """فریم صفحه لاگین که در ابتدای برنامه نمایش داده می‌شود."""
    def __init__(self, parent, on_login_success_callback):
        # استایل‌های برنامه اصلی را برای ظاهر یکپارچه کپی می‌کنیم
        BG_COLOR, FRAME_COLOR, PRIMARY_COLOR, ACCENT_COLOR = "#F4F6F7", "#FFFFFF", "#1ABC9C", "#16A085"
        TEXT_COLOR, HEADER_TEXT_COLOR = "#34495E", "#FFFFFF"
        FONT_FAMILY = ("Vazirmatn", "Tahoma", "Arial")
        HEADER_FONT = (FONT_FAMILY[0], 15, 'bold')
        BODY_FONT = (FONT_FAMILY[0], 11)
        BOLD_FONT = (FONT_FAMILY[0], 11, 'bold')
        
        super().__init__(parent, padding=(40, 20))
        self.parent = parent
        self.on_login_success = on_login_success_callback
        self.configure(style='TFrame')

        # تنظیم استایل‌ها فقط برای این فریم
        style = ttk.Style(self)
        style.configure('Login.TFrame', background=BG_COLOR)
        style.configure('Login.TLabelframe', font=BOLD_FONT, background=FRAME_COLOR, borderwidth=1, relief="solid")
        style.configure('Login.TLabelframe.Label', font=HEADER_FONT, background=FRAME_COLOR, foreground=PRIMARY_COLOR, padding=10)
        style.configure('Login.TLabel', font=BOLD_FONT, background=FRAME_COLOR, foreground=TEXT_COLOR)
        style.configure('Login.TButton', font=BOLD_FONT, padding=12, background=PRIMARY_COLOR, foreground=HEADER_TEXT_COLOR, borderwidth=0)
        style.map('Login.TButton', background=[('active', ACCENT_COLOR)])
        style.configure('Login.TEntry', font=BODY_FONT, padding=8, relief='flat')

        # کانتینر اصلی برای وسط‌چین کردن فرم لاگین
        container = ttk.Frame(self, style='Login.TFrame')
        container.place(relx=0.5, rely=0.5, anchor="center")

        login_lframe = ttk.LabelFrame(container, text="ورود کاربر", style='Login.TLabelframe', padding=30)
        login_lframe.pack()

        # فیلد نام کاربری
        ttk.Label(login_lframe, text="نام کاربری:", style='Login.TLabel').grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.username_entry = ttk.Entry(login_lframe, style='Login.TEntry', justify='center', font=BODY_FONT)
        self.username_entry.grid(row=0, column=1, padx=10, pady=10)
        self.username_entry.insert(0, 'admin')

        # فیلد کلمه عبور
        ttk.Label(login_lframe, text="کلمه عبور:", style='Login.TLabel').grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.password_entry = ttk.Entry(login_lframe, show="*", style='Login.TEntry', justify='center', font=BODY_FONT)
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)
        self.password_entry.insert(0, 'admin')
        self.password_entry.bind("<Return>", self._attempt_login) # امکان ورود با کلید Enter

        # دکمه ورود
        login_button = ttk.Button(login_lframe, text="ورود", command=self._attempt_login, style='Login.TButton')
        login_button.grid(row=2, column=0, columnspan=2, pady=(20, 0), sticky="ew")

        self.username_entry.focus()

    def _attempt_login(self, event=None):
        """بررسی اطلاعات ورود کاربر."""
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if username == "admin" and password == "admin":
            logging.info(f"ورود موفق کاربر '{username}'")
            self.on_login_success()
        else:
            logging.warning(f"تلاش ناموفق برای ورود با نام کاربری: '{username}'")
            messagebox.showerror("خطای ورود", "نام کاربری یا کلمه عبور اشتباه است.", parent=self.parent)


# ====================================================================================
# تابع اصلی و اجرای برنامه
# ====================================================================================
def main():
    """تابع اصلی برای راه‌اندازی برنامه با صفحه لاگین."""
    root = tk.Tk()
    root.title("ورود به سامانه مدیریت مهدکودک")
    root.geometry("500x400") # اندازه اولیه برای پنجره لاگین
    root.minsize(450, 350)
    root.configure(bg="#F4F6F7")
    
    db_manager = None
    
    try:
        # 1. اتصال به دیتابیس
        db_manager = DatabaseManager()
        db_manager.connect()
        atexit.register(db_manager.close)

        def on_login_success():
            """این تابع پس از ورود موفق اجرا می‌شود."""
            # 1. حذف فریم لاگین
            login_frame.destroy()
            
            # 2. اطلاعات کاربر (در این نسخه ثابت است)
            user_info = {'username': 'admin', 'role': 'admin'}
            logging.info(f"کاربر '{user_info['username']}' با موفقیت وارد شد.")
            
            # 3. تنظیمات پنجره اصلی برای برنامه
            kindergarten_name = db_manager.get_setting('kindergarten_name', 'مهدکودک')
            root.title(f"سامانه مدیریت {kindergarten_name}")
            root.geometry("1400x950")
            root.minsize(1200, 800)

            # 4. ساخت و نمایش برنامه اصلی
            app = KindergartenApp(root, db_manager, user_info)
            
            # 5. مدیریت بسته شدن پنجره
            def on_closing():
                if messagebox.askokcancel("خروج", "آیا می‌خواهید از برنامه خارج شوید؟"):
                    root.destroy()
            root.protocol("WM_DELETE_WINDOW", on_closing)

        # ایجاد و نمایش فریم لاگین
        login_frame = LoginFrame(root, on_login_success)
        login_frame.pack(expand=True, fill="both")
        
        root.mainloop()

    except (DatabaseCriticalError, Exception) as e:
        # مدیریت هرگونه خطای بحرانی در هنگام راه‌اندازی
        logging.critical(f"خطای پیش‌بینی نشده در راه‌اندازی برنامه: {e}", exc_info=True)
        traceback.print_exc()
        messagebox.showerror("خطای بحرانی", f"یک خطای پیش‌بینی نشده در راه‌اندازی رخ داد:\n{e}\n\nبرنامه بسته خواهد شد.")
        if root:
            root.destroy()

if __name__ == "__main__":
    main()
