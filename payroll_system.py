# -*- coding: utf-8 -*-
import os
import sqlite3
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# =====[ PDF Export with Arabic Support ]=====
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    import platform
    HAS_PDF = True
    
    # محاولة استيراد مكتبات معالجة النص العربي
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        HAS_ARABIC_SUPPORT = True
        print("✅ تم تحميل مكتبات دعم العربية بنجاح")
    except ImportError:
        HAS_ARABIC_SUPPORT = False
        print("⚠️ مكتبات دعم العربية غير متوفرة - سيتم استخدام التحويل اليدوي")
        
except Exception:
    HAS_PDF = False
    HAS_ARABIC_SUPPORT = False

APP_TITLE = "نظام إدارة مرتبات العمال (حسب الأوردر)"
DB_PATH = os.path.join("data", "payroll_new.db")
REPORTS_DIR = "reports"

def register_arabic_font():
    """تسجيل خط عربي يدعمه reportlab مع دعم الاتجاه الصحيح"""
    try:
        system = platform.system()
        
        if system == "Windows":
            # خطوط تدعم العربية بشكل أفضل في Windows
            arabic_fonts = [
                "C:/Windows/Fonts/arial.ttf",
                "C:/Windows/Fonts/tahoma.ttf",
                "C:/Windows/Fonts/calibri.ttf",
                "C:/Windows/Fonts/segoeui.ttf",
                "C:/Windows/Fonts/times.ttf"
            ]
        elif system == "Linux":
            # خطوط تدعم العربية بشكل أفضل في Linux
            arabic_fonts = [
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
                "/usr/share/fonts/TTF/DejaVuSans.ttf",
                "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
                "/usr/share/fonts/truetype/noto/NotoNastaliqUrdu-Regular.ttf"
            ]
        else:  # macOS
            arabic_fonts = [
                "/System/Library/Fonts/Arial.ttf",
                "/System/Library/Fonts/Helvetica.ttf",
                "/Library/Fonts/Arial.ttf"
            ]
        
        for font_path in arabic_fonts:
            if os.path.exists(font_path):
                try:
                    # تسجيل الخط مع اسم مخصص
                    pdfmetrics.registerFont(TTFont('ArabicFont', font_path))
                    print(f"✅ تم تسجيل الخط العربي: {font_path}")
                    return 'ArabicFont'
                except Exception as e:
                    print(f"❌ فشل تسجيل الخط {font_path}: {e}")
                    continue
        
        # إذا لم يتم العثور على خط عربي، استخدم الخط الافتراضي
        print("⚠️ لم يتم العثور على خط عربي، استخدام الخط الافتراضي")
        return 'Helvetica'
            
    except Exception as e:
        print(f"❌ خطأ في تسجيل الخط العربي: {e}")
        return 'Helvetica'

def process_arabic_text(text):
    """معالجة النص العربي لعرضه بشكل صحيح في PDF"""
    try:
        if not isinstance(text, str):
            text = str(text)
            
        # التحقق من وجود نص عربي
        has_arabic = any('\u0600' <= char <= '\u06FF' for char in text)
        
        if has_arabic and HAS_ARABIC_SUPPORT:
            # استخدام مكتبات معالجة النص العربي
            try:
                reshaped_text = arabic_reshaper.reshape(text)
                bidi_text = get_display(reshaped_text)
                return bidi_text
            except Exception as e:
                print(f"خطأ في معالجة النص العربي بالمكتبات: {e}")
                return transliterate_arabic(text)
        elif has_arabic:
            # استخدام التحويل اليدوي
            return transliterate_arabic(text)
        else:
            return text
            
    except Exception as e:
        print(f"خطأ في معالجة النص: {e}")
        return str(text)

def transliterate_arabic(text):
    """تحويل النص العربي لأحرف لاتينية"""
    # أسماء شائعة - الأولوية للأسماء الكاملة
    common_names = {
        'أحمد': 'Ahmed',
        'محمد': 'Mohamed', 
        'خالد': 'Khaled',
        'محمود': 'Mahmoud',
        'سارة': 'Sara',
        'الغردقة': 'Hurghada',
        'القاهرة': 'Cairo',
        'الإسكندرية': 'Alexandria',
        'المرتب': 'Salary',
        'بدل الانتقالات': 'Transport',
        'الإجمالي': 'Total',
        'الموظف': 'Employee',
        'المنطقة': 'Area',
        'العنوان': 'Address',
        'التاريخ': 'Date',
        'تقرير الأوردر رقم': 'Order Report #',
        'إجمالي المبلغ': 'Total Amount',
        'جنيه مصري': 'EGP'
    }
    
    # البحث عن الأسماء الشائعة أولاً
    if text in common_names:
        return common_names[text]
    
    # البحث عن الكلمات المركبة
    for arabic, english in common_names.items():
        if arabic in text:
            text = text.replace(arabic, english)
    
    # التحويل حرف بحرف للكلمات المتبقية
    arabic_to_latin = {
        'أ': 'A', 'ا': 'A', 'ب': 'B', 'ت': 'T', 'ث': 'Th', 'ج': 'J', 'ح': 'H', 'خ': 'Kh',
        'د': 'D', 'ذ': 'Th', 'ر': 'R', 'ز': 'Z', 'س': 'S', 'ش': 'Sh', 'ص': 'S',
        'ض': 'D', 'ط': 'T', 'ظ': 'Z', 'ع': 'A', 'غ': 'Gh', 'ف': 'F', 'ق': 'Q',
        'ك': 'K', 'ل': 'L', 'م': 'M', 'ن': 'N', 'ه': 'H', 'و': 'W', 'ي': 'Y',
        'ة': 'h', 'ى': 'a', 'ئ': 'Y', 'ء': 'A', 'ؤ': 'W'
    }
    
    transliterated = ''
    for char in text:
        if char in arabic_to_latin:
            transliterated += arabic_to_latin[char]
        else:
            transliterated += char
    
    return transliterated

def draw_arabic_text(canvas, text, x, y, font_name, font_size):
    """رسم نص عربي مع تحسين الاتجاه"""
    try:
        # تحديد الخط
        try:
            canvas.setFont(font_name, font_size)
        except:
            canvas.setFont('Helvetica', font_size)
            font_name = 'Helvetica'
        
        # معالجة النص العربي
        processed_text = process_arabic_text(text)
        
        # رسم النص
        canvas.drawString(x, y, processed_text)
            
    except Exception as e:
        print(f"خطأ في رسم النص: {e}")
        # استخدام الخط الافتراضي كبديل
        try:
            canvas.setFont('Helvetica', font_size)
            canvas.drawString(x, y, str(text))
        except:
            canvas.setFont('Helvetica', 12)
            canvas.drawString(x, y, "Text Error")

# ============================= Helpers =============================
def ensure_dirs():
    os.makedirs("data", exist_ok=True)
    os.makedirs(REPORTS_DIR, exist_ok=True)

def db():
    return sqlite3.connect(DB_PATH)

def init_db():
    if not os.path.exists(DB_PATH):
        print("إنشاء قاعدة بيانات جديدة...")
    else:
        print("استخدام قاعدة البيانات الموجودة...")
    
    conn = db()
    c = conn.cursor()
    
    # Employees
    c.execute("""
    CREATE TABLE IF NOT EXISTS employees(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE
    )""")
    
    # Areas
    c.execute("""
    CREATE TABLE IF NOT EXISTS areas(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE
    )""")
    
    # Salary per (employee × area)
    c.execute("""
    CREATE TABLE IF NOT EXISTS employee_area_salary(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER NOT NULL,
        area_id INTEGER NOT NULL,
        salary REAL NOT NULL,
        UNIQUE(employee_id, area_id),
        FOREIGN KEY(employee_id) REFERENCES employees(id),
        FOREIGN KEY(area_id) REFERENCES areas(id)
    )""")
    
    # Orders
    c.execute("""
    CREATE TABLE IF NOT EXISTS orders(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        area_id INTEGER NOT NULL,
        address TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(area_id) REFERENCES areas(id)
    )""")
    
    # Order details (employees in each order)
    c.execute("""
    CREATE TABLE IF NOT EXISTS order_employees(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        employee_id INTEGER NOT NULL,
        salary REAL NOT NULL,
        transport REAL NOT NULL DEFAULT 0,
        total REAL NOT NULL,
        FOREIGN KEY(order_id) REFERENCES orders(id),
        FOREIGN KEY(employee_id) REFERENCES employees(id)
    )""")
    
    # إضافة بيانات تجريبية
    c.execute("SELECT COUNT(*) FROM employees")
    if c.fetchone()[0] == 0:
        print("إضافة بيانات تجريبية...")
        employees = [("أحمد",), ("محمد",), ("خالد",), ("محمود",), ("سارة",)]
        c.executemany("INSERT INTO employees(name) VALUES(?)", employees)
        
        areas = [("الغردقة",), ("القاهرة",), ("الإسكندرية",)]
        c.executemany("INSERT INTO areas(name) VALUES(?)", areas)
        
        # رواتب تجريبية
        c.execute("SELECT id,name FROM employees")
        emp_map = {name: emp_id for emp_id, name in c.fetchall()}
        c.execute("SELECT id,name FROM areas")
        area_map = {name: area_id for area_id, name in c.fetchall()}
        
        salary_data = [
            (emp_map["أحمد"], area_map["الغردقة"], 5000),
            (emp_map["محمد"], area_map["الغردقة"], 5400),
            (emp_map["خالد"], area_map["الغردقة"], 4800),
            (emp_map["محمود"], area_map["القاهرة"], 6100),
            (emp_map["سارة"], area_map["الإسكندرية"], 5900),
            (emp_map["أحمد"], area_map["القاهرة"], 7000),
        ]
        c.executemany("INSERT INTO employee_area_salary(employee_id, area_id, salary) VALUES (?,?,?)", salary_data)
    else:
        print("البيانات التجريبية موجودة بالفعل...")
    
    conn.commit()
    conn.close()
    print("تم إنشاء قاعدة البيانات بنجاح!")

# ============================= GUI =============================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1000x650")
        self.minsize(950, 600)
        style = ttk.Style(self)
        try:
            self.call("tk", "scaling", 1.15)
        except:
            pass
        if "clam" in style.theme_names():
            style.theme_use("clam")
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab_employees = EmployeesTab(nb)
        self.tab_areas = AreasTab(nb)
        self.tab_mapping = MappingTab(nb)
        self.tab_add_order = AddOrderTab(nb)
        self.tab_reports = ReportsTab(nb)
        for t, n in [(self.tab_employees, "الموظفون"),
                     (self.tab_areas, "المناطق"),
                     (self.tab_mapping, "رواتب (موظف × منطقة)"),
                     (self.tab_add_order, "إضافة أوردر جديد"),
                     (self.tab_reports, "التقارير")]:
            nb.add(t, text=n)
        # refresh hooks
        for t in [self.tab_employees, self.tab_areas, self.tab_mapping, self.tab_add_order, self.tab_reports]:
            t.set_refresh_hooks(self.refresh_all)

    def refresh_all(self):
        for t in [self.tab_employees, self.tab_areas, self.tab_mapping, self.tab_add_order, self.tab_reports]:
            t.refresh()

# ============================= Employees Tab =============================
class EmployeesTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.refresh_hook = None
        frm = ttk.LabelFrame(self, text="إضافة موظف")
        frm.pack(side="top", fill="x", padx=8, pady=8)
        ttk.Label(frm, text="اسم الموظف:").grid(row=0, column=0, padx=5, pady=8, sticky="e")
        self.entry_name = ttk.Entry(frm, width=40)
        self.entry_name.grid(row=0, column=1, padx=5, pady=8, sticky="w")
        ttk.Button(frm, text="إضافة", command=self.add_employee).grid(row=0, column=2, padx=5, pady=8)
        
        lst = ttk.LabelFrame(self, text="قائمة الموظفين")
        lst.pack(fill="both", expand=True, padx=8, pady=8)
        self.tree = ttk.Treeview(lst, columns=("id", "name"), show="headings", height=12)
        self.tree.heading("id", text="ID")
        self.tree.heading("name", text="الاسم")
        self.tree.column("id", width=70, anchor="center")
        self.tree.column("name", anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(lst, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btns, text="تعديل المحدد", command=self.edit_selected).pack(side="left", padx=5)
        ttk.Button(btns, text="حذف المحدد", command=self.delete_selected).pack(anchor="e")
        self.tree.bind('<Double-1>', lambda e: self.edit_selected())
        self.refresh()

    def set_refresh_hooks(self, fn):
        self.refresh_hook = fn

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        conn = db()
        c = conn.cursor()
        c.execute("SELECT id,name FROM employees ORDER BY id DESC")
        for rid, name in c.fetchall():
            self.tree.insert("", "end", values=(rid, name))
        conn.close()

    def add_employee(self):
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showerror("خطأ", "اكتب اسم الموظف.")
            return
        conn = db()
        c = conn.cursor()
        try:
            c.execute("INSERT INTO employees(name) VALUES(?)", (name,))
            conn.commit()
            self.entry_name.delete(0, tk.END)
            self.refresh()
            if self.refresh_hook:
                self.refresh_hook()
        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "الاسم موجود بالفعل.")
        finally:
            conn.close()

    def edit_selected(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("تحذير", "اختر موظف للتعديل")
            return
        emp_data = self.tree.item(selection[0])['values']
        emp_id, emp_name = emp_data[0], emp_data[1]
        
        dialog = tk.Toplevel(self)
        dialog.title("تعديل الموظف")
        dialog.geometry("300x120")
        dialog.grab_set()
        frame = ttk.Frame(dialog)
        frame.pack(fill="both", expand=True, padx=15, pady=15)
        ttk.Label(frame, text="الاسم الجديد:").pack(anchor="w")
        entry = ttk.Entry(frame, width=25)
        entry.insert(0, emp_name)
        entry.pack(fill="x", pady=5)
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=5)
        
        def save():
            new_name = entry.get().strip()
            if not new_name:
                messagebox.showerror("خطأ", "ادخل الاسم")
                return
            conn = db()
            c = conn.cursor()
            try:
                c.execute("UPDATE employees SET name=? WHERE id=?", (new_name, emp_id))
                conn.commit()
                dialog.destroy()
                self.refresh()
                messagebox.showinfo("تم", "تم التعديل")
            except sqlite3.IntegrityError:
                messagebox.showerror("خطأ", "الاسم موجود")
            finally:
                conn.close()
        
        ttk.Button(btn_frame, text="حفظ", command=save).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="إلغاء", command=dialog.destroy).pack(side="left", padx=5)
        entry.focus()
        entry.select_range(0, tk.END)

    def delete_selected(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("تنبيه", "اختر موظفًا أولًا.")
            return
        rid = self.tree.item(selection[0], "values")[0]
        if not messagebox.askyesno("تأكيد", "هل تريد حذف الموظف المحدد؟"):
            return
        conn = db()
        c = conn.cursor()
        try:
            c.execute("DELETE FROM employees WHERE id=?", (rid,))
            conn.commit()
            self.refresh()
        except sqlite3.IntegrityError as e:
            messagebox.showerror("خطأ", f"تعذر الحذف.\n{e}")
        finally:
            conn.close()

# ============================= Areas Tab =============================
class AreasTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.refresh_hook = None
        frm = ttk.LabelFrame(self, text="إضافة منطقة")
        frm.pack(side="top", fill="x", padx=8, pady=8)
        ttk.Label(frm, text="اسم المنطقة:").grid(row=0, column=0, padx=5, pady=8, sticky="e")
        self.entry_name = ttk.Entry(frm, width=40)
        self.entry_name.grid(row=0, column=1, padx=5, pady=8, sticky="w")
        ttk.Button(frm, text="إضافة", command=self.add_area).grid(row=0, column=2, padx=5, pady=8)
        lst = ttk.LabelFrame(self, text="قائمة المناطق")
        lst.pack(fill="both", expand=True, padx=8, pady=8)
        self.tree = ttk.Treeview(lst, columns=("id", "name"), show="headings", height=12)
        self.tree.heading("id", text="ID")
        self.tree.heading("name", text="الاسم")
        self.tree.column("id", width=70, anchor="center")
        self.tree.column("name", anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(lst, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btns, text="حذف المحدد", command=self.delete_selected).pack(anchor="e")
        self.refresh()

    def set_refresh_hooks(self, fn):
        self.refresh_hook = fn

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        conn = db()
        c = conn.cursor()
        c.execute("SELECT id,name FROM areas ORDER BY id DESC")
        for rid, name in c.fetchall():
            self.tree.insert("", "end", values=(rid, name))
        conn.close()

    def add_area(self):
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showerror("خطأ", "اكتب اسم المنطقة.")
            return
        conn = db()
        c = conn.cursor()
        try:
            c.execute("INSERT INTO areas(name) VALUES(?)", (name,))
            conn.commit()
            self.entry_name.delete(0, tk.END)
            self.refresh()
        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "الاسم موجود بالفعل.")
        finally:
            conn.close()

    def delete_selected(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("تنبيه", "اختر منطقة أولًا.")
            return
        rid = self.tree.item(selection[0], "values")[0]
        if not messagebox.askyesno("تأكيد", "هل تريد حذف المنطقة المحددة؟"):
            return
        conn = db()
        c = conn.cursor()
        try:
            c.execute("DELETE FROM areas WHERE id=?", (rid,))
            conn.commit()
            self.refresh()
        except sqlite3.IntegrityError as e:
            messagebox.showerror("خطأ", f"تعذر الحذف.\n{e}")
        finally:
            conn.close()

# ============================= Mapping Tab =============================
class MappingTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.refresh_hook = None
        frm = ttk.LabelFrame(self, text="تحديد المرتب (موظف × منطقة)")
        frm.pack(side="top", fill="x", padx=8, pady=8)
        ttk.Label(frm, text="الموظف:").grid(row=0, column=0, padx=5, pady=8, sticky="e")
        self.cmb_employee = ttk.Combobox(frm, state="readonly")
        self.cmb_employee.grid(row=0, column=1, padx=5, pady=8, sticky="w")
        ttk.Label(frm, text="المنطقة:").grid(row=0, column=2, padx=5, pady=8, sticky="e")
        self.cmb_area = ttk.Combobox(frm, state="readonly")
        self.cmb_area.grid(row=0, column=3, padx=5, pady=8, sticky="w")
        ttk.Label(frm, text="المرتب:").grid(row=0, column=4, padx=5, pady=8, sticky="e")
        self.entry_salary = ttk.Entry(frm, width=12)
        self.entry_salary.grid(row=0, column=5, padx=5, pady=8, sticky="w")
        ttk.Button(frm, text="حفظ", command=self.save_mapping).grid(row=0, column=6, padx=5, pady=8)
        lst = ttk.LabelFrame(self, text="الرواتب المسجلة")
        lst.pack(fill="both", expand=True, padx=8, pady=8)
        self.tree = ttk.Treeview(lst, columns=("id", "employee", "area", "salary"), show="headings", height=12)
        self.tree.heading("id", text="ID")
        self.tree.heading("employee", text="الموظف")
        self.tree.heading("area", text="المنطقة")
        self.tree.heading("salary", text="المرتب")
        self.tree.column("id", width=70, anchor="center")
        self.tree.column("employee", anchor="center")
        self.tree.column("area", anchor="center")
        self.tree.column("salary", anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(lst, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btns, text="حذف المحدد", command=self.delete_selected).pack(anchor="e")
        self.refresh()

    def set_refresh_hooks(self, fn):
        self.refresh_hook = fn

    def refresh(self):
        conn = db()
        c = conn.cursor()
        c.execute("SELECT id,name FROM employees ORDER BY name")
        emps = c.fetchall()
        self.emp_map = {name: emp_id for emp_id, name in emps}
        self.cmb_employee['values'] = [name for emp_id, name in emps]
        c.execute("SELECT id,name FROM areas ORDER BY name")
        areas = c.fetchall()
        self.area_map = {name: area_id for area_id, name in areas}
        self.cmb_area['values'] = [name for area_id, name in areas]
        for i in self.tree.get_children():
            self.tree.delete(i)
        c.execute("SELECT mas.id,e.name,a.name,mas.salary FROM employee_area_salary mas JOIN employees e ON e.id=mas.employee_id JOIN areas a ON a.id=mas.area_id ORDER BY a.name,e.name")
        for rid, emp, area, sal in c.fetchall():
            self.tree.insert("", "end", values=(rid, emp, area, sal))
        conn.close()

    def save_mapping(self):
        emp = self.cmb_employee.get()
        area = self.cmb_area.get()
        sal = self.entry_salary.get().strip()
        if not emp or not area or not sal:
            messagebox.showerror("خطأ", "اكمل جميع الحقول")
            return
        try:
            sal = float(sal)
        except:
            messagebox.showerror("خطأ", "المرتب يجب أن يكون رقم")
            return
        emp_id = self.emp_map[emp]
        area_id = self.area_map[area]
        conn = db()
        c = conn.cursor()
        try:
            c.execute("INSERT OR REPLACE INTO employee_area_salary(employee_id, area_id, salary) VALUES(?,?,?)", (emp_id, area_id, sal))
            conn.commit()
            self.entry_salary.delete(0, tk.END)
            self.refresh()
        except Exception as e:
            messagebox.showerror("خطأ", str(e))
        finally:
            conn.close()

    def delete_selected(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("تنبيه", "اختر عنصرًا أولًا.")
            return
        rid = self.tree.item(selection[0], "values")[0]
        if not messagebox.askyesno("تأكيد", "هل تريد حذف العنصر المحدد؟"):
            return
        conn = db()
        c = conn.cursor()
        try:
            c.execute("DELETE FROM employee_area_salary WHERE id=?", (rid,))
            conn.commit()
            self.refresh()
        finally:
            conn.close()

# ============================= Add Order Tab =============================
class AddOrderTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.refresh_hook = None
        self.selected_employees = []
        
        frm = ttk.LabelFrame(self, text="إضافة أوردر جديد")
        frm.pack(side="top", fill="x", padx=8, pady=8)
        
        ttk.Label(frm, text="المنطقة:").grid(row=0, column=0, padx=5, pady=8, sticky="e")
        self.cmb_area = ttk.Combobox(frm, state="readonly")
        self.cmb_area.grid(row=0, column=1, padx=5, pady=8, sticky="w")
        ttk.Label(frm, text="العنوان:").grid(row=0, column=2, padx=5, pady=8, sticky="e")
        self.entry_address = ttk.Entry(frm, width=40)
        self.entry_address.grid(row=0, column=3, padx=5, pady=8, sticky="w")
        
        self.btn_pick_employees = ttk.Button(frm, text="اختيار الموظفين (متعددين)", command=self.pick_employees, state="disabled")
        self.btn_pick_employees.grid(row=0, column=4, padx=5, pady=8)
        
        self.btn_edit_transport = ttk.Button(frm, text="تعديل بدل الانتقالات", command=self.edit_transport, state="disabled")
        self.btn_edit_transport.grid(row=1, column=0, columnspan=2, padx=5, pady=8, sticky="w")
        
        self.btn_save_order = ttk.Button(frm, text="حفظ الأوردر", command=self.save_order, state="disabled")
        self.btn_save_order.grid(row=1, column=2, columnspan=2, padx=5, pady=8, sticky="w")
        
        ttk.Button(frm, text="مسح الكل", command=self.clear_all).grid(row=1, column=4, padx=5, pady=8)
        
        self.lbl_selected = ttk.Label(frm, text="الموظفين المختارين: 0")
        self.lbl_selected.grid(row=2, column=0, columnspan=6, padx=5, pady=8, sticky="w")
        
        self.cmb_area.bind('<<ComboboxSelected>>', self.on_area_selected)
        
        preview = ttk.LabelFrame(self, text="معاينة الأوردر")
        preview.pack(fill="both", expand=True, padx=8, pady=8)
        
        self.preview_tree = ttk.Treeview(preview, columns=("employee", "salary", "transport", "total"), show="headings", height=8)
        for col, txt in [("employee", "الموظف"), ("salary", "المرتب"), ("transport", "بدل انتقالات"), ("total", "الإجمالي")]:
            self.preview_tree.heading(col, text=txt)
            self.preview_tree.column(col, width=120, anchor="center")
        self.preview_tree.pack(fill="both", expand=True)
        
        self.lbl_total = ttk.Label(preview, text="إجمالي الأوردر: 0 جنيه مصري", font=('Arial', 12, 'bold'))
        self.lbl_total.pack(anchor="e", padx=10, pady=5)
        
        self.refresh()
        
    def set_refresh_hooks(self, fn):
        self.refresh_hook = fn

    def on_area_selected(self, event=None):
        if self.cmb_area.get():
            self.btn_pick_employees.config(state="normal")

    def refresh(self):
        conn = db()
        c = conn.cursor()
        c.execute("SELECT id,name FROM areas ORDER BY name")
        areas_data = c.fetchall()
        self.area_map = {name: area_id for area_id, name in areas_data}
        self.cmb_area['values'] = [name for area_id, name in areas_data]
        conn.close()
        
    def pick_employees(self):
        area = self.cmb_area.get()
        if not area:
            messagebox.showerror("خطأ", "اختر المنطقة أولًا")
            return
        
        area_id = self.area_map.get(area)
        if area_id is None:
            messagebox.showerror("خطأ", f"لم يتم العثور على معرف للمنطقة: {area}")
            return
        
        conn = db()
        c = conn.cursor()
        
        c.execute("SELECT e.id,e.name,mas.salary FROM employees e JOIN employee_area_salary mas ON mas.employee_id=e.id WHERE mas.area_id=?", (area_id,))
        rows = c.fetchall()
        
        if not rows:
            c.execute("SELECT id,name FROM employees ORDER BY name")
            all_employees = c.fetchall()
            rows = [(emp_id, emp_name, 5000) for emp_id, emp_name in all_employees]
        
        conn.close()
        
        if not rows:
            messagebox.showinfo("تنبيه", "لا يوجد موظفين في النظام.")
            return
        
        top = tk.Toplevel(self)
        top.title("اختيار الموظفين")
        top.geometry("400x300")
        top.grab_set()
        ttk.Label(top, text="اختر الموظفين (Ctrl للمتعدد):", font=("Arial", 10, "bold")).pack(pady=5)
        
        lst = tk.Listbox(top, selectmode="extended")
        lst.pack(fill="both", expand=True, padx=10, pady=10)
        self.emp_rows_map = {}
        for idx, (eid, name, sal) in enumerate(rows):
            lst.insert("end", f"{name} ({sal} جنيه مصري)")
            self.emp_rows_map[idx] = {'id': eid, 'name': name, 'salary': sal, 'transport': 0}
        
        def select_all():
            lst.select_set(0, tk.END)

        def on_done():
            sel = lst.curselection()
            if not sel:
                messagebox.showwarning("تحذير", "اختر موظف واحد على الأقل")
                return
            self.selected_employees = [self.emp_rows_map[i] for i in sel]
            self.update_preview()
            self.lbl_selected.config(text=f"الموظفين المختارين: {len(self.selected_employees)}")
            if self.selected_employees:
                self.btn_edit_transport.config(state="normal")
                self.btn_save_order.config(state="normal")
            top.destroy()
        
        btn_frame = ttk.Frame(top)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="اختيار الكل", command=select_all).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="تم", command=on_done).pack(side="right", padx=5)
    
    def update_preview(self):
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        total_amount = 0
        for emp in self.selected_employees:
            emp_total = emp['salary'] + emp['transport']
            total_amount += emp_total
            self.preview_tree.insert("", "end", values=(emp['name'], emp['salary'], emp['transport'], emp_total))
        self.lbl_total.config(text=f"إجمالي الأوردر: {total_amount} جنيه مصري")
    
    def edit_transport(self):
        if not self.selected_employees:
            messagebox.showwarning("تحذير", "لا يوجد موظفين محددين")
            return
        dialog = tk.Toplevel(self)
        dialog.title("بدل الانتقالات")
        dialog.geometry("350x250")
        dialog.grab_set()
        frame = ttk.Frame(dialog)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        ttk.Label(frame, text="بدل الانتقالات لكل موظف:", font=("Arial", 10, "bold")).pack(anchor="w", pady=5)
        
        entries = []
        for emp in self.selected_employees:
            emp_frame = ttk.Frame(frame)
            emp_frame.pack(fill="x", pady=2)
            ttk.Label(emp_frame, text=f"{emp['name']}:", width=15).pack(side="left")
            entry = ttk.Entry(emp_frame, width=8)
            entry.insert(0, str(emp['transport']))
            entry.pack(side="left", padx=5)
            ttk.Label(emp_frame, text="جنيه مصري").pack(side="left")
            entries.append(entry)
        
        def apply():
            try:
                for i, entry in enumerate(entries):
                    self.selected_employees[i]['transport'] = float(entry.get() or "0")
                self.update_preview()
                dialog.destroy()
                messagebox.showinfo("تم", "تم التحديث")
            except ValueError:
                messagebox.showerror("خطأ", "ادخل أرقام")
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=10)
        ttk.Button(btn_frame, text="تطبيق", command=apply).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="إلغاء", command=dialog.destroy).pack(side="right", padx=5)
    
    def save_order(self):
        area = self.cmb_area.get()
        address = self.entry_address.get().strip()
        if not area:
            messagebox.showerror("خطأ", "اختر المنطقة")
            return
        if not self.selected_employees:
            messagebox.showerror("خطأ", "اختر موظفين")
            return
        area_id = self.area_map[area]
        created_at = datetime.now().isoformat()
        conn = db()
        c = conn.cursor()
        try:
            c.execute("INSERT INTO orders(area_id,address,created_at) VALUES(?,?,?)", (area_id, address, created_at))
            order_id = c.lastrowid
            for emp in self.selected_employees:
                emp_total = emp['salary'] + emp['transport']
                c.execute("INSERT INTO order_employees(order_id,employee_id,salary,transport,total) VALUES(?,?,?,?,?)", (order_id, emp['id'], emp['salary'], emp['transport'], emp_total))
            conn.commit()
            total_amount = sum(emp['salary'] + emp['transport'] for emp in self.selected_employees)
            messagebox.showinfo("نجاح", f"تم إضافة الأوردر رقم {order_id}\nإجمالي: {total_amount} جنيه مصري")
            self.clear_all()
        except Exception as e:
            messagebox.showerror("خطأ", f"خطأ: {str(e)}")
        finally:
            conn.close()
    
    def clear_all(self):
        self.selected_employees = []
        self.lbl_selected.config(text="الموظفين المختارين: 0")
        self.entry_address.delete(0, tk.END)
        self.cmb_area.set('')
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self.lbl_total.config(text="إجمالي الأوردر: 0 جنيه مصري")
        self.btn_pick_employees.config(state="disabled")
        self.btn_edit_transport.config(state="disabled")
        self.btn_save_order.config(state="disabled")

# ============================= Reports Tab =============================
class ReportsTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.refresh_hook = None
        frm = ttk.LabelFrame(self, text="تقرير الأوردرات")
        frm.pack(side="top", fill="x", padx=8, pady=8)
        ttk.Label(frm, text="اختر الأوردر:").grid(row=0, column=0, padx=5, pady=8, sticky="e")
        self.cmb_orders = ttk.Combobox(frm, state="readonly", width=50)
        self.cmb_orders.grid(row=0, column=1, padx=5, pady=8, sticky="w")
        ttk.Button(frm, text="عرض التقرير", command=self.show_report).grid(row=0, column=2, padx=5, pady=8)
        ttk.Button(frm, text="تحديث القائمة", command=self.refresh).grid(row=0, column=3, padx=5, pady=8)
        if HAS_PDF:
            ttk.Button(frm, text="تصدير PDF", command=self.export_pdf).grid(row=0, column=4, padx=5, pady=8)
        lst = ttk.LabelFrame(self, text="تفاصيل الأوردر")
        lst.pack(fill="both", expand=True, padx=8, pady=8)
        self.tree = ttk.Treeview(lst, columns=("employee", "salary", "transport", "total"), show="headings", height=15)
        for col, txt in [("employee", "الموظف"), ("salary", "المرتب"), ("transport", "بدل الانتقالات"), ("total", "الإجمالي")]:
            self.tree.heading(col, text=txt)
            self.tree.column(col, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(lst, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        self.lbl_total = ttk.Label(self, text="إجمالي الأوردر: 0")
        self.lbl_total.pack(anchor="e", padx=10, pady=5)
        self.refresh()

    def set_refresh_hooks(self, fn):
        self.refresh_hook = fn

    def refresh(self):
        conn = db()
        c = conn.cursor()
        c.execute("SELECT o.id,a.name,o.address,o.created_at FROM orders o JOIN areas a ON a.id=o.area_id ORDER BY o.id DESC")
        self.orders_map = {f"{a} | {ad} | {dt}": oid for oid, a, ad, dt in c.fetchall()}
        self.cmb_orders['values'] = list(self.orders_map.keys())
        conn.close()

    def show_report(self):
        sel = self.cmb_orders.get()
        if not sel:
            messagebox.showerror("خطأ", "اختر أوردر")
            return
        order_id = self.orders_map[sel]
        conn = db()
        c = conn.cursor()
        c.execute("SELECT e.name,oe.salary,oe.transport,oe.total FROM order_employees oe JOIN employees e ON e.id=oe.employee_id WHERE oe.order_id=?", (order_id,))
        rows = c.fetchall()
        conn.close()
        [self.tree.delete(i) for i in self.tree.get_children()]
        total = 0
        for emp, sal, tran, tot in rows:
            self.tree.insert("", "end", values=(emp, sal, tran, tot))
            total += tot
        self.lbl_total.config(text=f"إجمالي الأوردر: {total}")

    def export_pdf(self):
        if not HAS_PDF:
            messagebox.showerror("خطأ", "مكتبة reportlab غير مثبتة")
            return
        sel = self.cmb_orders.get()
        if not sel:
            messagebox.showerror("خطأ", "اختر أوردر")
            return
        order_id = self.orders_map[sel]
        conn = db()
        c = conn.cursor()
        c.execute("SELECT a.name,o.address,o.created_at FROM orders o JOIN areas a ON a.id=o.area_id WHERE o.id=?", (order_id,))
        area, address, dt = c.fetchone()
        c.execute("SELECT e.name,oe.salary,oe.transport,oe.total FROM order_employees oe JOIN employees e ON e.id=oe.employee_id WHERE oe.order_id=?", (order_id,))
        rows = c.fetchall()
        conn.close()
        fname = f"{REPORTS_DIR}/Order_{order_id}.pdf"
        
        # تسجيل الخط العربي
        arabic_font_name = register_arabic_font()
        
        cpdf = canvas.Canvas(fname, pagesize=A4)
        
        # عنوان التقرير
        title = f"تقرير الأوردر رقم {order_id}"
        draw_arabic_text(cpdf, title, 2*cm, 27*cm, arabic_font_name, 16)
        
        # معلومات الأوردر
        y_pos = 26*cm
        
        # المنطقة
        area_text = f"المنطقة: {area}"
        draw_arabic_text(cpdf, area_text, 2*cm, y_pos, arabic_font_name, 12)
        y_pos -= 0.8*cm
        
        # العنوان
        if address:
            address_text = f"العنوان: {address}"
            draw_arabic_text(cpdf, address_text, 2*cm, y_pos, arabic_font_name, 12)
            y_pos -= 0.8*cm
        
        # التاريخ
        date_text = f"التاريخ: {dt}"
        draw_arabic_text(cpdf, date_text, 2*cm, y_pos, arabic_font_name, 12)
        y_pos -= 1.2*cm
        
        # رؤوس الأعمدة
        headers = ["الموظف", "المرتب", "بدل الانتقالات", "الإجمالي"]
        x_positions = [2*cm, 6*cm, 11*cm, 15*cm]
        
        for i, header in enumerate(headers):
            draw_arabic_text(cpdf, header, x_positions[i], y_pos, arabic_font_name, 11)
        
        y_pos -= 0.8*cm
        
        # خط فاصل
        cpdf.line(2*cm, y_pos + 0.2*cm, 19*cm, y_pos + 0.2*cm)
        y_pos -= 0.5*cm
        
        # بيانات الموظفين
        total_amount = 0
        for emp_name, salary, transport, total in rows:
            if y_pos < 3*cm:  # إذا وصلنا لنهاية الصفحة
                cpdf.showPage()
                y_pos = 27*cm
            
            draw_arabic_text(cpdf, str(emp_name), x_positions[0], y_pos, arabic_font_name, 10)
            draw_arabic_text(cpdf, str(salary), x_positions[1], y_pos, arabic_font_name, 10)
            draw_arabic_text(cpdf, str(transport), x_positions[2], y_pos, arabic_font_name, 10)
            draw_arabic_text(cpdf, str(total), x_positions[3], y_pos, arabic_font_name, 10)
            
            total_amount += total
            y_pos -= 0.6*cm
        
        # خط فاصل قبل الإجمالي
        y_pos -= 0.3*cm
        cpdf.line(2*cm, y_pos, 19*cm, y_pos)
        y_pos -= 0.8*cm
        
        # الإجمالي
        total_text = f"إجمالي المبلغ: {total_amount} جنيه مصري"
        draw_arabic_text(cpdf, total_text, 12*cm, y_pos, arabic_font_name, 12)
        
        cpdf.save()
        messagebox.showinfo("تم", f"تم تصدير التقرير إلى:\n{fname}")

# ============================= Main =============================
def main():
    ensure_dirs()
    init_db()
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()