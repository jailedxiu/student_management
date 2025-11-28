#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
学生档案管理系统（增强版：列分类、按大类筛选、学生详情编辑）
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import sqlite3
from datetime import datetime
import traceback

# 检查依赖
def check_dependencies():
    """检查并安装依赖"""
    missing = []
    
    try:
        import pandas  # noqa
    except ImportError:
        missing.append('pandas')
    
    try:
        import openpyxl  # noqa
    except ImportError:
        missing.append('openpyxl')
    
    if missing:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        msg = f"缺少以下依赖库：{', '.join(missing)}\n\n"
        msg += "请在命令行运行以下命令安装：\n"
        msg += f"pip install {' '.join(missing)}\n\n"
        msg += "或者使用国内镜像（更快）：\n"
        msg += f"pip install -i https://pypi.tuna.tsinghua.edu.cn/simple {' '.join(missing)}"
        messagebox.showerror("缺少依赖", msg)
        return False
    
    return True

if not check_dependencies():
    sys.exit(1)


def center_window_relative(child, parent, offset_x=0, offset_y=0):
    """让子窗口基于父窗口位置进行居中/偏移显示"""
    if parent is None or child is None:
        return
    parent.update_idletasks()
    child.update_idletasks()
    
    px, py = parent.winfo_x(), parent.winfo_y()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    
    cw, ch = child.winfo_width(), child.winfo_height()
    if cw <= 1 or ch <= 1:
        cw = child.winfo_reqwidth()
        ch = child.winfo_reqheight()
    
    x = px + (pw - cw) // 2 + offset_x
    y = py + (ph - ch) // 2 + offset_y
    x = max(x, 0)
    y = max(y, 0)
    
    child.geometry(f"{cw}x{ch}+{x}+{y}")
    child.deiconify()


class StudentManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("25级汽修一班学生档案管理系统")
        self.root.geometry("1200x700")
        
        if getattr(sys, 'frozen', False):
            self.base_path = os.path.dirname(sys.executable)
        else:
            self.base_path = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.base_path)
        
        self.db_path = os.path.join(self.base_path, "student_database.db")
        self.init_database()
        
        # 调整此列表即可改变列的默认展示顺序（未列出的字段按字母顺序排列在后）
        self.preferred_order = [
            "学号",
            "姓名",
            "身份证号码",
            "出生日期",
            "电话号码",
            "母亲姓名",
            "母亲电话号码",
            "父亲姓名",
            "父亲电话号码",
            "最后一次谈话日期",
            "谈话内容",
            "红黄蓝情况简述",
            "中考名次",
            "中考成绩",
        ]
        
        self.is_filtered = False
        self.current_filter_desc = ""
        self.column_category_map = {}
        self.column_categories = []
        self.all_columns = []
        
        self.create_widgets()
        self.refresh_column_list()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def on_closing(self):
        self.root.destroy()
    
    @staticmethod
    def normalize_cell_value(value):
        """统一处理写入数据库的单元格值"""
        if pd.isna(value):
            return None
        if isinstance(value, (pd.Timestamp, datetime)):
            return value.strftime("%Y-%m-%d")
        return str(value)
    
    def sort_columns(self, columns):
        """根据预设顺序对列名进行排序"""
        unique_columns = []
        for col in columns:
            if col not in unique_columns:
                unique_columns.append(col)
        order_map = {col: idx for idx, col in enumerate(self.preferred_order)}
        return sorted(unique_columns, key=lambda c: (order_map.get(c, len(order_map)), c.lower()))
    
    def get_columns_grouped_by_category(self, columns_override=None):
        """按列大类分组（供详情页等功能使用）"""
        if columns_override is None:
            source_columns = [col for col in self.all_columns if col != "学号"]
        else:
            source_columns = [col for col in columns_override if col != "学号"]
        ordered_columns = self.sort_columns(source_columns)
        grouped = {}
        for col in ordered_columns:
            category = self.column_category_map.get(col, "未分类")
            grouped.setdefault(category, []).append(col)
        ordered_categories = sorted(grouped.keys(), key=lambda c: (c == "未分类", c))
        return [(cat, grouped[cat]) for cat in ordered_categories]
    
    def init_database(self):
        """初始化数据库"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS students (
                    学号 TEXT PRIMARY KEY
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS column_history (
                    column_name TEXT PRIMARY KEY,
                    first_upload_time TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS column_categories (
                    category_name TEXT PRIMARY KEY,
                    created_at TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS column_category_map (
                    column_name TEXT PRIMARY KEY,
                    category_name TEXT,
                    FOREIGN KEY (column_name) REFERENCES column_history(column_name),
                    FOREIGN KEY (category_name) REFERENCES column_categories(category_name)
                )
            ''')
            
            conn.commit()
            conn.close()
        except Exception as e:
            messagebox.showerror("数据库错误", f"初始化数据库失败：\n{str(e)}")
    
    def create_widgets(self):
        """创建界面组件"""
        top_frame = tk.Frame(self.root, pady=10)
        top_frame.pack(fill=tk.X)
        
        btn_import = tk.Button(top_frame, text="导入Excel", command=self.import_excel, 
                 bg="#4CAF50", fg="white", font=("Arial", 11), padx=15, pady=5, width=10)
        btn_import.pack(side=tk.LEFT, padx=5)
        
        btn_export = tk.Button(top_frame, text="导出Excel", command=self.export_excel,
                 bg="#2196F3", fg="white", font=("Arial", 11), padx=15, pady=5, width=10)
        btn_export.pack(side=tk.LEFT, padx=5)
        
        btn_refresh = tk.Button(top_frame, text="刷新显示", command=self.refresh_display,
                 bg="#FF9800", fg="white", font=("Arial", 11), padx=15, pady=5, width=10)
        btn_refresh.pack(side=tk.LEFT, padx=5)
        
        btn_clear = tk.Button(top_frame, text="清空筛选", command=self.clear_filters,
                 bg="#9E9E9E", fg="white", font=("Arial", 11), padx=15, pady=5, width=10)
        btn_clear.pack(side=tk.LEFT, padx=5)
        
        btn_delete_row = tk.Button(top_frame, text="删除学生", command=self.delete_row,
                 bg="#F44336", fg="white", font=("Arial", 11), padx=15, pady=5, width=10)
        btn_delete_row.pack(side=tk.LEFT, padx=5)
        
        btn_delete_column = tk.Button(top_frame, text="删除列", command=self.delete_column,
                 bg="#E91E63", fg="white", font=("Arial", 11), padx=15, pady=5, width=10)
        btn_delete_column.pack(side=tk.LEFT, padx=5)
        
        btn_category = tk.Button(top_frame, text="列分类", command=self.open_column_category_manager,
                 bg="#8E24AA", fg="white", font=("Arial", 11), padx=15, pady=5, width=10)
        btn_category.pack(side=tk.LEFT, padx=5)
        
        filter_frame = tk.LabelFrame(self.root, text="筛选条件", font=("Arial", 11, "bold"), pady=10)
        filter_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(filter_frame, text="姓名搜索:", font=("Arial", 10)).grid(row=0, column=0, padx=5, sticky=tk.W)
        self.name_entry = tk.Entry(filter_frame, font=("Arial", 10), width=20)
        self.name_entry.grid(row=0, column=1, padx=5)
        tk.Button(filter_frame, text="搜索", command=self.search_by_name,
                 bg="#607D8B", fg="white", font=("Arial", 10), width=8).grid(row=0, column=2, padx=5)
        
        tk.Label(filter_frame, text="列大类:", font=("Arial", 10)).grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.filter_category_combo = ttk.Combobox(filter_frame, font=("Arial", 10), width=18, state="readonly")
        self.filter_category_combo.grid(row=1, column=1, padx=5, pady=5)
        self.filter_category_combo.bind("<<ComboboxSelected>>", self.on_filter_category_change)
        tk.Button(filter_frame, text="全部列", command=self.reset_category_filter,
                 bg="#B0BEC5", fg="white", font=("Arial", 10), width=8).grid(row=1, column=2, padx=5)
        
        tk.Label(filter_frame, text="按列筛选:", font=("Arial", 10)).grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.filter_column_combo = ttk.Combobox(filter_frame, font=("Arial", 10), width=18, state="readonly")
        self.filter_column_combo.grid(row=2, column=1, padx=5, pady=5)
        self.filter_column_combo.bind("<<ComboboxSelected>>", self.load_filter_values)
        
        tk.Label(filter_frame, text="筛选值:", font=("Arial", 10)).grid(row=2, column=3, padx=5, sticky=tk.W)
        self.filter_value_combo = ttk.Combobox(filter_frame, font=("Arial", 10), width=18, state="readonly")
        self.filter_value_combo.grid(row=2, column=4, padx=5, pady=5)
        tk.Button(filter_frame, text="应用筛选", command=self.apply_filter,
                 bg="#607D8B", fg="white", font=("Arial", 10), width=8).grid(row=2, column=5, padx=5)
        
        display_frame = tk.Frame(self.root)
        display_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tree_scroll_x = tk.Scrollbar(display_frame, orient=tk.HORIZONTAL)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.tree_scroll_y = tk.Scrollbar(display_frame, orient=tk.VERTICAL)
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree = ttk.Treeview(display_frame, 
                                 yscrollcommand=self.tree_scroll_y.set,
                                 xscrollcommand=self.tree_scroll_x.set,
                                 selectmode="extended")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)
        
        self.status_bar = tk.Label(self.root, text="就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def refresh_column_list(self):
        """刷新列名和分类"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall()]
            conn.close()
            
            if "学号" not in columns:
                columns.insert(0, "学号")
            
            columns = self.sort_columns(columns)
            self.all_columns = columns
            self.column_category_map = self.fetch_column_category_map()
            self.column_categories = self.get_all_categories()
            self.refresh_category_filter()
            
            return columns
        except Exception as e:
            messagebox.showerror("错误", f"刷新列名失败：{str(e)}")
            return []
    
    def fetch_column_category_map(self):
        mapping = {}
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT column_name, category_name FROM column_category_map")
            mapping = {row[0]: row[1] for row in cursor.fetchall()}
            conn.close()
        except Exception as e:
            print(f"加载列分类失败: {e}")
        return mapping
    
    def get_all_categories(self):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT category_name FROM column_categories ORDER BY category_name COLLATE NOCASE")
            rows = cursor.fetchall()
            conn.close()
            return [row[0] for row in rows]
        except Exception as e:
            messagebox.showerror("错误", f"加载列大类失败：{str(e)}")
            return []
    
    def refresh_category_filter(self):
        if not hasattr(self, 'filter_category_combo'):
            return
        options = ["全部"] + self.column_categories if self.column_categories else ["全部"]
        current = self.filter_category_combo.get()
        self.filter_category_combo['values'] = options
        if current in options:
            self.filter_category_combo.set(current)
        else:
            self.filter_category_combo.set("全部")
        self.update_filter_column_options()
    
    def update_filter_column_options(self):
        if not hasattr(self, 'filter_column_combo'):
            return
        selected_category = self.filter_category_combo.get()
        columns = [col for col in self.all_columns if col != "学号"]
        if selected_category and selected_category not in ("", "全部"):
            columns = [col for col in columns if self.column_category_map.get(col) == selected_category]
        self.filter_column_combo['values'] = columns
        if columns:
            if self.filter_column_combo.get() not in columns:
                self.filter_column_combo.current(0)
                self.load_filter_values()
        else:
            self.filter_column_combo.set('')
            self.filter_value_combo.set('')
            self.filter_value_combo['values'] = []
    
    def on_filter_category_change(self, event=None):
        self.update_filter_column_options()
    
    def reset_category_filter(self):
        if hasattr(self, 'filter_category_combo'):
            self.filter_category_combo.set("全部")
            self.update_filter_column_options()
    
    def import_excel(self):
        """导入Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            
            if "学号" not in df.columns:
                messagebox.showerror("错误", "Excel文件中必须包含'学号'列！")
                return
            
            df = df.dropna(subset=["学号"])
            
            if len(df) == 0:
                messagebox.showwarning("警告", "没有有效的学生数据（学号列为空）")
                return
            
            df["学号"] = df["学号"].astype(str)
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(students)")
            existing_columns = [row[1] for row in cursor.fetchall()]
            
            new_columns = []
            overlap_columns = []
            
            for col in df.columns:
                if col in existing_columns:
                    overlap_columns.append(col)
                else:
                    new_columns.append(col)
            
            if overlap_columns:
                overlap_msg = f"以下列已存在，导入将覆盖这些列的数据：\n"
                overlap_msg += ", ".join([c for c in overlap_columns if c != "学号"])
                if "学号" in overlap_columns:
                    overlap_msg += "\n\n注：学号列用于匹配学生，不会被覆盖"
                
                if not messagebox.askyesno("确认覆盖", overlap_msg + "\n\n是否继续？"):
                    conn.close()
                    return
            
            if new_columns:
                new_msg = f"以下列是新列，将添加到系统中：\n"
                new_msg += ", ".join(new_columns)
                
                if not messagebox.askyesno("确认添加", new_msg + "\n\n是否继续？"):
                    conn.close()
                    return
            
            existing_students = []
            for student_id in df["学号"]:
                cursor.execute("SELECT 学号 FROM students WHERE 学号 = ?", (student_id,))
                if cursor.fetchone():
                    existing_students.append(student_id)
            
            if existing_students:
                msg = f"发现 {len(existing_students)} 个学生的数据已存在，将被覆盖：\n"
                msg += ", ".join(existing_students[:10])
                if len(existing_students) > 10:
                    msg += f"\n...等共{len(existing_students)}个学生"
                msg += "\n\n是否继续？"
                
                if not messagebox.askyesno("确认覆盖学生数据", msg):
                    conn.close()
                    return
            
            for col in df.columns:
                try:
                    cursor.execute(f'ALTER TABLE students ADD COLUMN "{col}" TEXT')
                    cursor.execute('''
                        INSERT OR IGNORE INTO column_history (column_name, first_upload_time)
                        VALUES (?, ?)
                    ''', (col, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                except sqlite3.OperationalError:
                    pass
            
            for _, row in df.iterrows():
                student_id = row["学号"]
                
                cursor.execute("SELECT 学号 FROM students WHERE 学号 = ?", (student_id,))
                exists = cursor.fetchone()
                
                if exists:
                    set_clause = ", ".join([f'"{col}" = ?' for col in df.columns if col != "学号"])
                    values = [self.normalize_cell_value(row[col]) for col in df.columns if col != "学号"]
                    values.append(student_id)
                    
                    cursor.execute(f'UPDATE students SET {set_clause} WHERE 学号 = ?', values)
                else:
                    columns = ", ".join([f'"{col}"' for col in df.columns])
                    placeholders = ", ".join(["?" for _ in df.columns])
                    values = [self.normalize_cell_value(row[col]) for col in df.columns]
                    
                    cursor.execute(f'INSERT INTO students ({columns}) VALUES ({placeholders})', values)
            
            conn.commit()
            conn.close()
            
            self.refresh_column_list()
            self.refresh_display()
            
            messagebox.showinfo("成功", f"成功导入 {len(df)} 条学生记录！")
            self.status_bar.config(text=f"导入完成: {len(df)} 条记录")
            
        except Exception as e:
            messagebox.showerror("错误", f"导入失败：\n{str(e)}\n\n{traceback.format_exc()}")
    
    def delete_row(self):
        """删除选中的学生记录"""
        selected_items = self.tree.selection()
        
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要删除的学生记录")
            return
        
        student_ids = []
        for item in selected_items:
            values = self.tree.item(item)['values']
            if values:
                student_ids.append(str(values[0]))
        
        if not student_ids:
            return
        
        msg = f"确定要删除以下 {len(student_ids)} 个学生的记录吗？\n\n"
        msg += "学号: " + ", ".join(student_ids[:10])
        if len(student_ids) > 10:
            msg += f"\n...等共{len(student_ids)}个学生"
        msg += "\n\n此操作不可恢复！"
        
        if not messagebox.askyesno("确认删除", msg):
            return
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            for student_id in student_ids:
                cursor.execute("DELETE FROM students WHERE 学号 = ?", (student_id,))
            
            conn.commit()
            conn.close()
            
            self.refresh_display()
            
            messagebox.showinfo("成功", f"成功删除 {len(student_ids)} 条学生记录")
            self.status_bar.config(text=f"删除完成: {len(student_ids)} 条记录")
            
        except Exception as e:
            messagebox.showerror("错误", f"删除失败：{str(e)}")
    
    def delete_column(self):
        """删除指定的列"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall() if row[1] != "学号"]
            conn.close()
            
            columns = self.sort_columns(columns)
            
            if not columns:
                messagebox.showwarning("警告", "没有可删除的列（学号列不能删除）")
                return
            
            delete_window = tk.Toplevel(self.root)
            delete_window.withdraw()
            delete_window.title("选择要删除的列")
            delete_window.geometry("450x550")
            delete_window.transient(self.root)
            delete_window.grab_set()
            
            title_frame = tk.Frame(delete_window)
            title_frame.pack(fill=tk.X, pady=10)
            
            tk.Label(title_frame, text="请选择要删除的列：", 
                    font=("Arial", 11, "bold"), fg="#F44336").pack()
            
            tk.Label(title_frame, text="警告：删除列将永久删除该列的所有数据！", 
                    font=("Arial", 9), fg="red").pack(pady=5)
            
            middle_frame = tk.Frame(delete_window)
            middle_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            canvas = tk.Canvas(middle_frame, bg="white")
            scrollbar = tk.Scrollbar(middle_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg="white")
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            check_vars = {}
            
            for col in columns:
                var = tk.BooleanVar(value=False)
                check_vars[col] = var
                cb = tk.Checkbutton(scrollable_frame, text=col, variable=var, 
                                   font=("Arial", 10), bg="white")
                cb.pack(anchor=tk.W, padx=20, pady=3)
            
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            button_frame = tk.Frame(delete_window)
            button_frame.pack(side=tk.BOTTOM, pady=15)
            
            def do_delete():
                selected_columns = [col for col, var in check_vars.items() if var.get()]
                
                if not selected_columns:
                    messagebox.showwarning("警告", "请至少选择一列")
                    return
                
                confirm_msg = f"确定要删除以下列吗？\n\n"
                confirm_msg += ", ".join(selected_columns)
                confirm_msg += "\n\n此操作将永久删除这些列的所有数据，不可恢复！"
                
                if not messagebox.askyesno("最终确认", confirm_msg):
                    return
                
                try:
                    conn_in = sqlite3.connect(self.db_path)
                    cursor_in = conn_in.cursor()
                    
                    cursor_in.execute("PRAGMA table_info(students)")
                    all_columns = [row[1] for row in cursor_in.fetchall()]
                    keep_columns = [col for col in all_columns if col not in selected_columns]
                    
                    columns_def = ", ".join([f'"{col}" TEXT' for col in keep_columns])
                    cursor_in.execute(f'CREATE TABLE students_new ({columns_def}, PRIMARY KEY (学号))')
                    
                    columns_str = ", ".join([f'"{col}"' for col in keep_columns])
                    cursor_in.execute(f'INSERT INTO students_new SELECT {columns_str} FROM students')
                    
                    cursor_in.execute('DROP TABLE students')
                    cursor_in.execute('ALTER TABLE students_new RENAME TO students')
                    
                    for col in selected_columns:
                        cursor_in.execute('DELETE FROM column_history WHERE column_name = ?', (col,))
                        cursor_in.execute('DELETE FROM column_category_map WHERE column_name = ?', (col,))
                    
                    conn_in.commit()
                    conn_in.close()
                    
                    self.refresh_column_list()
                    self.refresh_display()
                    
                    messagebox.showinfo("成功", f"成功删除 {len(selected_columns)} 列")
                    delete_window.destroy()
                    
                except Exception as err:
                    messagebox.showerror("错误", f"删除列失败：\n{str(err)}")
            
            tk.Button(button_frame, text="确定删除", command=do_delete,
                     bg="#F44336", fg="white", font=("Arial", 11, "bold"), 
                     padx=30, pady=8).pack(side=tk.LEFT, padx=10)
            tk.Button(button_frame, text="取消", command=delete_window.destroy,
                     bg="#9E9E9E", fg="white", font=("Arial", 11), 
                     padx=30, pady=8).pack(side=tk.LEFT, padx=10)
            
            center_window_relative(delete_window, self.root)
            
        except Exception as e:
            messagebox.showerror("错误", f"删除列功能错误：\n{str(e)}")
    
    def export_excel(self):
        """导出Excel文件"""
        try:
            columns = self.refresh_column_list()
            
            if not columns or columns == ["学号"]:
                messagebox.showwarning("警告", "没有可导出的数据")
                return
            
            export_type = "all"
            
            if self.is_filtered:
                msg = f"当前显示的是筛选后的数据（{self.current_filter_desc}）\n\n"
                msg += "请选择导出范围：\n\n"
                msg += "• 点击【是】：仅导出当前筛选的学生\n"
                msg += "• 点击【否】：导出所有学生数据"
                
                result = messagebox.askyesnocancel("选择导出范围", msg)
                if result is None:
                    return
                elif result:
                    export_type = "filtered"
                else:
                    export_type = "all"
            
            export_window = tk.Toplevel(self.root)
            export_window.withdraw()
            export_window.title("选择导出列")
            export_window.geometry("450x650")
            export_window.transient(self.root)
            export_window.grab_set()
            
            title_frame = tk.Frame(export_window)
            title_frame.pack(fill=tk.X, pady=10)
            
            title_text = "请选择要导出的列："
            if export_type == "filtered":
                title_text += f"\n（将导出筛选结果：{self.current_filter_desc}）"
            else:
                title_text += "\n（将导出所有学生数据）"
            
            tk.Label(title_frame, text=title_text, 
                    font=("Arial", 11, "bold"), fg="#2196F3").pack()
            
            middle_frame = tk.Frame(export_window)
            middle_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            canvas = tk.Canvas(middle_frame, bg="white")
            scrollbar = tk.Scrollbar(middle_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg="white")
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            check_vars = {}
            
            default_selected = ["学号", "姓名", "身份证号"]
            
            var = tk.BooleanVar(value=True)
            check_vars["学号"] = var
            cb = tk.Checkbutton(scrollable_frame, text="学号 ✓", variable=var, 
                               font=("Arial", 10, "bold"), state="disabled", 
                               fg="#4CAF50", bg="white")
            cb.pack(anchor=tk.W, padx=20, pady=3)
            
            for col in columns:
                if col != "学号":
                    should_select = col in default_selected
                    var = tk.BooleanVar(value=should_select)
                    check_vars[col] = var
                    if should_select:
                        cb = tk.Checkbutton(scrollable_frame, text=f"{col} (推荐)", variable=var,
                                           font=("Arial", 10), fg="#2196F3", bg="white")
                    else:
                        cb = tk.Checkbutton(scrollable_frame, text=col, variable=var,
                                           font=("Arial", 10), bg="white")
                    cb.pack(anchor=tk.W, padx=20, pady=3)
            
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            button_frame = tk.Frame(export_window)
            button_frame.pack(side=tk.BOTTOM, pady=15)
            
            def do_export():
                try:
                    selected_columns = [col for col, var in check_vars.items() if var.get()]
                    
                    if not selected_columns:
                        messagebox.showwarning("警告", "请至少选择一列")
                        return
                    
                    default_filename = "学生数据_筛选结果.xlsx" if export_type == "filtered" else "学生数据_全部.xlsx"
                    file_path = filedialog.asksaveasfilename(
                        title="保存Excel文件",
                        defaultextension=".xlsx",
                        initialfile=default_filename,
                        filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
                    )
                    
                    if not file_path:
                        return
                    
                    conn = sqlite3.connect(self.db_path)
                    columns_str = ", ".join([f'"{col}"' for col in selected_columns])
                    
                    if export_type == "filtered":
                        data = []
                        all_cols = [col[1] for col in conn.execute("PRAGMA table_info(students)").fetchall()]
                        ordered_all = self.sort_columns(all_cols)
                        for item in self.tree.get_children():
                            values = self.tree.item(item)['values']
                            row_dict = dict(zip(ordered_all, values))
                            filtered_row = [row_dict.get(col, '') for col in selected_columns]
                            data.append(filtered_row)
                        
                        df = pd.DataFrame(data, columns=selected_columns)
                    else:
                        query = f"SELECT {columns_str} FROM students"
                        df = pd.read_sql_query(query, conn)
                    
                    conn.close()
                    
                    df.to_excel(file_path, index=False, engine='openpyxl')
                    
                    export_desc = f"筛选结果：{self.current_filter_desc}" if export_type == "filtered" else "全部数据"
                    messagebox.showinfo("成功", f"成功导出 {len(df)} 条记录（{export_desc}）到：\n{file_path}")
                    export_window.destroy()
                    self.status_bar.config(text=f"导出完成: {len(df)} 条记录")
                    
                except Exception as e:
                    messagebox.showerror("错误", f"导出失败：\n{str(e)}")
            
            tk.Button(button_frame, text="确定导出", command=do_export,
                     bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), 
                     padx=40, pady=8).pack(side=tk.LEFT, padx=10)
            tk.Button(button_frame, text="取消", command=export_window.destroy,
                     bg="#F44336", fg="white", font=("Arial", 11), 
                     padx=40, pady=8).pack(side=tk.LEFT, padx=10)
            
            center_window_relative(export_window, self.root)
        
        except Exception as e:
            messagebox.showerror("错误", f"导出功能错误：\n{str(e)}")
    
    def refresh_display(self):
        """刷新显示所有数据"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall()]
            columns = self.sort_columns(columns)
            
            if not columns:
                conn.close()
                self.status_bar.config(text="暂无数据")
                self.is_filtered = False
                return
            
            self.tree['columns'] = columns
            self.tree['show'] = 'headings'
            
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor=tk.W)
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            columns_str = ", ".join([f'"{col}"' for col in columns])
            cursor.execute(f"SELECT {columns_str} FROM students")
            rows = cursor.fetchall()
            
            for row in rows:
                self.tree.insert('', tk.END, values=row)
            
            conn.close()
            self.status_bar.config(text=f"显示 {len(rows)} 条记录（全部数据）")
            self.is_filtered = False
            self.current_filter_desc = ""
            
        except Exception as e:
            messagebox.showerror("错误", f"刷新显示失败：{str(e)}")
    
    def search_by_name(self):
        """按姓名搜索"""
        name = self.name_entry.get().strip()
        
        if not name:
            messagebox.showwarning("警告", "请输入姓名")
            return
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall()]
            columns = self.sort_columns(columns)
            
            if "姓名" not in columns:
                conn.close()
                messagebox.showwarning("提示", "数据库中没有'姓名'列，无法按姓名搜索")
                return
            
            self.tree['columns'] = columns
            self.tree['show'] = 'headings'
            
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor=tk.W)
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            columns_str = ", ".join([f'"{col}"' for col in columns])
            cursor.execute(f'SELECT {columns_str} FROM students WHERE 姓名 LIKE ?', (f'%{name}%',))
            rows = cursor.fetchall()
            
            for row in rows:
                self.tree.insert('', tk.END, values=row)
            
            conn.close()
            
            if len(rows) == 0:
                messagebox.showinfo("提示", f"未找到姓名包含 '{name}' 的学生")
                self.is_filtered = False
                self.current_filter_desc = ""
            else:
                self.status_bar.config(text=f"找到 {len(rows)} 条匹配记录")
                self.is_filtered = True
                self.current_filter_desc = f"姓名包含'{name}'"
            
        except Exception as e:
            messagebox.showerror("错误", f"搜索失败：{str(e)}")
    
    def load_filter_values(self, event=None):
        column = self.filter_column_combo.get()
        
        if not column:
            return
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(f'SELECT DISTINCT "{column}" FROM students WHERE "{column}" IS NOT NULL AND "{column}" != ""')
            values = sorted([row[0] for row in cursor.fetchall()])
            
            conn.close()
            
            self.filter_value_combo['values'] = values
            
            if values:
                self.filter_value_combo.current(0)
            
        except Exception as e:
            messagebox.showerror("错误", f"加载筛选值失败：{str(e)}")
    
    def apply_filter(self):
        column = self.filter_column_combo.get()
        value = self.filter_value_combo.get()
        
        if not column or not value:
            messagebox.showwarning("警告", "请选择筛选列和筛选值")
            return
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall()]
            columns = self.sort_columns(columns)
            
            self.tree['columns'] = columns
            self.tree['show'] = 'headings'
            
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor=tk.W)
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            columns_str = ", ".join([f'"{col}"' for col in columns])
            cursor.execute(f'SELECT {columns_str} FROM students WHERE "{column}" = ?', (value,))
            rows = cursor.fetchall()
            
            for row in rows:
                self.tree.insert('', tk.END, values=row)
            
            conn.close()
            
            self.status_bar.config(text=f"筛选结果: {len(rows)} 条记录 ({column} = {value})")
            self.is_filtered = True
            self.current_filter_desc = f"{column}={value}"
            
        except Exception as e:
            messagebox.showerror("错误", f"筛选失败：{str(e)}")
    
    def clear_filters(self):
        """清空筛选条件并显示所有数据"""
        self.name_entry.delete(0, tk.END)
        self.filter_column_combo.set('')
        self.filter_value_combo.set('')
        self.filter_value_combo['values'] = []
        if hasattr(self, 'filter_category_combo'):
            self.filter_category_combo.set("全部")
            self.update_filter_column_options()
        self.is_filtered = False
        self.current_filter_desc = ""
        self.refresh_display()
    
    def on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        values = self.tree.item(item_id)['values']
        if not values:
            return
        columns = self.tree['columns']
        if "学号" in columns:
            student_idx = columns.index("学号")
            student_id = values[student_idx]
        else:
            student_id = values[0]
        if not student_id:
            return
        StudentDetailWindow(self, str(student_id))
    
    def open_column_category_manager(self):
        self.refresh_column_list()
        ColumnCategoryManager(self)


class ColumnCategoryManager:
    def __init__(self, parent):
        self.parent = parent
        self.window = tk.Toplevel(parent.root)
        self.window.withdraw()
        self.window.title("列分类管理")
        self.window.geometry("700x600")
        self.window.transient(parent.root)
        self.window.grab_set()
        
        self.column_vars = {}
        self.column_combos = []
        
        self.build_ui()
        self.reload_categories()
        self.load_columns()
        center_window_relative(self.window, self.parent.root)
    
    def build_ui(self):
        main = tk.Frame(self.window, padx=10, pady=10)
        main.pack(fill=tk.BOTH, expand=True)
        
        category_frame = tk.LabelFrame(main, text="一级类目", padx=10, pady=10)
        category_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.category_list = tk.Listbox(category_frame, height=5)
        self.category_list.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        btn_frame = tk.Frame(category_frame)
        btn_frame.pack(side=tk.LEFT, padx=10)
        
        tk.Button(btn_frame, text="新增", width=8, command=self.add_category).pack(pady=2)
        tk.Button(btn_frame, text="重命名", width=8, command=self.rename_category).pack(pady=2)
        tk.Button(btn_frame, text="删除", width=8, command=self.delete_category).pack(pady=2)
        
        mapping_frame = tk.LabelFrame(main, text="列归属设置（学号默认不分类）", padx=10, pady=10)
        mapping_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(mapping_frame)
        scrollbar = tk.Scrollbar(mapping_frame, orient="vertical", command=canvas.yview)
        self.form_frame = tk.Frame(canvas)
        
        self.form_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.form_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        tk.Button(main, text="保存映射", bg="#4CAF50", fg="white", font=("Arial", 11),
                  command=self.save_mappings).pack(fill=tk.X, pady=(10, 0))
    
    def reload_categories(self):
        self.categories = self.parent.get_all_categories()
        self.category_list.delete(0, tk.END)
        for cat in self.categories:
            self.category_list.insert(tk.END, cat)
        self.update_combo_options()
    
    def update_combo_options(self):
        options = ["未分类"] + self.categories if self.categories else ["未分类"]
        for combo in self.column_combos:
            combo['values'] = options
            if combo.get() not in options:
                combo.set("未分类")
    
    def load_columns(self):
        for widget in self.form_frame.winfo_children():
            widget.destroy()
        self.column_vars.clear()
        self.column_combos.clear()
        
        columns = [col for col in self.parent.all_columns if col != "学号"]
        columns = self.parent.sort_columns(columns)
        for idx, col in enumerate(columns):
            var = tk.StringVar(value=self.parent.column_category_map.get(col, "未分类"))
            lbl = tk.Label(self.form_frame, text=col + ":", anchor="w")
            lbl.grid(row=idx, column=0, sticky="w", padx=5, pady=4)
            combo = ttk.Combobox(self.form_frame, textvariable=var, state="readonly", width=25)
            combo.grid(row=idx, column=1, sticky="ew", padx=5, pady=4)
            self.form_frame.grid_columnconfigure(1, weight=1)
            self.column_vars[col] = var
            self.column_combos.append(combo)
        
        self.update_combo_options()
    
    def add_category(self):
        name = simpledialog.askstring("新增类目", "请输入类目名称：", parent=self.window)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name in self.categories:
            messagebox.showwarning("提示", "该类目已存在", parent=self.window)
            return
        conn = sqlite3.connect(self.parent.db_path)
        cursor = conn.cursor()
        cursor.execute(
            'INSERT INTO column_categories (category_name, created_at) VALUES (?, ?)',
            (name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()
        conn.close()
        self.reload_categories()
        self.parent.refresh_column_list()
    
    def rename_category(self):
        selection = self.category_list.curselection()
        if not selection:
            messagebox.showwarning("提示", "请选择要重命名的类目", parent=self.window)
            return
        old_name = self.category_list.get(selection[0])
        new_name = simpledialog.askstring("重命名", "新的类目名称：", parent=self.window, initialvalue=old_name)
        if not new_name:
            return
        new_name = new_name.strip()
        if not new_name or new_name == old_name:
            return
        if new_name in self.categories:
            messagebox.showwarning("提示", "该名称已存在", parent=self.window)
            return
        conn = sqlite3.connect(self.parent.db_path)
        cursor = conn.cursor()
        cursor.execute('UPDATE column_categories SET category_name = ? WHERE category_name = ?', (new_name, old_name))
        cursor.execute('UPDATE column_category_map SET category_name = ? WHERE category_name = ?', (new_name, old_name))
        conn.commit()
        conn.close()
        self.reload_categories()
        self.parent.refresh_column_list()
    
    def delete_category(self):
        selection = self.category_list.curselection()
        if not selection:
            messagebox.showwarning("提示", "请选择要删除的类目", parent=self.window)
            return
        name = self.category_list.get(selection[0])
        if not messagebox.askyesno("确认", f"确定删除类目“{name}”并清除其下列的归属吗？", parent=self.window):
            return
        conn = sqlite3.connect(self.parent.db_path)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM column_categories WHERE category_name = ?', (name,))
        cursor.execute('DELETE FROM column_category_map WHERE column_name IN (SELECT column_name FROM column_category_map WHERE category_name = ?)', (name,))
        cursor.execute('DELETE FROM column_category_map WHERE category_name = ?', (name,))
        conn.commit()
        conn.close()
        self.reload_categories()
        self.parent.refresh_column_list()
    
    def save_mappings(self):
        try:
            conn = sqlite3.connect(self.parent.db_path)
            cursor = conn.cursor()
            for col, var in self.column_vars.items():
                category = var.get()
                if not category or category == "未分类":
                    cursor.execute('DELETE FROM column_category_map WHERE column_name = ?', (col,))
                else:
                    cursor.execute('''
                        INSERT INTO column_category_map (column_name, category_name)
                        VALUES (?, ?)
                        ON CONFLICT(column_name) DO UPDATE SET category_name = excluded.category_name
                    ''', (col, category))
            conn.commit()
            conn.close()
            messagebox.showinfo("成功", "列分类设置已保存", parent=self.window)
            self.parent.refresh_column_list()
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}", parent=self.window)


class StudentDetailWindow:
    def __init__(self, parent, student_id):
        self.parent = parent
        self.student_id = student_id
        self.window = tk.Toplevel(parent.root)
        self.window.withdraw()
        self.window.title(f"学生详情 - {student_id}")
        self.window.geometry("600x700")
        self.window.transient(parent.root)
        self.window.grab_set()
        
        self.entry_vars = {}
        
        if not self.load_student_data():
            self.window.destroy()
            return
        
        self.column_groups = self.parent.get_columns_grouped_by_category(self.columns)
        self.build_form()
        center_window_relative(self.window, self.parent.root)
    
    def load_student_data(self):
        try:
            conn = sqlite3.connect(self.parent.db_path)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall()]
            columns = self.parent.sort_columns(columns)
            self.columns = columns
            columns_str = ", ".join([f'"{col}"' for col in self.columns])
            cursor.execute(f'SELECT {columns_str} FROM students WHERE 学号 = ?', (self.student_id,))
            row = cursor.fetchone()
            conn.close()
            if not row:
                messagebox.showwarning("提示", "该学生记录不存在", parent=self.window)
                return False
            self.data = dict(zip(self.columns, row))
            return True
        except Exception as e:
            messagebox.showerror("错误", f"加载数据失败：{str(e)}", parent=self.window)
            return False
    
    def build_form(self):
        container = tk.Frame(self.window)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        notebook = ttk.Notebook(container)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        basic_inner = self._create_scrollable_tab(notebook, "基础信息")
        self._add_entry_widget(basic_inner, 0, "学号", disabled=True)
        
        if not self.column_groups:
            placeholder = tk.Label(basic_inner, text="暂无其他可编辑字段", fg="#757575")
            placeholder.grid(row=1, column=0, columnspan=2, pady=10)
        else:
            for category, columns in self.column_groups:
                tab_inner = self._create_scrollable_tab(notebook, category)
                for idx, col in enumerate(columns):
                    self._add_entry_widget(tab_inner, idx, col)
        
        btn_frame = tk.Frame(self.window, pady=10)
        btn_frame.pack(fill=tk.X)
        
        tk.Button(btn_frame, text="保存", bg="#4CAF50", fg="white",
                  font=("Arial", 11), padx=30, pady=5,
                  command=self.save_changes).pack(side=tk.LEFT, expand=True, padx=10)
        tk.Button(btn_frame, text="关闭", bg="#9E9E9E", fg="white",
                  font=("Arial", 11), padx=30, pady=5,
                  command=self.window.destroy).pack(side=tk.LEFT, expand=True, padx=10)
    
    def _create_scrollable_tab(self, notebook, title):
        frame = tk.Frame(notebook)
        notebook.add(frame, text=title)
        canvas = tk.Canvas(frame)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        inner.grid_columnconfigure(1, weight=1)
        return inner
    
    def _add_entry_widget(self, parent_frame, row_index, column_name, disabled=False):
        tk.Label(parent_frame, text=f"{column_name}：", anchor="w").grid(row=row_index, column=0, sticky="w", padx=5, pady=4)
        value = '' if self.data.get(column_name) is None else str(self.data.get(column_name))
        var = tk.StringVar(value=value)
        entry = tk.Entry(parent_frame, textvariable=var, font=("Arial", 10))
        if disabled:
            entry.config(state="disabled")
        entry.grid(row=row_index, column=1, sticky="ew", padx=5, pady=4)
        self.entry_vars[column_name] = var
    
    def save_changes(self):
        try:
            conn = sqlite3.connect(self.parent.db_path)
            cursor = conn.cursor()
            update_columns = [col for col in self.columns if col != "学号"]
            set_clause = ", ".join([f'"{col}" = ?' for col in update_columns])
            values = []
            for col in update_columns:
                if col not in self.entry_vars:
                    values.append(None)
                    continue
                val = self.entry_vars[col].get().strip()
                values.append(val if val != "" else None)
            values.append(self.student_id)
            cursor.execute(f'UPDATE students SET {set_clause} WHERE 学号 = ?', values)
            conn.commit()
            conn.close()
            messagebox.showinfo("成功", "学生信息已更新", parent=self.window)
            self.parent.refresh_display()
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}", parent=self.window)


def main():
    """主函数"""
    try:
        root = tk.Tk()
        app = StudentManagementSystem(root)
        root.mainloop()
    except Exception as e:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("错误", f"程序运行出错：\n{str(e)}\n\n{traceback.format_exc()}")
        except:
            print(f"程序运行出错：{str(e)}")
            print(traceback.format_exc())
            input("按回车键退出...")
        sys.exit(1)

from datetime import datetime, date

DEADLINE = date(2025, 12, 31)  # 截止日期（含当天）

if datetime.now().date() > DEADLINE:
    messagebox.showerror("试用已过期", "该版本已过期，请联系开发者获取最新版。")
    sys.exit(0)

if __name__ == "__main__":
    main()