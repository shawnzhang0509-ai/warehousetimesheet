#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤工时解析工具 - 图形界面版
支持多打卡点合并、更改说明表、名字映射、9AM封顶规则

双击运行，无需命令行
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import os
import sys
import re
from datetime import time, datetime

try:
    import xlrd
    import pandas as pd
    import openpyxl
except ImportError:
    # 尝试自动安装
    import subprocess
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "xlrd", "pandas", "openpyxl"],
        capture_output=True, text=True
    )
    try:
        import xlrd
        import pandas as pd
        import openpyxl
    except ImportError:
        print("请手动安装依赖: python -m pip install xlrd pandas openpyxl")
        sys.exit(1)


# ============================================================
# 核心解析函数（与命令行版相同）
# ============================================================

def parse_time(val):
    if val is None or val == '':
        return None
    if isinstance(val, str):
        try:
            parts = val.strip().split(':')
            if len(parts) >= 2:
                return time(int(parts[0]), int(parts[1]), int(parts[2]) if len(parts) > 2 else 0)
        except:
            return None
    elif isinstance(val, (int, float)):
        total_seconds = int(val * 24 * 3600)
        return time(total_seconds // 3600, (total_seconds % 3600) // 60, total_seconds % 60)
    return None


def format_time(t):
    if t is None:
        return ''
    return t.strftime('%H:%M:%S')


def calc_time_diff(start, end):
    if start is None or end is None:
        return 0
    s = start.hour * 3600 + start.minute * 60 + start.second
    e = end.hour * 3600 + end.minute * 60 + end.second
    if e < s:
        e += 24 * 3600
    return round((e - s) / 3600, 2)


def get_adjusted_start(name, date_key, original_time, adj_map):
    if (name, date_key) in adj_map:
        return time_from_hour(adj_map[(name, date_key)])
    if (name, 'ALL') in adj_map:
        return time_from_hour(adj_map[(name, 'ALL')])
    if 'mandeep' in name.lower():
        return original_time
    if original_time is not None and original_time < time(9, 0, 0):
        return time(9, 0, 0)
    return original_time


def time_from_hour(hour_val):
    if hour_val is None:
        return None
    try:
        return time(int(float(hour_val)), 0, 0)
    except:
        return None


def extract_from_file(filepath):
    """从xls提取打卡记录"""
    if not filepath or not os.path.exists(filepath):
        return []
    book = xlrd.open_workbook(filepath)
    punch_sheets = [n for n in book.sheet_names() if n not in ['排班记录表', '考勤汇总表']]
    records = []
    for sheet_name in punch_sheets:
        sheet = book.sheet_by_name(sheet_name)
        for emp_idx in range(3):
            base_col = emp_idx * 15
            if base_col + 12 >= sheet.ncols:
                continue
            dept = sheet.cell_value(3, base_col + 1) if base_col + 1 < sheet.ncols else ''
            name = sheet.cell_value(3, base_col + 9) if base_col + 9 < sheet.ncols else ''
            emp_id = sheet.cell_value(4, base_col + 9) if base_col + 9 < sheet.ncols else ''
            if not name or not emp_id:
                continue
            try:
                emp_id = int(emp_id)
            except:
                continue
            dates = ['20', '21', '22', '23', '24', '25', '26']
            weekdays = ['一', '二', '三', '四', '五', '六', '日']
            for i, (day, wd) in enumerate(zip(dates, weekdays)):
                row = 12 + i
                if row >= sheet.nrows:
                    break
                am_s = parse_time(sheet.cell_value(row, base_col + 1))
                am_e = parse_time(sheet.cell_value(row, base_col + 3))
                pm_s = parse_time(sheet.cell_value(row, base_col + 6))
                pm_e = parse_time(sheet.cell_value(row, base_col + 8))
                ot_s = parse_time(sheet.cell_value(row, base_col + 10))
                ot_e = parse_time(sheet.cell_value(row, base_col + 12))
                records.append({
                    '员工号': emp_id, '姓名': str(name).strip(),
                    '部门': str(dept).strip() if dept else '',
                    '日期': f"4月{day}日", '星期': wd,
                    '上午上班': format_time(am_s), '上午下班': format_time(am_e),
                    '下午上班': format_time(pm_s), '下午下班': format_time(pm_e),
                    '加班签到': format_time(ot_s), '加班签退': format_time(ot_e),
                })
    return records


def parse_adjustment_table(filepath):
    """解析更改说明表"""
    if not filepath or not os.path.exists(filepath):
        return {}
    try:
        df = pd.read_excel(filepath, header=None)
    except:
        return {}
    header_row = None
    name_col = date_col = time_col = None
    for r in range(min(10, len(df))):
        row_vals = [str(v).strip() if pd.notna(v) else '' for v in df.iloc[r]]
        if '姓名' in row_vals and '未打卡时间' in row_vals:
            header_row = r
            for c, v in enumerate(row_vals):
                if v == '姓名': name_col = c
                elif v == '未打卡时间': time_col = c
                elif v == '日期': date_col = c
            break
    if header_row is None or name_col is None or time_col is None:
        return {}
    if date_col is None:
        date_col = name_col + 1
    adj_map = {}
    for r in range(header_row + 1, len(df)):
        name = df.iloc[r, name_col]
        time_val = df.iloc[r, time_col]
        date_val = df.iloc[r, date_col]
        if pd.isna(name) or pd.isna(time_val):
            continue
        try:
            start_hour = float(time_val)
        except:
            continue
        date_key = None
        if pd.notna(date_val):
            m = re.match(r'(\d{1,2})[/\.](\d{1,2})', str(date_val).strip())
            if m:
                d, mon = int(m.group(1)), int(m.group(2))
                date_key = f"{mon:02d}月{d:02d}日"
        if date_key:
            adj_map[(str(name).strip(), date_key)] = start_hour
    return adj_map


def merge_and_calc(records, adj_map=None):
    """合并多来源记录并计算工时"""
    if not records:
        return pd.DataFrame()
    
    # 按姓名+日期分组
    groups = {}
    for r in records:
        key = (r['姓名'], r['日期'])
        if key not in groups:
            groups[key] = []
        groups[key].append(r)
    
    results = []
    for (name, date_4m), rows in sorted(groups.items()):
        def to_time_obj(s):
            if not s: return None
            p = list(map(int, s.split(':')))
            return time(p[0], p[1], p[2])
        
        all_starts = [r[k] for r in rows for k in ['上午上班','下午上班','加班签到'] if r.get(k)]
        all_ends = [r[k] for r in rows for k in ['上午下班','下午下班','加班签退'] if r.get(k)]
        earliest = min(all_starts) if all_starts else ''
        latest = max(all_ends) if all_ends else ''
        
        orig_start = to_time_obj(earliest)
        end_time = to_time_obj(latest)
        
        if adj_map:
            date_key = date_4m.replace('4月', '04月')
            adj_start = get_adjusted_start(name, date_key, orig_start, adj_map)
        else:
            adj_start = orig_start
        
        total = calc_time_diff(adj_start, end_time)
        
        # 检查是否有加班
        ot_records = [r for r in rows if r.get('加班签到') and r.get('加班签退')]
        ot_hours = 0
        if ot_records:
            ot_s = to_time_obj(ot_records[0]['加班签到'])
            ot_e = to_time_obj(ot_records[0]['加班签退'])
            ot_hours = calc_time_diff(ot_s, ot_e)
            normal = max(0, total - ot_hours)
        else:
            normal = total
        
        # 调整标记
        adj_mark = ''
        if adj_map:
            date_key = date_4m.replace('4月', '04月')
            if (name, date_key) in adj_map or (name, 'ALL') in adj_map:
                adj_mark = '已调整'
            elif orig_start and orig_start < time(9, 0, 0) and 'mandeep' not in name.lower():
                adj_mark = '9AM封顶'
        
        results.append({
            '姓名': name, '部门': rows[0]['部门'], '日期': date_4m, '星期': rows[0]['星期'],
            '上班打卡': format_time(adj_start), '下班打卡': format_time(end_time),
            '工作工时': round(normal, 2), '加班工时': round(ot_hours, 2),
            '合计工时': round(total, 2), '调整标记': adj_mark,
        })
    
    return pd.DataFrame(results)


# ============================================================
# GUI 界面
# ============================================================

class AttendanceGUI:
    @staticmethod
    def _get_safe_output_dir():
        """获取安全的默认输出目录，避免System32等系统目录"""
        cwd = os.getcwd()
        dangerous = ['system32', 'syswow64', 'windows', 'program files']
        is_dangerous = any(d in cwd.lower() for d in dangerous)
        
        if is_dangerous or not os.access(cwd, os.W_OK):
            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            if os.path.exists(desktop) and os.access(desktop, os.W_OK):
                return desktop
            docs = os.path.join(os.path.expanduser('~'), 'Documents')
            if os.path.exists(docs) and os.access(docs, os.W_OK):
                return docs
            return os.path.expanduser('~')
        return cwd

    def __init__(self, root):
        self.root = root
        self.root.title("考勤工时解析工具")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # 变量
        self.walls_path = tk.StringVar()
        self.carine_path = tk.StringVar()
        self.adjust_path = tk.StringVar()
        self.mapping_path = tk.StringVar(value="name_mapping.json")
        # 设置合理的默认输出目录（避免System32等系统目录）
        default_dir = self._get_safe_output_dir()
        self.output_dir = tk.StringVar(value=default_dir)
        
        self.name_mapping = {}  # 打卡名 -> {first, last}
        self.parsed_data = None  # 解析后的数据
        
        self._build_ui()
    
    def _build_ui(self):
        # 主容器
        main = ttk.Frame(self.root, padding="10")
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        
        # === 文件选择区域 ===
        file_frame = ttk.LabelFrame(main, text="文件选择", padding="10")
        file_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # Walls文件
        ttk.Label(file_frame, text="Walls打卡文件:").grid(row=0, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.walls_path).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(file_frame, text="浏览", command=lambda: self._browse_file(self.walls_path, "xls")).grid(row=0, column=2)
        
        # Carine文件
        ttk.Label(file_frame, text="Carine打卡文件:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(file_frame, textvariable=self.carine_path).grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(file_frame, text="浏览", command=lambda: self._browse_file(self.carine_path, "xls")).grid(row=1, column=2)
        
        # 更改说明表
        ttk.Label(file_frame, text="更改说明表:").grid(row=2, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.adjust_path).grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(file_frame, text="浏览", command=lambda: self._browse_file(self.adjust_path, "xlsx")).grid(row=2, column=2)
        
        # 名字映射
        ttk.Label(file_frame, text="名字映射文件:").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(file_frame, textvariable=self.mapping_path).grid(row=3, column=1, sticky="ew", padx=5)
        ttk.Button(file_frame, text="浏览", command=lambda: self._browse_file(self.mapping_path, "json")).grid(row=3, column=2)
        
        # 输出目录
        ttk.Label(file_frame, text="输出目录:").grid(row=4, column=0, sticky="w")
        ttk.Entry(file_frame, textvariable=self.output_dir).grid(row=4, column=1, sticky="ew", padx=5)
        ttk.Button(file_frame, text="浏览", command=self._browse_dir).grid(row=4, column=2)
        
        # === 操作按钮区域 ===
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=10)
        
        ttk.Button(btn_frame, text="1. 提取名字", command=self._extract_names, width=15).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="2. 解析工时", command=self._parse_attendance, width=15).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="3. 导出结果", command=self._export_results, width=15).grid(row=0, column=2, padx=5)
        
        # === 名字映射编辑区域 ===
        map_frame = ttk.LabelFrame(main, text="名字映射 (打卡名 → Legal名字)", padding="10")
        map_frame.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=(0, 10))
        map_frame.columnconfigure(0, weight=1)
        map_frame.rowconfigure(1, weight=1)
        
        ttk.Label(map_frame, text="提示: 左侧是打卡机里的名字，右侧填写对应的 Legal First Name 和 Last Name").grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        # Treeview
        columns = ('punch_name', 'first_name', 'last_name')
        self.tree = ttk.Treeview(map_frame, columns=columns, show='headings', height=10)
        self.tree.heading('punch_name', text='打卡机名字')
        self.tree.heading('first_name', text='Legal First Name')
        self.tree.heading('last_name', text='Legal Last Name')
        self.tree.column('punch_name', width=200)
        self.tree.column('first_name', width=250)
        self.tree.column('last_name', width=250)
        
        vsb = ttk.Scrollbar(map_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(map_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")
        
        # 编辑按钮
        edit_frame = ttk.Frame(map_frame)
        edit_frame.grid(row=3, column=0, columnspan=2, pady=5)
        ttk.Button(edit_frame, text="编辑选中行", command=self._edit_selected).grid(row=0, column=0, padx=5)
        ttk.Button(edit_frame, text="保存映射", command=self._save_mapping).grid(row=0, column=1, padx=5)
        ttk.Button(edit_frame, text="加载映射", command=self._load_mapping).grid(row=0, column=2, padx=5)
        
        # 双击编辑
        self.tree.bind('<Double-1>', lambda e: self._edit_selected())
        
        # === 结果显示区域 ===
        result_frame = ttk.LabelFrame(main, text="解析结果预览", padding="10")
        result_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(0, 10))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        self.result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, height=8)
        self.result_text.grid(row=0, column=0, sticky="nsew")
        
        main.rowconfigure(2, weight=2)
        main.rowconfigure(3, weight=1)
    
    def _browse_file(self, var, ext):
        if ext == "json":
            f = filedialog.askopenfilename(filetypes=[("JSON", "*.json"), ("All", "*.*")])
        elif ext == "xlsx":
            f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
        else:
            f = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls"), ("All", "*.*")])
        if f:
            var.set(f)
    
    def _browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.output_dir.set(d)
    
    def _extract_names(self):
        """从打卡文件中提取所有名字"""
        files = []
        if self.walls_path.get(): files.append(self.walls_path.get())
        if self.carine_path.get(): files.append(self.carine_path.get())
        
        if not files:
            messagebox.showwarning("提示", "请先选择至少一个打卡文件")
            return
        
        all_names = set()
        for f in files:
            records = extract_from_file(f)
            for r in records:
                all_names.add(r['姓名'])
        
        # 填充到Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for name in sorted(all_names):
            if name in self.name_mapping:
                mapped = self.name_mapping[name]
                self.tree.insert('', 'end', values=(name, mapped.get('first_name', ''), mapped.get('last_name', '')))
            else:
                # 英文名自动拆分
                if all(c.isalpha() or c.isspace() for c in name):
                    parts = name.split()
                    if len(parts) >= 2:
                        self.tree.insert('', 'end', values=(name, parts[0], ' '.join(parts[1:])))
                    else:
                        self.tree.insert('', 'end', values=(name, name, ''))
                else:
                    self.tree.insert('', 'end', values=(name, '', ''))
        
        self.result_text.insert(tk.END, f"✅ 已提取 {len(all_names)} 个名字，请填写 Legal First/Last Name\n")
    
    def _edit_selected(self):
        """编辑选中的名字映射"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一行")
            return
        
        item = self.tree.item(selected[0])
        punch_name, first, last = item['values']
        
        # 弹出编辑窗口
        edit_win = tk.Toplevel(self.root)
        edit_win.title(f"编辑: {punch_name}")
        edit_win.geometry("400x150")
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        ttk.Label(edit_win, text="Legal First Name:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        first_var = tk.StringVar(value=first)
        ttk.Entry(edit_win, textvariable=first_var, width=30).grid(row=0, column=1, padx=10, pady=5)
        
        ttk.Label(edit_win, text="Legal Last Name:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        last_var = tk.StringVar(value=last)
        ttk.Entry(edit_win, textvariable=last_var, width=30).grid(row=1, column=1, padx=10, pady=5)
        
        def save():
            self.tree.item(selected[0], values=(punch_name, first_var.get(), last_var.get()))
            edit_win.destroy()
        
        ttk.Button(edit_win, text="保存", command=save).grid(row=2, column=0, columnspan=2, pady=10)
    
    def _save_mapping(self):
        """保存名字映射到JSON"""
        mapping = {}
        for item in self.tree.get_children():
            punch, first, last = self.tree.item(item)['values']
            mapping[punch] = {"first_name": first, "last_name": last}
        
        path = os.path.join(self.output_dir.get(), self.mapping_path.get())
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
        
        self.name_mapping = mapping
        self.result_text.insert(tk.END, f"✅ 名字映射已保存: {path}\n")
    
    def _load_mapping(self):
        """从JSON加载名字映射"""
        path = os.path.join(self.output_dir.get(), self.mapping_path.get())
        if not os.path.exists(path):
            messagebox.showwarning("提示", f"找不到文件: {path}")
            return
        
        with open(path, 'r', encoding='utf-8') as f:
            self.name_mapping = json.load(f)
        
        # 刷新Treeview
        for item in self.tree.get_children():
            punch = self.tree.item(item)['values'][0]
            if punch in self.name_mapping:
                m = self.name_mapping[punch]
                self.tree.item(item, values=(punch, m.get('first_name', ''), m.get('last_name', '')))
        
        self.result_text.insert(tk.END, f"✅ 已加载名字映射: {len(self.name_mapping)} 条\n")
    
    def _parse_attendance(self):
        """解析考勤数据"""
        files = []
        if self.walls_path.get(): files.append(self.walls_path.get())
        if self.carine_path.get(): files.append(self.carine_path.get())
        
        if not files:
            messagebox.showwarning("提示", "请先选择至少一个打卡文件")
            return
        
        self.result_text.insert(tk.END, "⏳ 正在解析打卡数据...\n")
        self.root.update()
        
        # 提取所有记录
        all_records = []
        for f in files:
            records = extract_from_file(f)
            all_records.extend(records)
        
        if not all_records:
            messagebox.showerror("错误", "没有提取到任何打卡记录")
            return
        
        # 读取更改说明表
        adj_map = {}
        if self.adjust_path.get() and os.path.exists(self.adjust_path.get()):
            adj_map = parse_adjustment_table(self.adjust_path.get())
            self.result_text.insert(tk.END, f"📋 更改说明表: {len(adj_map)} 条规则\n")
        
        # 合并计算
        self.parsed_data = merge_and_calc(all_records, adj_map)
        
        # 显示统计
        total_emp = self.parsed_data['姓名'].nunique()
        valid_emp = self.parsed_data[self.parsed_data['合计工时'] > 0]['姓名'].nunique()
        adj_count = (self.parsed_data['调整标记'] == '已调整').sum()
        cap9_count = (self.parsed_data['调整标记'] == '9AM封顶').sum()
        
        self.result_text.insert(tk.END, f"✅ 解析完成!\n")
        self.result_text.insert(tk.END, f"   总员工: {total_emp} 人\n")
        self.result_text.insert(tk.END, f"   有打卡: {valid_emp} 人\n")
        self.result_text.insert(tk.END, f"   总记录: {len(self.parsed_data)} 条\n")
        if adj_count > 0:
            self.result_text.insert(tk.END, f"   特殊调整: {adj_count} 条\n")
        if cap9_count > 0:
            self.result_text.insert(tk.END, f"   9AM封顶: {cap9_count} 条\n")
        
        # 显示前10行预览
        self.result_text.insert(tk.END, f"\n📋 前10条记录预览:\n")
        preview = self.parsed_data[self.parsed_data['合计工时'] > 0].head(10)
        for _, r in preview.iterrows():
            self.result_text.insert(tk.END, f"   {r['姓名']:12s} {r['日期']:8s} {r['上班打卡']:10s} -> {r['下班打卡']:10s} = {r['合计工时']:.2f}h {r['调整标记']}\n")
    
    def _export_results(self):
        """导出timesheet.csv和详细xlsx"""
        if self.parsed_data is None or self.parsed_data.empty:
            messagebox.showwarning("提示", "请先点击'解析工时'")
            return
        
        # 收集当前的名字映射
        mapping = {}
        for item in self.tree.get_children():
            punch, first, last = self.tree.item(item)['values']
            mapping[punch] = {"first_name": first, "last_name": last}
        
        out_dir = self.output_dir.get()
        
        # 生成timesheet.csv
        csv_rows = []
        for _, r in self.parsed_data.iterrows():
            punch_name = r['姓名']
            if punch_name in mapping and mapping[punch_name].get('last_name'):
                first = mapping[punch_name].get('first_name', '')
                last = mapping[punch_name].get('last_name', '')
            else:
                first = punch_name
                last = ''
            
            m = re.match(r'(\d+)月(\d+)日', r['日期'])
            if m:
                month, day = int(m.group(1)), int(m.group(2))
                date_str = f"{month}/{day}/{datetime.now().year}"
            else:
                date_str = r['日期']
            
            csv_rows.append({
                'first_name': first, 'last_name': last,
                'type': 'Ordinary Time', 'Date': date_str,
                'end_date': '', 'description': '',
                'hours': round(r['合计工时'], 2), 'rate': ''
            })
        
        df_csv = pd.DataFrame(csv_rows)
        csv_path = os.path.join(out_dir, 'timesheet.csv')
        df_csv.to_csv(csv_path, index=False)
        
        # 生成详细xlsx
        xlsx_path = os.path.join(out_dir, '工时明细.xlsx')
        pivot = self.parsed_data.pivot_table(
            index=['姓名', '部门'], columns='日期', values='合计工时', aggfunc='first'
        ).reset_index()
        weekly = self.parsed_data.groupby(['姓名', '部门']).agg({
            '工作工时': 'sum', '加班工时': 'sum', '合计工时': 'sum', '日期': 'count'
        }).reset_index()
        weekly.columns = ['姓名', '部门', '本周工作工时', '本周加班工时', '本周合计工时', '打卡天数']
        pivot = pivot.merge(weekly, on=['姓名', '部门'], how='left')
        
        with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
            self.parsed_data.to_excel(writer, sheet_name='每日打卡明细', index=False)
            self.parsed_data[self.parsed_data['合计工时'] > 0].to_excel(writer, sheet_name='有打卡记录', index=False)
            pivot.to_excel(writer, sheet_name='每周汇总', index=False)
        
        self.result_text.insert(tk.END, f"\n✅ 导出完成!\n")
        self.result_text.insert(tk.END, f"   Timesheet CSV: {csv_path}\n")
        self.result_text.insert(tk.END, f"   详细报表: {xlsx_path}\n")
        
        messagebox.showinfo("完成", f"文件已导出到:\n{csv_path}\n{xlsx_path}")


def main():
    root = tk.Tk()
    app = AttendanceGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
