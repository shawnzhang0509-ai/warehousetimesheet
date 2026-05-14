#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤工时解析程序
支持多文件合并（仓库打卡 + 公司打卡），自动计算每日工时
支持考勤更改说明表（特殊起始时间 + 9AM封顶规则）
输出 timesheet.csv（工资系统格式）+ 详细Excel报表

用法:
    python attendance_parser.py 文件1.xls [文件2.xls ...] -o 输出.xlsx
    python attendance_parser.py walls.xls carine.xls -a 更改表.xlsx -m 名字映射.json

示例:
    python attendance_parser.py 20.04-26.04_walls.xls 20.04-26.04_carine.xls \
        -a 考勤更改说明表.xlsx -m name_mapping.json --sources walls carine -o 工时汇总.xlsx
"""

import argparse
import sys
import os
import re
import json
from datetime import time, datetime

try:
    import xlrd
except ImportError:
    print("错误: 需要安装 xlrd 库 (pip install xlrd)")
    sys.exit(1)

try:
    import pandas as pd
    import openpyxl
except ImportError:
    print("错误: 需要安装 pandas 和 openpyxl (pip install pandas openpyxl)")
    sys.exit(1)


# ============================================================
# 时间解析与计算工具
# ============================================================

def parse_time(val):
    """智能解析时间值：支持xlrd数值时间和字符串时间"""
    if val is None or val == '':
        return None
    if isinstance(val, str):
        try:
            parts = val.strip().split(':')
            if len(parts) >= 2:
                return time(int(parts[0]), int(parts[1]), int(parts[2]) if len(parts) > 2 else 0)
        except (ValueError, IndexError):
            return None
    elif isinstance(val, (int, float)):
        total_seconds = int(val * 24 * 3600)
        return time(total_seconds // 3600, (total_seconds % 3600) // 60, total_seconds % 60)
    return None


def format_time(t):
    """time对象 → HH:MM:SS 字符串"""
    if t is None:
        return ''
    return t.strftime('%H:%M:%S')


def time_from_hour(hour_val):
    """数值小时 → time对象，如 7.0 → 07:00:00"""
    if hour_val is None:
        return None
    try:
        h = int(float(hour_val))
        return time(h, 0, 0)
    except (ValueError, TypeError):
        return None


def calc_time_diff(start, end):
    """计算两个time的差值（小时），支持跨天"""
    if start is None or end is None:
        return 0
    s = start.hour * 3600 + start.minute * 60 + start.second
    e = end.hour * 3600 + end.minute * 60 + end.second
    if e < s:
        e += 24 * 3600
    return round((e - s) / 3600, 2)


# ============================================================
# 考勤更改说明表解析
# ============================================================

def parse_adjustment_table(filepath):
    """解析考勤更改说明表，返回 (姓名, 日期) → 起始时间(小时) 的字典"""
    if not filepath or not os.path.exists(filepath):
        return {}

    try:
        df = pd.read_excel(filepath, header=None)
    except Exception as e:
        print(f"警告: 无法读取更改说明表: {e}")
        return {}

    # 找到表头行
    header_row = None
    name_col = None
    date_col = None
    time_col = None

    for r in range(min(10, len(df))):
        row_vals = [str(v).strip() if pd.notna(v) else '' for v in df.iloc[r]]
        if '姓名' in row_vals and '未打卡时间' in row_vals:
            header_row = r
            for c, v in enumerate(row_vals):
                if v == '姓名':
                    name_col = c
                elif v == '未打卡时间':
                    time_col = c
                elif v == '日期':
                    date_col = c
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

        name = str(name).strip()
        try:
            start_hour = float(time_val)
        except (ValueError, TypeError):
            continue

        # 解析日期
        date_key = None
        if pd.notna(date_val):
            date_str = str(date_val).strip()
            m = re.match(r'(\d{1,2})[/\.](\d{1,2})', date_str)
            if m:
                d, mon = int(m.group(1)), int(m.group(2))
                date_key = f"{mon:02d}月{d:02d}日"

        if date_key:
            adj_map[(name, date_key)] = start_hour

    print(f"  更改说明表: 读取了 {len(adj_map)} 条规则")
    return adj_map


# ============================================================
# 名字映射工具
# ============================================================

def load_name_mapping(filepath):
    """加载名字映射JSON，返回 打卡机名 → {first_name, last_name} 的字典"""
    if not filepath or not os.path.exists(filepath):
        return {}
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            mapping = json.load(f)
        print(f"  名字映射: 读取了 {len(mapping)} 条")
        return mapping
    except Exception as e:
        print(f"警告: 无法读取名字映射文件: {e}")
        return {}


def generate_name_mapping_template(all_names, output_path='name_mapping.json'):
    """根据所有打卡机名字生成映射模板JSON"""
    mapping = {}
    for name in sorted(all_names):
        name_str = str(name).strip()
        if all(c.isalpha() or c.isspace() for c in name_str):
            parts = name_str.split()
            if len(parts) >= 2:
                mapping[name_str] = {"first_name": parts[0], "last_name": " ".join(parts[1:])}
            else:
                mapping[name_str] = {"first_name": name_str, "last_name": ""}
        else:
            mapping[name_str] = {"first_name": "", "last_name": ""}

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    print(f"\n📝 名字映射模板已生成: {output_path}")
    print("   请打开JSON文件，把空的名字填写为正式的 first_name + last_name")
    return mapping


def normalize_text_key(value):
    """姓名等文本合并键：忽略大小写和多余空格。"""
    return re.sub(r'\s+', ' ', str(value or '').strip()).casefold()


def normalize_employee_id(value):
    """员工号合并键：把 Excel 常见的 1001.0 形式归一成 1001。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    text = str(value).strip()
    if not text or text.lower() == 'nan':
        return ''
    try:
        num = float(text)
        if num.is_integer():
            return str(int(num))
    except (ValueError, TypeError):
        pass
    return text


def mapped_name_for_merge(name, name_mapping=None):
    """返回用于识别同一人的映射姓名；没有映射时保留打卡机姓名。"""
    punch_name = str(name or '').strip()
    if not name_mapping or punch_name not in name_mapping:
        return punch_name

    mapped = name_mapping[punch_name] or {}
    first = str(mapped.get('first_name', '') or '').strip()
    last = str(mapped.get('last_name', '') or '').strip()
    mapped_name = ' '.join(part for part in (first, last) if part)
    return mapped_name or punch_name


def record_identity_tokens(row, name_mapping=None):
    """同一日期内任一身份 token 相同即合并，兼容跨仓库员工号或姓名差异。"""
    tokens = []
    emp_id = normalize_employee_id(row.get('员工号'))
    if emp_id:
        tokens.append(('员工号', emp_id))

    mapped_name = mapped_name_for_merge(row.get('姓名', ''), name_mapping)
    name_key = normalize_text_key(mapped_name)
    if name_key:
        tokens.append(('姓名', name_key))
    return tokens


def build_identity_groups(records, name_mapping=None):
    """把记录按日期和身份 token 合成连通分组。"""
    groups = []
    token_to_group = {}

    for row in records:
        date_key = row.get('日期', '')
        tokens = [(date_key, kind, value) for kind, value in record_identity_tokens(row, name_mapping)]
        matching_groups = []
        for token in tokens:
            group_idx = token_to_group.get(token)
            if group_idx is not None and group_idx not in matching_groups:
                matching_groups.append(group_idx)

        if not matching_groups:
            group_idx = len(groups)
            groups.append([row])
        else:
            group_idx = matching_groups[0]
            groups[group_idx].append(row)
            for other_idx in matching_groups[1:]:
                if other_idx == group_idx or not groups[other_idx]:
                    continue
                groups[group_idx].extend(groups[other_idx])
                groups[other_idx] = []
                for token, existing_idx in list(token_to_group.items()):
                    if existing_idx == other_idx:
                        token_to_group[token] = group_idx

        for token in tokens:
            token_to_group[token] = group_idx

    return [group for group in groups if group]


def ordered_unique(values):
    """保持首次出现顺序去重。"""
    seen = set()
    result = []
    for value in values:
        if value is None:
            continue
        text = str(value).strip()
        if not text:
            continue
        if text not in seen:
            seen.add(text)
            result.append(text)
    return result


def group_name_aliases(rows, name_mapping=None):
    """同一合并组内可能出现的打卡名和映射名，供调整规则匹配。"""
    names = []
    for row in rows:
        raw_name = row.get('姓名', '')
        names.append(raw_name)
        mapped_name = mapped_name_for_merge(raw_name, name_mapping)
        if mapped_name != raw_name:
            names.append(mapped_name)
    return ordered_unique(names)


def choose_group_name(rows, name_mapping=None):
    """报表显示姓名优先选择有名字映射的打卡名，否则沿用第一条记录。"""
    if name_mapping:
        for row in rows:
            raw_name = str(row.get('姓名', '') or '').strip()
            if raw_name in name_mapping:
                return raw_name
    return str(rows[0].get('姓名', '') or '').strip()


# ============================================================
# 打卡记录提取
# ============================================================

def extract_from_file(filepath, source_name):
    """从单个xls文件中提取所有员工的每日打卡记录"""
    if not os.path.exists(filepath):
        print(f"警告: 文件不存在 {filepath}")
        return []

    book = xlrd.open_workbook(filepath)
    punch_sheets = [n for n in book.sheet_names() if n not in ['排班记录表', '考勤汇总表']]
    if not punch_sheets:
        print(f"警告: {filepath} 中没有打卡记录表")
        return []

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
            except (ValueError, TypeError):
                continue

            dates = ['20', '21', '22', '23', '24', '25', '26']
            weekdays = ['一', '二', '三', '四', '五', '六', '日']

            for i, (day, weekday) in enumerate(zip(dates, weekdays)):
                row = 12 + i
                if row >= sheet.nrows:
                    break

                am_start = parse_time(sheet.cell_value(row, base_col + 1))
                am_end = parse_time(sheet.cell_value(row, base_col + 3))
                pm_start = parse_time(sheet.cell_value(row, base_col + 6))
                pm_end = parse_time(sheet.cell_value(row, base_col + 8))
                ot_start = parse_time(sheet.cell_value(row, base_col + 10))
                ot_end = parse_time(sheet.cell_value(row, base_col + 12))

                records.append({
                    '员工号': emp_id,
                    '姓名': str(name).strip(),
                    '部门': str(dept).strip() if dept else '',
                    '日期': f"4月{day}日",
                    '原始日期': f"04月{day}日",
                    '星期': weekday,
                    '上午上班': format_time(am_start),
                    '上午下班': format_time(am_end),
                    '下午上班': format_time(pm_start),
                    '下午下班': format_time(pm_end),
                    '加班签到': format_time(ot_start),
                    '加班签退': format_time(ot_end),
                    '数据来源': source_name
                })

    print(f"  {filepath}: 提取了 {len(records)} 条记录")
    return records


# ============================================================
# 多来源合并与工时计算
# ============================================================

def find_adjustment(names, date_key, adj_map):
    """在同一人的多个姓名别名中查找调整表规则。"""
    for candidate in names:
        if (candidate, date_key) in adj_map:
            return adj_map[(candidate, date_key)]
    for candidate in names:
        if (candidate, 'ALL') in adj_map:
            return adj_map[(candidate, 'ALL')]
    return None


def get_adjusted_start(name, date_key, original_time, adj_map, aliases=None):
    """获取调整后的起始时间"""
    names = ordered_unique([name] + (aliases or []))
    adjustment = find_adjustment(names, date_key, adj_map)
    if adjustment is not None:
        return time_from_hour(adjustment)
    if any('mandeep' in candidate.lower() for candidate in names):
        return original_time
    if original_time is not None and original_time < time(9, 0, 0):
        return time(9, 0, 0)
    return original_time


def merge_cross_source(records_df, adj_map=None, name_mapping=None):
    """合并多来源打卡记录，应用更改说明表规则，计算最终工时（纯Python实现，避免pandas groupby兼容性问题）"""
    if records_df.empty:
        return records_df

    def to_time_obj(s):
        if not s:
            return None
        p = list(map(int, s.split(':')))
        return time(p[0], p[1], p[2])

    # 同一天内员工号相同或规范化姓名相同都合入同一组，兼容不同仓库的编码/姓名差异。
    groups = build_identity_groups([row.to_dict() for _, row in records_df.iterrows()], name_mapping)

    results = []
    for rows in sorted(groups, key=lambda group: (str(group[0].get('日期', '')), normalize_text_key(choose_group_name(group, name_mapping)))):
        first = rows[0]
        date_4m = first['日期']
        name = choose_group_name(rows, name_mapping)
        aliases = group_name_aliases(rows, name_mapping)

        # 收集所有打卡时间
        all_am_s = sorted(set(r['上午上班'] for r in rows if r['上午上班']))
        all_am_e = sorted(set(r['上午下班'] for r in rows if r['上午下班']))
        all_pm_s = sorted(set(r['下午上班'] for r in rows if r['下午上班']))
        all_pm_e = sorted(set(r['下午下班'] for r in rows if r['下午下班']))
        all_ot_s = sorted(set(r['加班签到'] for r in rows if r['加班签到']))
        all_ot_e = sorted(set(r['加班签退'] for r in rows if r['加班签退']))

        earliest_start_str = min([v for v in (all_am_s + all_pm_s + all_ot_s) if v], default='')
        latest_end_str = max([v for v in (all_am_e + all_pm_e + all_ot_e) if v], default='')

        original_start = to_time_obj(earliest_start_str)
        end_time = to_time_obj(latest_end_str)

        # 应用更改说明表规则
        if adj_map:
            date_key = date_4m.replace('4月', '04月')
            adjusted_start = get_adjusted_start(name, date_key, original_start, adj_map, aliases)
        else:
            adjusted_start = original_start

        total_hours = calc_time_diff(adjusted_start, end_time)

        # 分离正常工时和加班工时
        ot_hours = 0
        if all_ot_s and all_ot_e:
            ot_s = to_time_obj(all_ot_s[0])
            ot_e = to_time_obj(all_ot_e[-1])
            ot_hours = calc_time_diff(ot_s, ot_e)
            normal_hours = max(0, total_hours - ot_hours)
        else:
            normal_hours = total_hours

        # 调整标记
        adj_mark = ''
        if adj_map:
            date_key = date_4m.replace('4月', '04月')
            if find_adjustment(aliases, date_key, adj_map) is not None:
                adj_mark = '已调整'
            elif original_start and original_start < time(9, 0, 0) and not any('mandeep' in candidate.lower() for candidate in aliases):
                adj_mark = '9AM封顶'

        sources = '+'.join(sorted(set(r.get('数据来源', '') for r in rows if r.get('数据来源', ''))))

        results.append({
            '员工号': first.get('员工号', ''),
            '姓名': name,
            '部门': first.get('部门', ''),
            '日期': date_4m,
            '星期': first.get('星期', ''),
            '上班打卡': format_time(adjusted_start),
            '下班打卡': format_time(end_time),
            '上午上班': all_am_s[0] if all_am_s else '',
            '上午下班': all_am_e[-1] if all_am_e else '',
            '下午上班': all_pm_s[0] if all_pm_s else '',
            '下午下班': all_pm_e[-1] if all_pm_e else '',
            '加班签到': all_ot_s[0] if all_ot_s else '',
            '加班签退': all_ot_e[-1] if all_ot_e else '',
            '工作工时': round(normal_hours, 2),
            '加班工时': round(ot_hours, 2),
            '合计工时': round(total_hours, 2),
            '调整标记': adj_mark,
            '数据来源': sources
        })

    merged = pd.DataFrame(results)
    return merged.sort_values(['姓名', '日期']) if not merged.empty else merged


# ============================================================
# 生成 timesheet.csv
# ============================================================

def generate_timesheet(merged_df, name_mapping, output_path):
    """
    生成工资系统用的 timesheet.csv
    列: first_name, last_name, type, Date, end_date, description, hours, rate
    """
    valid_rows = []

    for _, row in merged_df.iterrows():
        punch_name = row['姓名']
        hours = row['合计工时']

        # 跳过无工时的记录（不需要出现在timesheet中）
        # 但示例中有hours=0的记录，所以保留

        # 名字映射
        if punch_name in name_mapping and name_mapping[punch_name].get('last_name'):
            mapped = name_mapping[punch_name]
            first = mapped.get('first_name', '')
            last = mapped.get('last_name', '')
        elif punch_name in name_mapping and name_mapping[punch_name].get('first_name'):
            # 只有first_name，作为fallback
            first = name_mapping[punch_name].get('first_name', punch_name)
            last = name_mapping[punch_name].get('last_name', '')
        else:
            # 未映射：用原始名字作为first_name
            first = punch_name
            last = ''

        # 解析日期: "4月20日" → datetime(2026, 4, 20)
        date_str = row['日期']  # 如 "4月20日"
        m = re.match(r'(\d+)月(\d+)日', date_str)
        if m:
            month, day = int(m.group(1)), int(m.group(2))
            # 推断年份：假设是当前年或次年
            year = datetime.now().year
            try:
                dt = datetime(year, month, day)
            except ValueError:
                dt = None
        else:
            dt = None

        if dt:
            date_formatted = dt.strftime('%-m/%-d/%Y')  # M/D/YYYY 如 4/20/2026
        else:
            date_formatted = date_str

        valid_rows.append({
            'first_name': first,
            'last_name': last,
            'type': 'Ordinary Time',
            'Date': date_formatted,
            'end_date': '',
            'description': '',
            'hours': round(hours, 2),
            'rate': ''
        })

    df_timesheet = pd.DataFrame(valid_rows)

    # 保存为CSV
    csv_path = output_path.replace('.xlsx', '.csv')
    if not csv_path.endswith('.csv'):
        csv_path = csv_path + '.csv'

    df_timesheet.to_csv(csv_path, index=False)
    print(f"\n✅ Timesheet CSV已生成: {csv_path}")
    print(f"   - 总行数: {len(df_timesheet)}")
    print(f"   - 有工时: {len(df_timesheet[df_timesheet['hours'] > 0])}")
    print(f"   - 零工时: {len(df_timesheet[df_timesheet['hours'] == 0])}")

    return df_timesheet


# ============================================================
# 生成详细Excel报表
# ============================================================

def generate_report(merged_df, output_path):
    """生成Excel报表"""
    detail_cols = [
        '员工号', '姓名', '部门', '日期', '星期',
        '上班打卡', '下班打卡',
        '上午上班', '上午下班', '下午上班', '下午下班',
        '加班签到', '加班签退',
        '工作工时', '加班工时', '合计工时', '调整标记', '数据来源'
    ]
    detail = merged_df[[c for c in detail_cols if c in merged_df.columns]]
    detail_valid = detail[detail['合计工时'] > 0] if '合计工时' in detail.columns else detail

    pivot = merged_df.pivot_table(
        index=['姓名', '部门'], columns='日期', values='合计工时', aggfunc='first'
    ).reset_index()

    weekly = merged_df.groupby(['姓名', '部门']).agg({
        '工作工时': 'sum', '加班工时': 'sum', '合计工时': 'sum', '日期': 'count'
    }).reset_index()
    weekly.columns = ['姓名', '部门', '本周工作工时', '本周加班工时', '本周合计工时', '打卡天数']
    pivot = pivot.merge(weekly, on=['姓名', '部门'], how='left')

    has_adj = '调整标记' in merged_df.columns
    adj_count = merged_df[merged_df['调整标记'] == '已调整'].shape[0] if has_adj else 0
    cap9_count = merged_df[merged_df['调整标记'] == '9AM封顶'].shape[0] if has_adj else 0

    readme = pd.DataFrame({
        '项目': ['数据周期', '员工总数', '有打卡记录', '数据来源', '工时规则', '输出文件', '调整统计'],
        '内容': [
            '2026-04-20 至 2026-04-26',
            f"{merged_df['姓名'].nunique()}人",
            f"{merged_df[merged_df['合计工时'] > 0]['姓名'].nunique()}人",
            ', '.join(merged_df['数据来源'].unique()),
            '1.更改表→使用表内起始时间\n2.Mandeep→按实际\n3.其他早于9AM→9AM',
            'timesheet.csv + 详细报表.xlsx',
            f"特殊调整{adj_count}条, 9AM封顶{cap9_count}条" if has_adj else "未应用"
        ]
    })

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        detail.to_excel(writer, sheet_name='每日打卡明细', index=False)
        detail_valid.to_excel(writer, sheet_name='有打卡记录的明细', index=False)
        pivot.to_excel(writer, sheet_name='每周工时汇总', index=False)
        readme.to_excel(writer, sheet_name='数据说明', index=False)

    print(f"\n✅ 详细报表已生成: {output_path}")
    print(f"   - 每日打卡明细: {len(detail)} 行")
    print(f"   - 有效打卡明细: {len(detail_valid)} 行")
    print(f"   - 每周工时汇总: {len(pivot)} 人")


# ============================================================
# 主程序
# ============================================================

def main():
    parser = argparse.ArgumentParser(description='考勤工时解析程序')
    parser.add_argument('files', nargs='+', help='输入的xls考勤文件')
    parser.add_argument('-o', '--output', default='timesheet.xlsx', help='输出Excel文件名')
    parser.add_argument('-a', '--adjustment', help='考勤更改说明表(xlsx)')
    parser.add_argument('-m', '--mapping', help='名字映射JSON文件')
    parser.add_argument('--gen-mapping', action='store_true', help='只生成名字映射模板')
    parser.add_argument('--sources', nargs='+', help='自定义数据来源名称')

    args = parser.parse_args()

    print("=" * 50)
    print("考勤工时解析程序")
    print("=" * 50)

    valid_files = [f for f in args.files if os.path.exists(f)]
    if not valid_files:
        print("错误: 没有有效的输入文件")
        sys.exit(1)

    # 提取数据
    source_names = args.sources if args.sources and len(args.sources) == len(valid_files) else \
        [os.path.splitext(os.path.basename(f))[0].split()[-1] if ' ' in os.path.splitext(os.path.basename(f))[0] 
         else os.path.splitext(os.path.basename(f))[0] for f in valid_files]

    all_records = []
    for filepath, source in zip(valid_files, source_names):
        records = extract_from_file(filepath, source)
        all_records.extend(records)

    if not all_records:
        print("错误: 没有提取到任何打卡记录")
        sys.exit(1)

    df = pd.DataFrame(all_records)
    all_punch_names = sorted(df['姓名'].dropna().unique())

    # 如果只生成映射模板
    if args.gen_mapping:
        mapping_path = args.mapping or 'name_mapping.json'
        generate_name_mapping_template(all_punch_names, mapping_path)
        return

    print(f"\n总共提取: {len(all_records)} 条记录")

    # 读取更改说明表
    adj_map = parse_adjustment_table(args.adjustment) if args.adjustment else {}

    # 读取名字映射
    name_mapping = load_name_mapping(args.mapping) if args.mapping else {}

    # 合并并计算工时
    merged = merge_cross_source(df, adj_map, name_mapping)
    print(f"合并后: {len(merged)} 条记录, {merged['姓名'].nunique()} 人")

    # 检查未映射的名字
    unmapped = [n for n in merged['姓名'].unique() if n not in name_mapping]
    if unmapped and name_mapping:
        print(f"\n⚠️  未映射的名字 ({len(unmapped)}个): {', '.join(unmapped)}")
        print("   这些名字将使用原始打卡名作为first_name")

    # 生成输出
    generate_report(merged, args.output)
    generate_timesheet(merged, name_mapping, args.output)

    print("\n完成!")


if __name__ == '__main__':
    main()
