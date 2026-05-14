import unittest

import pandas as pd

from attendance_parser import merge_cross_source


def record(emp_id, name, source, am_start='', am_end='', pm_start='', pm_end=''):
    return {
        '员工号': emp_id,
        '姓名': name,
        '部门': '仓库',
        '日期': '4月20日',
        '星期': '一',
        '上午上班': am_start,
        '上午下班': am_end,
        '下午上班': pm_start,
        '下午下班': pm_end,
        '加班签到': '',
        '加班签退': '',
        '数据来源': source,
    }


class AttendanceMergeTest(unittest.TestCase):
    def test_merges_same_employee_id_with_different_names(self):
        df = pd.DataFrame([
            record(1001, 'YING YU', 'warehouse-a', am_start='08:00:00', am_end='12:00:00'),
            record(1001, 'YINGYU', 'warehouse-b', pm_start='13:00:00', pm_end='18:00:00'),
        ])

        merged = merge_cross_source(df)

        self.assertEqual(len(merged), 1)
        row = merged.iloc[0]
        self.assertEqual(row['上班打卡'], '08:00:00')
        self.assertEqual(row['下班打卡'], '18:00:00')
        self.assertEqual(row['合计工时'], 10)
        self.assertEqual(row['数据来源'], 'warehouse-a+warehouse-b')

    def test_keeps_existing_same_name_merge_when_employee_ids_differ(self):
        df = pd.DataFrame([
            record(2001, 'ISABEL', 'warehouse-a', am_start='09:00:00', am_end='12:00:00'),
            record(3001, 'ISABEL', 'warehouse-b', pm_start='13:00:00', pm_end='17:00:00'),
        ])

        merged = merge_cross_source(df)

        self.assertEqual(len(merged), 1)
        self.assertEqual(merged.iloc[0]['合计工时'], 8)

    def test_merges_different_names_with_same_mapping(self):
        df = pd.DataFrame([
            record(2001, 'YING YU', 'warehouse-a', am_start='09:00:00', am_end='12:00:00'),
            record(3001, 'YINGYU', 'warehouse-b', pm_start='13:00:00', pm_end='17:00:00'),
        ])
        mapping = {
            'YING YU': {'first_name': 'YING', 'last_name': 'YU'},
            'YINGYU': {'first_name': 'YING', 'last_name': 'YU'},
        }

        merged = merge_cross_source(df, name_mapping=mapping)

        self.assertEqual(len(merged), 1)
        self.assertEqual(merged.iloc[0]['姓名'], 'YING YU')
        self.assertEqual(merged.iloc[0]['合计工时'], 8)

    def test_adjustment_table_matches_any_merged_alias(self):
        df = pd.DataFrame([
            record(1001, 'YING YU', 'warehouse-a', am_start='08:00:00', am_end='12:00:00'),
            record(1001, 'YINGYU', 'warehouse-b', pm_start='13:00:00', pm_end='18:00:00'),
        ])

        merged = merge_cross_source(df, adj_map={('YINGYU', '04月20日'): 7})

        self.assertEqual(len(merged), 1)
        self.assertEqual(merged.iloc[0]['上班打卡'], '07:00:00')
        self.assertEqual(merged.iloc[0]['合计工时'], 11)
        self.assertEqual(merged.iloc[0]['调整标记'], '已调整')


if __name__ == '__main__':
    unittest.main()
