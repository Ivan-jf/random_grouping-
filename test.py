import random
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from app import app, add_volume_column, calculate_volume, detect_dimension_columns, random_grouping, select_volume_window


class VolumeWorkflowTests(unittest.TestCase):
    def test_detects_unique_dimension_columns_and_calculates_volume(self):
        columns = ['耳标号', '体重', '长(mm)', '宽(mm)']
        length_col, width_col = detect_dimension_columns(columns)
        self.assertEqual((length_col, width_col), ('长(mm)', '宽(mm)'))
        self.assertEqual(calculate_volume(10, 4), 80)

    def test_rejects_multiple_dimension_candidates(self):
        with self.assertRaisesRegex(ValueError, '多个长'):
            detect_dimension_columns(['长1', '长2', '宽'])

    def test_selects_exact_contiguous_window(self):
        df = pd.DataFrame({'耳标号': ['a', 'b', 'c', 'd'], '体积': [40, 30, 20, 10]})
        selected = select_volume_window(df, start=1, count=2)
        self.assertEqual(selected['耳标号'].tolist(), ['b', 'c'])

    def test_volume_is_rounded_to_two_decimal_places(self):
        df = pd.DataFrame({'长': [3], '宽': [1.234]})
        result = add_volume_column(df, '长', '宽')
        self.assertEqual(result['体积'].iloc[0], 2.28)

    def test_rejects_window_out_of_bounds(self):
        df = pd.DataFrame({'体积': [3, 2, 1]})
        with self.assertRaisesRegex(ValueError, '范围'):
            select_volume_window(df, start=2, count=2)

    def test_unequal_groups_match_requested_counts_and_spread_across_volume(self):
        random.seed(4)
        df = pd.DataFrame({'耳标号': range(12), '体积': range(120, 0, -10)})
        result, summary, _ = random_grouping(df, 2, ['A', 'B', 'C'], '耳标号', '体积',
                                        group_sizes={'A': 6, 'B': 4, 'C': 2})
        self.assertEqual(result['group'].value_counts().to_dict(), {'A': 6, 'B': 4, 'C': 2})
        self.assertEqual(result.iloc[:6]['group'].value_counts().to_dict(), {'A': 3, 'B': 2, 'C': 1})
        self.assertEqual(dict(zip(summary['分组'], summary['只数'])), {'A': 6, 'B': 4, 'C': 2})
        self.assertLessEqual(max(abs(value - df['体积'].mean()) for value in summary['均值']), 2.0)

    def test_default_equal_groups_still_fill_each_block(self):
        df = pd.DataFrame({'耳标号': range(6), '体积': range(6, 0, -1)})
        result, _, _ = random_grouping(df, 2, ['A', 'B', 'C'], '耳标号', '体积')
        for _, block in result.groupby('block'):
            self.assertEqual(set(block['group']), {'A', 'B', 'C'})

    def test_unequal_balancing_keeps_random_assignments(self):
        volumes = [150, 143, 141, 126, 120, 110, 97, 95, 86, 82, 79, 77, 64, 50, 47]
        df = pd.DataFrame({'耳标号': range(15), '体积': volumes})
        outcomes = set()
        for seed in (7, 8, 9):
            random.seed(seed)
            result, _, _ = random_grouping(df, 4, ['A', 'B', 'C'], '耳标号', '体积',
                                           group_sizes={'A': 6, 'B': 5, 'C': 4})
            outcomes.add(tuple(result['group']))
        self.assertGreater(len(outcomes), 1)

    def test_run_preview_and_result_use_unequal_total(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'mice.xlsx'
            pd.DataFrame({'耳标号': range(1, 13), '长': range(12, 0, -1),
                          '宽': [2] * 12}).to_excel(source, index=False)
            payload = {'filepath': str(source), 'group_num': 2,
                       'group_names': 'A,B,C', 'group_sizes': {'A': 4, 'C': 3},
                       'id_col': '耳标号', 'length_col': '长', 'width_col': '宽'}
            client = app.test_client()
            preview = client.post('/run', json={**payload, 'preview_only': True})
            self.assertEqual(preview.status_code, 200)
            self.assertEqual(preview.json['required_count'], 9)
            self.assertEqual(preview.json['group_sizes'], {'A': 4, 'B': 2, 'C': 3})
            result = client.post('/run', json={**payload, 'start_index': 1})
            self.assertEqual(result.status_code, 200)
            self.assertEqual(result.json['total'], 9)
            counts = pd.Series([row['group'] for row in result.json['preview']]).value_counts().to_dict()
            self.assertEqual(counts, {'A': 4, 'B': 2, 'C': 3})


if __name__ == '__main__':
    unittest.main()
