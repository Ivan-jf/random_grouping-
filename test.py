import unittest

import pandas as pd

from app import add_volume_column, calculate_volume, detect_dimension_columns, select_volume_window


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


if __name__ == '__main__':
    unittest.main()
