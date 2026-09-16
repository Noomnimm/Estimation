import unittest
from pathlib import Path

from web_app.material_logic import MaterialWorkbook, natural_key


class HeadSortingTests(unittest.TestCase):
    def test_punctuation_and_numeric_order(self):
        self.assertEqual(
            sorted(['SP.st.10m', 'SP.st.4.5m', 'SP.st.3m'], key=natural_key),
            ['SP.st.3m', 'SP.st.4.5m', 'SP.st.10m'],
        )
        for name in ['2BA st.4.5m', 'SP,DDE.BL st.4.5m', 'A..4.5m', '.', '']:
            natural_key(name)

    def test_all_workbook_head_menus(self):
        workbook = MaterialWorkbook()
        workbook.load_base(Path(__file__).resolve().parents[1] / 'Newdata.xlsx')
        for size in workbook.get_sizes():
            with self.subTest(size=size):
                self.assertTrue(workbook.get_heads(size))
        for size in ['12.2', '14.3', '22']:
            heads = workbook.get_heads(size)
            for name in ['BA.st4.5m', '2BA st.4.5m', 'SP,DDE.BL st.4.5m', 'DP,DDE.BL st.4.5m', 'DP,DDE st.4.5m', 'DP,DE st.4.5m']:
                self.assertIn(name, heads)


if __name__ == '__main__':
    unittest.main()
