import unittest

import pandas as pd

from web_app.material_logic import (
    MaterialWorkbook, ohgw_material, SIZE_COL, HEAD_COL, MATERIAL_COL,
    CODE_COL, QTY_COL, TOTAL_COL, SET_COL, SET_DESC_COL, SET_INSTALL_COL,
)


class OHGWTests(unittest.TestCase):
    def test_mapping(self):
        cases = [
            ('DDE บน', '14', 'Set25258'), ('DDE.BL', '16', 'Set25258'),
            ('DE', '12', 'Set25256'), ('BA', '14', 'Set25261'),
            ('SP', '12.20', 'Set25251'), ('SP', '14.00', 'Set25251'),
            ('SP', '14.30', 'Set25264'), ('DP', '12.2', 'Set25262'),
            ('CCB ประกบบน', '14.0', 'Set25262'), ('CCB', '14.3', 'Set25266'),
            ('DP', '14.30', 'Set25266'), ('CCB,CCB', '14.3', 'Set25266'),
        ]
        for head, size, expected in cases:
            with self.subTest(head=head, size=size):
                self.assertEqual(ohgw_material(head, size)[1], expected)

    def test_unknown(self):
        for head, size in [('SP', '16'), ('BA', '12.2'), ('DDE', '14.3'), ('2SP', '14'), ('CSC', '14')]:
            with self.assertRaises(ValueError):
                ohgw_material(head, size)

    def test_calculate_expand_and_legacy(self):
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame([{
            SIZE_COL: '14.3', HEAD_COL: 'SP', MATERIAL_COL: 'Base', CODE_COL: '100', QTY_COL: 1,
        }])
        row = {'size': '14.3', 'head': 'SP', 'count': '4+4-2'}
        self.assertEqual(len(workbook.calculate([[row]])['items']), 1)
        result = workbook.calculate([[dict(row, ohgw=True)], [dict(row, ohgw=True)]])
        added = next(item for item in result['items'] if item[CODE_COL] == 'Set25264')
        self.assertEqual(added[TOTAL_COL], 12)
        workbook.set_df = pd.DataFrame([{
            SET_COL: 'Set25264', CODE_COL: '200', SET_DESC_COL: 'OHGW part', SET_INSTALL_COL: 2,
        }])
        expanded = workbook.expand_set()
        self.assertEqual(next(item[TOTAL_COL] for item in expanded['items'] if item[CODE_COL] == '200'), 24)
        self.assertTrue(workbook.export_summary())


if __name__ == '__main__':
    unittest.main()
