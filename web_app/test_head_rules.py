import unittest
import pandas as pd
from web_app.material_logic import (
    MaterialWorkbook, classify_wire_head, wire_head_multiplier, add_wire_materials,
    SIZE_COL, HEAD_COL, MATERIAL_COL, CODE_COL, QTY_COL, TOTAL_COL,
)


class HeadRuleTests(unittest.TestCase):
    def equipment(self, head, wire1='185 SAC', wire2='185 SAC'):
        totals = {}
        add_wire_materials(totals, classify_wire_head(head), wire1, wire2, wire_head_multiplier(head))
        return {r[CODE_COL]: r[TOTAL_COL] for r in totals.values()}

    def test_double_heads(self):
        for head, preform in [('2BA.st 4.5 m', 6), ('2DE.st4.5', 6), ('2DDE.st 4.5m', 12)]:
            with self.subTest(head=head):
                values = self.equipment(head)
                self.assertEqual(values['1020260205'], preform)
                self.assertEqual(values['1030140011'], preform)
        values = self.equipment('2DDE.st 4.5m')
        for code in ['1020410027', '1020180001', '1020180008']:
            self.assertEqual(values[code], 6)
        self.assertEqual(self.equipment('2BA')['1020300103'], 12)
        small = self.equipment('2BA', wire2='50 SAC')
        self.assertEqual(small['1020260202'], 6)
        self.assertEqual(small['1020330104'], 6)
        self.assertEqual(small['1020330006'], 6)

    def test_single_phase(self):
        for head, expected in [('BA 1-P', 2), ('BA.AL 1P', 2), ('DE.CON 1-P', 2), ('DDE 1-P', 4), ('DDE.BL 1-P', 4)]:
            self.assertEqual(self.equipment(head)['1020260205'], expected)
        self.assertEqual(self.equipment('DDE 1-P')['1020410027'], 2)
        self.assertNotIn('1020410027', self.equipment('DDE.BL 1-P'))
        self.assertEqual(self.equipment('BA 1-P')['1020300103'], 4)

    def test_strain_and_pending(self):
        values = self.equipment('2DE', wire1='185 A')
        self.assertEqual(values['1030110004'], 6)
        self.assertNotIn('1030140011', values)
        for head in ['LAT.SLK บน', '2BA.st 4.5m+DE.CON', '2DE.st4.5 + DE.CON']:
            self.assertIsNone(classify_wire_head(head))
        self.assertEqual(classify_wire_head('BA.SLK บน'), 'ba')  # Preserve existing wire rule.
        self.assertEqual(self.equipment('DDE.st 3m, LAT.SLK')['1020260205'], 6)

    def test_base_not_doubled(self):
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame([{SIZE_COL: '99', HEAD_COL: '2DE.st4.5', MATERIAL_COL: 'Base', CODE_COL: 'SetExample', QTY_COL: 1}])
        row = {'size': '99', 'head': '2DE.st4.5', 'count': '1+1', 'wire1': '185 SAC'}
        result = workbook.calculate([[row]])
        values = {r[CODE_COL]: r[TOTAL_COL] for r in result['items']}
        self.assertEqual(values['SetExample'], 2)
        self.assertEqual(values['1020260205'], 12)


if __name__ == '__main__':
    unittest.main()
