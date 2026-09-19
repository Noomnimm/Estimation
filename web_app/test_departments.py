import unittest

import pandas as pd

from web_app.material_logic import (
    CODE_COL,
    DEPARTMENT_COL,
    HEAD_COL,
    MATERIAL_COL,
    QTY_COL,
    SIZE_COL,
    TOTAL_COL,
    MaterialWorkbook,
)


class DepartmentTests(unittest.TestCase):
    def setUp(self):
        self.workbook = MaterialWorkbook()
        self.workbook.base_df = pd.DataFrame([
            {SIZE_COL: "1", HEAD_COL: "HEAD", MATERIAL_COL: "HV", CODE_COL: "100", QTY_COL: 1, DEPARTMENT_COL: "แผนกแรงสูง"},
            {SIZE_COL: "1", HEAD_COL: "HEAD", MATERIAL_COL: "TX", CODE_COL: "200", QTY_COL: 1, DEPARTMENT_COL: "แผนกหม้อแปลง"},
        ])

    def test_department_filters_selectors_and_calculation(self):
        self.assertEqual(self.workbook.get_sizes("แผนกแรงสูง TAC"), [])
        self.assertEqual(self.workbook.get_heads("1", "แผนกหม้อแปลง"), ["HEAD"])
        result = self.workbook.calculate([[
            {"department": "แผนกหม้อแปลง", "size": "1", "head": "HEAD", "count": "3"},
        ]])
        self.assertEqual(result["items"], [{MATERIAL_COL: "TX", CODE_COL: "200", TOTAL_COL: 3.0}])


if __name__ == "__main__":
    unittest.main()
