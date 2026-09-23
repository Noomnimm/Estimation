import unittest
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook

from web_app.material_logic import (
    CODE_COL,
    DEPARTMENT_COL,
    HEAD_COL,
    MATERIAL_COL,
    QTY_COL,
    SIZE_COL,
    TOTAL_COL,
    INSULATOR_UPRIGHT_COL,
    INSULATOR_HORIZONTAL_COL,
    MaterialWorkbook,
)


class DepartmentTests(unittest.TestCase):
    def setUp(self):
        self.workbook = MaterialWorkbook()
        self.workbook.base_df = pd.DataFrame([
            {SIZE_COL: "1", HEAD_COL: "HEAD", MATERIAL_COL: "HV", CODE_COL: "100", QTY_COL: 1, DEPARTMENT_COL: "แผนกแรงสูง"},
            {SIZE_COL: "1", HEAD_COL: "HEAD", MATERIAL_COL: "TX", CODE_COL: "200", QTY_COL: 1, DEPARTMENT_COL: "แผนกหม้อแปลง", INSULATOR_UPRIGHT_COL: 7, INSULATOR_HORIZONTAL_COL: 11},
        ])

    def test_department_filters_selectors_and_calculation(self):
        self.assertEqual(self.workbook.get_sizes("แผนกแรงสูง TAC"), [])
        self.assertEqual(self.workbook.get_heads("1", "แผนกหม้อแปลง"), ["HEAD"])
        result = self.workbook.calculate([[
            {"department": "แผนกหม้อแปลง", "size": "1", "head": "HEAD", "count": "3"},
        ]])
        self.assertEqual(result["items"], [{MATERIAL_COL: "TX", CODE_COL: "200", TOTAL_COL: 3.0}])
        self.assertEqual(self.workbook.get_insulator_rate("1", "HEAD", "แผนกหม้อแปลง"), (7.0, 11.0))

    def test_combined_hardware_export_contains_insulator_sheets(self):
        data = self.workbook.export_page_hardware_combined([[
            {"department": "แผนกหม้อแปลง", "size": "1", "head": "HEAD", "count": "2"},
        ]])
        workbook = load_workbook(BytesIO(data), read_only=True)
        self.assertEqual(workbook.sheetnames, [
            "ลูกถ้วยและอุปกรณ์", "ที่มารายการ", "สรุปลูกถ้วยแยกหน้า", "ที่มาลูกถ้วย",
        ])


if __name__ == "__main__":
    unittest.main()
