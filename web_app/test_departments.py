import unittest
from io import BytesIO
from unittest.mock import patch

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

    def test_catalog_code_override_is_stored_as_ten_digit_text(self):
        raw = pd.DataFrame([[None] * 9 for _ in range(8)])
        raw.iloc[7, 1] = "15"
        raw.iloc[7, 2] = "หม้อแปลงขนาด 30 เควีเอ"
        raw.iloc[7, 8] = "1050010004"
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame(columns=[SIZE_COL, HEAD_COL, MATERIAL_COL, CODE_COL, QTY_COL, DEPARTMENT_COL])
        with patch("web_app.material_logic.pd.read_excel", return_value=raw):
            workbook.load_keycode_catalog("unused.xlsx", "แผนกหม้อแปลง", {"15": "1050000011"})
        self.assertEqual(workbook.base_df.iloc[0][CODE_COL], "1050000011")

    def test_department_base_preserves_negative_quantities_and_normalizes_codes(self):
        source = pd.DataFrame([{
            "รายการ": "30kVA 1P", MATERIAL_COL: "อุปกรณ์ทดสอบ",
            CODE_COL: "1-05-000-0011", QTY_COL: -2,
        }])
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame(columns=[SIZE_COL, HEAD_COL, MATERIAL_COL, CODE_COL, QTY_COL, DEPARTMENT_COL])
        with patch("web_app.material_logic.pd.read_excel", return_value=source):
            workbook.load_department_base("unused.xlsx", "แผนกหม้อแปลง", "หม้อแปลง")
        row = workbook.base_df.iloc[0]
        self.assertEqual(row[CODE_COL], "1050000011")
        self.assertEqual(row[QTY_COL], -2)

    def test_transformer_surge_arrester_follows_ngr_and_distance(self):
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame([
            {SIZE_COL: "หม้อแปลง", HEAD_COL: "30kVA 1P", MATERIAL_COL: "OLD SA", CODE_COL: "1040000007", QTY_COL: -2, DEPARTMENT_COL: "แผนกหม้อแปลง"},
            {SIZE_COL: "หม้อแปลง", HEAD_COL: "30kVA 1P", MATERIAL_COL: "NEW LA", CODE_COL: "1040000000", QTY_COL: 2, DEPARTMENT_COL: "แผนกหม้อแปลง"},
        ])
        cases = [
            (False, False, "1040000007", "1040000000"),
            (False, True, "1040000007", "1040000001"),
            (True, False, "1040000008", "1040000002"),
            (True, True, "1040000008", "1040000003"),
        ]
        for ngr, within, removed_code, installed_code in cases:
            result = workbook.calculate([[
                {"department": "แผนกหม้อแปลง", "size": "หม้อแปลง", "head": "30kVA 1P", "count": 1,
                 "surgeNgr": ngr, "surgeWithin3km": within},
            ]])
            values = {row[CODE_COL]: row[TOTAL_COL] for row in result["items"]}
            self.assertEqual(values, {removed_code: -2.0, installed_code: 2.0})
        for ngr, tank_code in [(False, "1040000007"), (True, "1040000008")]:
            result = workbook.calculate([[
                {"department": "แผนกหม้อแปลง", "size": "หม้อแปลง", "head": "30kVA 1P", "count": 1,
                 "surgeNgr": ngr, "surgeWithin3km": False, "surgeMounting": "tank"},
            ]])
            self.assertEqual({row[CODE_COL]: row[TOTAL_COL] for row in result["items"]}, {tank_code: 2.0})


if __name__ == "__main__":
    unittest.main()
