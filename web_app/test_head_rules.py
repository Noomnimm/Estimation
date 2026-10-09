import unittest
from io import BytesIO
import pandas as pd
from openpyxl import load_workbook
from web_app.material_logic import (
    MaterialWorkbook, classify_wire_head, wire_head_multiplier, add_wire_materials,
    insulator_rate, page_labels, parse_number, SIZE_COL, HEAD_COL, MATERIAL_COL, CODE_COL, QTY_COL, TOTAL_COL,
)


class HeadRuleTests(unittest.TestCase):
    def test_count_expressions_support_basic_arithmetic(self):
        self.assertEqual(parse_number("(4+4)*2"), 16)
        self.assertEqual(parse_number("1200/3+100"), 500)
        self.assertEqual(parse_number("10-2*3"), 4)
        self.assertEqual(parse_number("1/0"), 0)
        self.assertEqual(parse_number("2**8"), 0)

    def test_manual_poles_and_stubs_are_added_to_selected_work_type(self):
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame(columns=[SIZE_COL, HEAD_COL, MATERIAL_COL, CODE_COL, QTY_COL])
        pages = [[{"workType": "install"}], [{"workType": "demolition"}]]
        structures = [
            [{"workType": "install", "code": "1000010012", "material": "POLE 12.20", "count": "2+3"}],
            [{"workType": "demolition", "code": "Set14316", "material": "CONCRETE STUB", "count": "4"}],
        ]
        result = workbook.calculate(pages, structures)
        by_code = {item[CODE_COL]: item for item in result["items"]}
        self.assertEqual(by_code["1000010012"]["จำนวนติดตั้ง"], 5)
        self.assertEqual(by_code["1000010012"]["จำนวนรื้อถอน"], 0)
        self.assertEqual(by_code["Set14316"]["จำนวนติดตั้ง"], 0)
        self.assertEqual(by_code["Set14316"]["จำนวนรื้อถอน"], 4)

    def test_inline_set_components_only_return_ten_digit_codes(self):
        workbook = MaterialWorkbook()
        workbook.set_df = pd.DataFrame([
            {"Set": "Set20202", "รหัสพัสดุ": "1010110202", "คำอธิบาย": "BOLT", "ติดตั้ง": 2},
            {"Set": "Set20202", "รหัสพัสดุ": "Set99999", "คำอธิบาย": "NESTED SET", "ติดตั้ง": 1},
        ])
        result = workbook.get_set_components("set20202", 3)
        self.assertEqual(result["items"], [{"material": "BOLT", "code": "1010110202", "quantity": 6.0}])

    def test_demolition_set_components_use_reuse_quantity(self):
        workbook = MaterialWorkbook()
        workbook.set_df = pd.DataFrame([
            {"Set": "Set20202", "รหัสพัสดุ": "1010110202", "คำอธิบาย": "BOLT", "ติดตั้ง": 5, "จำนวนนำกลับมาใช้ใหม่": 2},
            {"Set": "Set20202", "รหัสพัสดุ": "1010110203", "คำอธิบาย": "OTHER BOLT", "ติดตั้ง": 3, "จำนวนนำกลับมาใช้ใหม่": 0},
        ])

        result = workbook.get_set_components("set20202", 3, "demolition")
        self.assertEqual(result["items"], [{"material": "BOLT", "code": "1010110202", "quantity": 6.0}])

    def test_demolition_set_components_do_not_assume_install_quantity_is_reusable(self):
        workbook = MaterialWorkbook()
        workbook.set_df = pd.DataFrame([
            {"Set": "Set21326", "รหัสพัสดุ": "1010180100", "คำอธิบาย": "SQUARE WASHER", "ติดตั้ง": 18},
        ])

        result = workbook.get_set_components("set21326", 1, "demolition")
        self.assertEqual(result["items"], [])

    def test_combined_set_components_separate_install_and_reuse_quantities(self):
        workbook = MaterialWorkbook()
        workbook.set_df = pd.DataFrame([
            {"Set": "Set20202", "รหัสพัสดุ": "1010110202", "คำอธิบาย": "BOLT", "ติดตั้ง": 5, "จำนวนนำกลับมาใช้ใหม่": 2},
        ])

        result = workbook.get_set_components("set20202", 2, "combined")
        self.assertEqual(result["items"], [{
            "material": "BOLT", "code": "1010110202", "installQuantity": 10.0,
            "demolitionQuantity": 4.0, "quantity": 10.0,
        }])

    def test_expand_set_uses_reuse_quantity_for_demolition(self):
        workbook = MaterialWorkbook()
        workbook.set_df = pd.DataFrame([
            {"Set": "Set20202", "รหัสพัสดุ": "1010110202", "คำอธิบาย": "BOLT", "ติดตั้ง": 5, "จำนวนนำกลับมาใช้ใหม่": 2},
            {"Set": "Set20202", "รหัสพัสดุ": "1010110203", "คำอธิบาย": "OTHER BOLT", "ติดตั้ง": 3, "จำนวนนำกลับมาใช้ใหม่": 0},
        ])
        workbook.summary = [{"ก่อนแตก Set": True}]
        workbook.separate_work_types = True
        workbook.summary_by_work_type = {
            "install": [],
            "demolition": [{MATERIAL_COL: "BOLT SET", CODE_COL: "Set20202", TOTAL_COL: 3}],
        }

        result = workbook.expand_set()
        by_code = {item[CODE_COL]: item for item in result["items"]}
        self.assertEqual(by_code["1010110202"]["จำนวนรื้อถอน"], 6.0)
        self.assertNotIn("1010110203", by_code)

    def test_page_labels_separate_installation_and_demolition(self):
        pages = [
            [{"workType": "install"}],
            [{"workType": "install"}],
            [{"workType": "demolition"}],
            [{"workType": "demolition"}],
        ]
        self.assertEqual(page_labels(pages), ["ติดตั้ง 1", "ติดตั้ง 2", "รื้อถอน 1", "รื้อถอน 2"])
        self.assertEqual(page_labels([[{}], [{}]]), ["หน้า 1", "หน้า 2"])

    def test_page_export_groups_work_types_in_one_sheet(self):
        workbook = MaterialWorkbook()
        pages = [
            [{"workType": "install", "size": "12.2", "head": "SP", "count": 1}],
            [{"workType": "install", "size": "12.2", "head": "SP", "count": 2}],
            [{"workType": "demolition", "size": "12.2", "head": "SP", "count": 3}],
        ]
        exported = load_workbook(BytesIO(workbook.export_page_summary(pages)), data_only=True)
        sheet = exported["สรุปแต่ละหน้า"]
        self.assertEqual(sheet["C2"].value, "งานติดตั้ง")
        self.assertEqual(sheet["F2"].value, "งานรื้อถอน")
        self.assertEqual([sheet.cell(3, column).value for column in range(3, 8)], ["1", "2", "รวม", "1", "รวม"])
        self.assertIn("C2:E2", {str(cell_range) for cell_range in sheet.merged_cells.ranges})
        self.assertIn("F2:G2", {str(cell_range) for cell_range in sheet.merged_cells.ranges})
        self.assertEqual([sheet.cell(4, column).value for column in range(3, 8)], [1, 2, 3, 3, 3])

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

    def test_dde_185_sac_to_acsr_uses_pg3_instead_of_tensionless(self):
        for left, right in (("185 SAC", "185 ACSR"), ("185 ACSR", "185 SAC")):
            with self.subTest(left=left, right=right):
                values = self.equipment("DDE", wire1=left, wire2=right)
                self.assertEqual(values["1020300103"], 6)
                self.assertNotIn("1020410027", values)
                self.assertNotIn("1020180001", values)
                self.assertNotIn("1020180008", values)
                self.assertEqual(values["1020260205"], 3)
                self.assertEqual(values["1030140011"], 3)
                self.assertEqual(values["1030110007"], 3)

    def test_dde_tensionless_sleeves_follow_new_construction_codes(self):
        cases = [
            ("50 SAC", "1020410022"),
            ("185 SAC", "1020410027"),
            ("50 ACSR", "1020410014"),
            ("185 ACSR", "1020410017"),
        ]
        for wire, expected_code in cases:
            with self.subTest(wire=wire):
                values = self.equipment("DDE", wire1=wire, wire2=wire)
                self.assertEqual(values[expected_code], 3)
                self.assertNotIn("1020410002", values)
                self.assertNotIn("1020410007", values)

    def test_strain_and_pending(self):
        values = self.equipment('2DE', wire1='185 A')
        self.assertEqual(values['1030110004'], 6)
        self.assertNotIn('1030140011', values)
        for head in ['2BA.st 4.5m+DE.CON', '2DE.st4.5 + DE.CON']:
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

    def test_calculation_keeps_installation_and_demolition_totals_separate(self):
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame([
            {SIZE_COL: '12.2', HEAD_COL: 'SP', MATERIAL_COL: 'TEST MATERIAL', CODE_COL: '1000000001', QTY_COL: 2},
        ])
        result = workbook.calculate([
            [{'workType': 'install', 'size': '12.2', 'head': 'SP', 'count': 3}],
            [{'workType': 'demolition', 'size': '12.2', 'head': 'SP', 'count': 4}],
        ])
        item = next(row for row in result['items'] if row[CODE_COL] == '1000000001')
        self.assertEqual(item['จำนวนติดตั้ง'], 6)
        self.assertEqual(item['จำนวนรื้อถอน'], 8)

    def test_wire_accessory_addons_are_not_added_to_demolition(self):
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame([
            {SIZE_COL: '99', HEAD_COL: 'DDE', MATERIAL_COL: 'BASE HEAD MATERIAL', CODE_COL: '1000000001', QTY_COL: 1},
        ])
        result = workbook.calculate([[
            {'workType': 'demolition', 'size': '99', 'head': 'DDE', 'count': 1},
        ]])
        by_code = {row[CODE_COL]: row for row in result['items']}
        self.assertEqual(by_code['1000000001']['จำนวนรื้อถอน'], 1)
        for code in ('1020260205', '1030140011', '1020410027', '1020180001', '1020180008'):
            self.assertNotIn(code, by_code)

    def test_demolition_filters_non_reusable_base_rows_using_allset(self):
        workbook = MaterialWorkbook()
        head = 'X-ARM-C, 3-P, DEAD END CONNECTION'
        codes = [
            ('SET20211', '1010180001', 'EYE NUT', 6),
            ('SET21326', '1010180100', 'SQUARE WASHER', 12),
            ('SET99901', '1010110202', 'MACHINE BOLT', 4),
            ('SET99902', '1010130001', 'DOUBLE ARMING', -2),
            ('SET99903', '1010180301', 'LOCK WASHER', 1),
            ('SET99904', '1000000001', 'REUSABLE MATERIAL', 3),
        ]
        workbook.base_df = pd.DataFrame([
            {SIZE_COL: '99', HEAD_COL: head, MATERIAL_COL: 'SET 20211', CODE_COL: 'Set20211', QTY_COL: 2},
            {SIZE_COL: '99', HEAD_COL: head, MATERIAL_COL: 'SET 21326', CODE_COL: 'Set21326', QTY_COL: 1},
            *[
                {SIZE_COL: '99', HEAD_COL: head, MATERIAL_COL: material, CODE_COL: code, QTY_COL: quantity}
                for _, code, material, quantity in codes
            ],
        ])
        workbook.set_df = pd.DataFrame([
            {'Set': set_code, CODE_COL: code, 'คำอธิบาย': material, 'ติดตั้ง': abs(quantity), 'จำนวนนำกลับมาใช้ใหม่': 0}
            for set_code, code, material, quantity in codes
        ] + [
            {'Set': 'Set12000', CODE_COL: '1010180100', 'คำอธิบาย': 'SQUARE WASHER', 'ติดตั้ง': 4, 'จำนวนนำกลับมาใช้ใหม่': 4},
            {'Set': 'Set99905', CODE_COL: '1000000001', 'คำอธิบาย': 'REUSABLE MATERIAL', 'ติดตั้ง': 3, 'จำนวนนำกลับมาใช้ใหม่': 3},
        ])

        result = workbook.calculate([
            [{'workType': 'install', 'size': '99', 'head': head, 'count': 1}],
            [{'workType': 'demolition', 'size': '99', 'head': head, 'count': 1}],
        ])
        by_code = {row[CODE_COL]: row for row in result['items']}
        for _, code, _, _ in codes[:5]:
            self.assertNotEqual(by_code[code]['จำนวนติดตั้ง'], 0)
            self.assertEqual(by_code[code]['จำนวนรื้อถอน'], 0)
        self.assertEqual(by_code['1000000001']['จำนวนรื้อถอน'], 3)

    def test_confirmed_unmatched_codes_remain_visible_in_demolition(self):
        workbook = MaterialWorkbook()
        confirmed_codes = {
            '1020440000', '1020440008', '1020440119', '1030010200', '1040000002',
            '1040010015', '1040010016', '1040030002', '1060020050',
        }
        workbook.base_df = pd.DataFrame([
            {SIZE_COL: '99', HEAD_COL: 'SP', MATERIAL_COL: f'MATERIAL {code}', CODE_COL: code, QTY_COL: 1}
            for code in confirmed_codes
        ])
        workbook.set_df = pd.DataFrame([
            {'Set': 'Set99999', CODE_COL: code, 'คำอธิบาย': f'MATERIAL {code}', 'ติดตั้ง': 1, 'จำนวนนำกลับมาใช้ใหม่': 0}
            for code in confirmed_codes
        ])

        result = workbook.calculate([[
            {'workType': 'demolition', 'size': '99', 'head': 'SP', 'count': 1},
        ]])
        by_code = {row[CODE_COL]: row for row in result['items']}
        for code in confirmed_codes:
            self.assertEqual(by_code[code]['จำนวนรื้อถอน'], 1)

    def test_page_hardware_export_omits_wire_accessory_addons_for_demolition(self):
        workbook = MaterialWorkbook()
        pages = [[{
            'workType': 'demolition', 'size': '99', 'head': 'DDE', 'count': 1,
            # Removal should not need wire-selection data just to calculate accessories.
        }]]
        exported = load_workbook(BytesIO(workbook.export_page_hardware(pages)), data_only=True)
        sheet = exported['ลูกถ้วยและอุปกรณ์']
        accessory_codes = {
            str(row[2]) for row in sheet.iter_rows(min_row=4, values_only=True)
            if row[0] == 'อุปกรณ์ยึดสาย' and row[2]
        }
        self.assertEqual(accessory_codes, set())

    def test_lat_as_de(self):
        for head in ['LAT.SLK บน', 'LAT.SLK ล่าง']:
            values = self.equipment(head, wire1='50 SAC')
            self.assertEqual(values['1020260202'], 3)
            self.assertEqual(values['1030140011'], 3)
            self.assertNotIn('1020410027', values)
            self.assertNotIn('1020180001', values)
            strain = self.equipment(head, wire1='185 A')
            self.assertEqual(strain['1030110004'], 3)
            self.assertNotIn('1030140011', strain)

    def test_mixed_lat(self):
        workbook = MaterialWorkbook()
        head = 'DDE.st 3m, LAT.SLK'
        workbook.base_df = pd.DataFrame([{SIZE_COL: '99', HEAD_COL: head, MATERIAL_COL: 'Base', CODE_COL: 'SetExample', QTY_COL: 1}])
        row = {'size': '99', 'head': head, 'count': '1+1', 'wire1': '185 SAC', 'wire2': '185 SAC'}
        with self.assertRaisesRegex(ValueError, 'LAT.SLK'):
            workbook.calculate([[row]])
        row['latWire'] = '50 SAC'
        result = workbook.calculate([[row]])
        values = {r[CODE_COL]: r[TOTAL_COL] for r in result['items']}
        self.assertEqual(values['SetExample'], 2)
        self.assertEqual(values['1020260205'], 12)
        self.assertEqual(values['1020260202'], 6)
        self.assertEqual(values['1030140011'], 18)
        self.assertEqual(values['1020410027'], 6)
        self.assertEqual(values['1020180001'], 6)

    def test_new_steel_combined_heads(self):
        cases = {
            'SP,DDE.BL st.4.5m': 'dde_bl',
            'DP,DDE.BL st.4.5m': 'dde_bl',
            'DP,DDE st.4.5m': 'dde',
            'DP,DE st.4.5m': 'de',
        }
        for head, expected_kind in cases.items():
            with self.subTest(head=head):
                self.assertEqual(classify_wire_head(head), expected_kind)
                values = self.equipment(head)
                expected_preform = 3 if expected_kind == 'de' else 6
                self.assertEqual(values['1020260205'], expected_preform)
                self.assertEqual(values['1030140011'], expected_preform)
                if expected_kind == 'dde':
                    self.assertEqual(values['1020410027'], 3)
                    self.assertEqual(values['1020180001'], 3)
                    self.assertEqual(values['1020180008'], 3)
                else:
                    self.assertNotIn('1020410027', values)

    def test_dde_de_combined_head(self):
        head = 'DDE,DE (St.4.5m)'
        self.assertEqual(classify_wire_head(head), 'dde_de')
        self.assertEqual(insulator_rate(head), (6, 36))
        totals = {}
        add_wire_materials(totals, 'dde_de', '185 SAC', '185 SAC', 1, '185 SAC')
        values = {row[CODE_COL]: row[TOTAL_COL] for row in totals.values()}
        self.assertEqual(values['1020260205'], 9)
        self.assertEqual(values['1030140011'], 9)
        self.assertEqual(values['1020410027'], 3)
        self.assertEqual(values['1020180001'], 3)
        self.assertEqual(values['1020180008'], 3)

    def test_page_hardware_export(self):
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame([{
            SIZE_COL: 'Dis & SF6', HEAD_COL: 'SF6', MATERIAL_COL: 'SF6 SET', CODE_COL: 'set24009', QTY_COL: 1,
        }])
        workbook.set_df = pd.DataFrame([
            {"Set": 'set24009', CODE_COL: '1020330006', "คำอธิบาย": 'HOTLINE BAIL-CLAMP,MAIN 70-185 SQ.MM.', "ติดตั้ง": 3},
            {"Set": 'set24009', CODE_COL: '1020330104', "คำอธิบาย": 'HOTLINE CLAMP,MAIN35-185,TAP50-185SQ.MM.', "ติดตั้ง": 3},
        ])
        pages = [[{
            'size': '14.3', 'head': 'DP,DDE st.4.5m', 'count': '1+1',
            'wire1': '185 SAC', 'wire2': '185 SAC',
        }, {
            'size': 'Dis & SF6', 'head': 'SF6', 'count': '1',
        }], [{
            'size': '12.2', 'head': 'DP,DE st.4.5m', 'count': '1',
            'wire1': '50 SAC',
        }]]
        exported = load_workbook(BytesIO(workbook.export_page_hardware(pages)), data_only=True)
        sheet = exported['ลูกถ้วยและอุปกรณ์']
        values = {(row[0], row[1], str(row[2] or '')): (row[3], row[4]) for row in sheet.iter_rows(min_row=4, values_only=True)}
        self.assertEqual(values[('ลูกถ้วย', 'ลูกถ้วยตั้ง', '')], (24, 6))
        self.assertEqual(values[('ลูกถ้วย', 'ลูกถ้วยนอน', '')], (48, 12))
        self.assertEqual(values[('อุปกรณ์ยึดสาย', 'PREFORMED D/E,SAC 22kV 185sq.mm. 29.78mm', '1020260205')], (12, None))
        self.assertEqual(values[('อุปกรณ์ยึดสาย', 'PREFORMED D/E,SAC 22kV 50sq.mm. 21.80mm', '1020260202')], (None, 3))
        self.assertEqual(values[('อุปกรณ์ยึดสาย', 'CLEVIS,THIMBLE,FOR PREFORMED DEAD-END', '1030140011')], (12, 3))
        self.assertEqual(values[('อุปกรณ์ยึดสาย', 'SLEEVE,TENSIONLESS COM.AL 185 SQ.MM.', '1020410027')], (6, None))
        self.assertEqual(values[('อุปกรณ์ยึดสาย', 'HOTLINE BAIL-CLAMP,MAIN 70-185 SQ.MM.', '1020330006')], (3, None))
        self.assertEqual(values[('อุปกรณ์ยึดสาย', 'HOTLINE CLAMP,MAIN35-185,TAP50-185SQ.MM.', '1020330104')], (3, None))
        self.assertEqual(sheet.freeze_panes, 'D4')
        detail_sheet = exported['ที่มารายการ']
        sf6_rows = [row for row in detail_sheet.iter_rows(min_row=2, values_only=True) if row[1] == 'SF6']
        self.assertEqual(len(sf6_rows), 2)
        self.assertTrue(all(row[2] == 'set24009' and row[5] == 3 for row in sf6_rows))

    def test_python_insulator_rules_match_confirmed_cases(self):
        expected = {
            'BA.st4.5m': (4, 12), '2BA st.4.5m': (12, 24),
            'SP,DDE.BL st.4.5m': (3, 24), 'DP,DDE.BL st.4.5m': (6, 24),
            'DP,DDE st.4.5m': (12, 24), 'DP,DE st.4.5m': (6, 12),
            'DDE.st 3m, LAT.SLK': (12, 36), 'DE.CON 1-P': (4, 8),
        }
        for head, rate in expected.items():
            self.assertEqual(insulator_rate(head), rate)

    def test_sf6_set_hardware_reconciles_with_wire_totals(self):
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame([{
            SIZE_COL: 'Dis & SF6', HEAD_COL: 'SF6', MATERIAL_COL: 'SF6 SET', CODE_COL: 'set24009', QTY_COL: 1,
        }])
        workbook.set_df = pd.DataFrame([
            {"Set": 'set24009', CODE_COL: '1020330006', "คำอธิบาย": 'HOTLINE BAIL-CLAMP,MAIN 70-185 SQ.MM.', "ติดตั้ง": 3},
            {"Set": 'set24009', CODE_COL: '1020330104', "คำอธิบาย": 'HOTLINE CLAMP,MAIN35-185,TAP50-185SQ.MM.', "ติดตั้ง": 3},
        ])
        pages = [[
            {'size': '14.3', 'head': 'BA', 'count': 12, 'wire1': '185 SAC', 'wire2': '50 SAC'},
            {'size': 'Dis & SF6', 'head': 'SF6', 'count': 1},
        ]]
        exported = load_workbook(BytesIO(workbook.export_page_hardware(pages)), data_only=True)
        rows = list(exported['ลูกถ้วยและอุปกรณ์'].iter_rows(min_row=4, values_only=True))
        by_code = {str(row[2] or ''): row[3] for row in rows}
        self.assertEqual(by_code['1020330006'], 39)
        self.assertEqual(by_code['1020330104'], 39)

    def test_page_insulator_export_shows_totals_and_sources(self):
        workbook = MaterialWorkbook()
        pages = [[
            {'size': '14.3', 'head': 'BA', 'count': '1+1'},
            {'size': '14.3', 'head': 'DE', 'count': 1},
        ], [
            {'size': '12.2', 'head': 'SP 1-P', 'count': 3},
        ]]
        exported = load_workbook(BytesIO(workbook.export_page_insulators(pages)), data_only=True)
        summary = list(exported['สรุปลูกถ้วยแยกหน้า'].iter_rows(min_row=2, values_only=True))
        self.assertEqual(summary[0], (1, 8, 36, 44))
        self.assertEqual(summary[1], (2, 6, 0, 6))
        self.assertEqual(summary[2], ('รวมทุกหน้า', 14, 36, 50))
        details = list(exported['ที่มาลูกถ้วย'].iter_rows(min_row=2, values_only=True))
        self.assertEqual(details[0], (1, '14.3', 'BA', 2, 4, 12, 8, 24))
        self.assertEqual(details[1], (1, '14.3', 'DE', 1, 0, 12, 0, 12))
        self.assertEqual(details[2], (2, '12.2', 'SP 1-P', 3, 2, 0, 6, 0))

    def test_page_crossarm_export_expands_sets_and_applies_adjustments(self):
        workbook = MaterialWorkbook()
        workbook.base_df = pd.DataFrame([
            {SIZE_COL: '14.3', HEAD_COL: 'BA', MATERIAL_COL: 'BA SET', CODE_COL: 'set-ba', QTY_COL: 1},
            {SIZE_COL: '14.3', HEAD_COL: 'BA', MATERIAL_COL: 'คอน 2.5 เมตร', CODE_COL: '1000110004', QTY_COL: 2},
            {SIZE_COL: '14.3', HEAD_COL: 'BA', MATERIAL_COL: 'คอน 2 เมตร', CODE_COL: '1000110003', QTY_COL: -2},
        ])
        workbook.set_df = pd.DataFrame([
            {'Set': 'set-ba', CODE_COL: '1000110003', 'คำอธิบาย': 'คอนคอนกรีต 2 เมตร', 'ติดตั้ง': 2},
            {'Set': 'set-ba', CODE_COL: '1010000302', 'คำอธิบาย': 'STEEL CHANNEL, 150x75x6.5 MM. 4,500 MM.LONG', 'ติดตั้ง': 2},
            {'Set': 'set-ba', CODE_COL: '1010200001', 'คำอธิบาย': 'เหล็กค้ำคอน', 'ติดตั้ง': 4},
        ])
        pages = [[{'size': '14.3', 'head': 'BA', 'count': 3}]]
        exported = load_workbook(BytesIO(workbook.export_page_crossarms(pages)), data_only=True)
        summary = list(exported['สรุปคอนแยกหน้า'].iter_rows(min_row=2, values_only=True))
        self.assertEqual(summary, [
            ('คอน 2.5 เมตร', '1000110004', 6, 6),
            ('STEEL CHANNEL, 150x75x6.5 MM. 4,500 MM.LONG', '1010000302', 6, 6),
        ])
        details = list(exported['ที่มาคอน'].iter_rows(min_row=2, values_only=True))
        self.assertEqual(details, [
            (1, '14.3', 'BA', 3, 'คอน 2.5 เมตร', '1000110004', 6),
            (1, '14.3', 'BA', 3, 'STEEL CHANNEL, 150x75x6.5 MM. 4,500 MM.LONG', '1010000302', 6),
        ])


if __name__ == '__main__':
    unittest.main()
