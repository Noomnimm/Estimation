from __future__ import annotations

import ast
import re
from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


BASE_SHEET = "BaseData"
SET_SHEET = "วัสดุทั้งหมด"

SIZE_COL = "ขนาดเสา (m)"
HEAD_COL = "รหัสหัวเสา"
MATERIAL_COL = "รายการวัสดุ"
CODE_COL = "รหัสพัสดุ"
QTY_COL = "จำนวน"
TOTAL_COL = "จำนวนรวม"
DEPARTMENT_COL = "แผนก"
DEFAULT_DEPARTMENT = "แผนกแรงสูง"
INSULATOR_UPRIGHT_COL = "ลูกถ้วยตั้ง"
INSULATOR_HORIZONTAL_COL = "ลูกถ้วยนอน"

SET_COL = "Set"
SET_DESC_COL = "คำอธิบาย"
SET_INSTALL_COL = "ติดตั้ง"

WIRE_MATERIALS = {
    "50 PIC": ("PREFORMED DEAD END,FOR AL PARTIALLY INSULATED CONDUCTOR 22 KV. 50 SQ.MM.", "1020250001", "50", True),
    "95 PIC": ("PREFORMED DEAD END,FOR AL PARTIALLY INSULATED CONDUCTOR 22 KV. 95 SQ.MM.", "1020250002", "95", True),
    "185 PIC": ("PREFORMED DEAD END,FOR AL PARTIALLY INSULATED CONDUCTOR 22 KV. 185 SQ.MM.", "1020250004", "185", True),
    "50 SAC": ("PREFORMED D/E,SAC 22kV 50sq.mm. 21.80mm", "1020260202", "50", True),
    "185 SAC": ("PREFORMED D/E,SAC 22kV 185sq.mm. 29.78mm", "1020260205", "185", True),
    "50 A": ("CLAMP,STRAIN,STRAIGHT TYPE,AL 35-70 sq.mm.ACSR 35-50 SQ.MM.", "1030110000", "50", False),
    "50 ACSR": ("CLAMP,STRAIN,STRAIGHT TYPE,AL 35-70 sq.mm.ACSR 35-50 SQ.MM.", "1030110000", "50", False),
    "185 ACSR": ("CLAMP,STRAIN,STRAIGHT TYPE FOR ACSR.120-185 sq.mm.", "1030110007", "185", False),
    "185 A": ("CLAMP,STRAIN,STRAIGHT TYPE,FOR AL 185 SQ.MM.", "1030110004", "185", False),
}

TENSIONLESS_MATERIALS = {
    "50": ("CONNECTOR,SPLICE,COMPRESSION TYPE,TENSIONLESS AL 50 SQ.MM.", "1020410002"),
    "185": ("CONNECTOR,SPLICE,COMPRESSION TYPE,TENSIONLESS AL 185 SQ.MM.", "1020410027"),
}

PG3_MATERIAL = (
    "CONNECTOR,PARALLEL GROOVE,TRIPLE BOLT,AL,AL-ALLOY AND ACSR 70-185 SQ.MM.",
    "1020300103",
)
HOTLINE_CLAMP_MATERIAL = ("HOTLINE CLAMP,MAIN35-185,TAP50-185SQ.MM.", "1020330104")
BAIL_CLAMP_MATERIAL = ("HOTLINE BAIL-CLAMP,MAIN 70-185 SQ.MM.", "1020330006")
CLEVIS_MATERIAL = ("CLEVIS,THIMBLE,FOR PREFORMED DEAD-END", "1030140011")
TENSIONLESS_TAPES = (
    ("PVC TAPE", "1020180001"),
    ("ERP TAPE", "1020180008"),
)


class MaterialWorkbook:
    def __init__(self) -> None:
        self.base_df: pd.DataFrame | None = None
        self.set_df: pd.DataFrame | None = None
        self.summary: list[dict[str, Any]] = []
        self.base_path: Path | None = None
        self.set_path: Path | None = None

    def load_base(self, path: str | Path, department: str = DEFAULT_DEPARTMENT) -> dict[str, Any]:
        df = pd.read_excel(path, sheet_name=BASE_SHEET)
        require_columns(df, [SIZE_COL, HEAD_COL, MATERIAL_COL, CODE_COL, QTY_COL], "BaseData")
        df = df[[SIZE_COL, HEAD_COL, MATERIAL_COL, CODE_COL, QTY_COL]].copy()
        df[DEPARTMENT_COL] = department
        df[INSULATOR_UPRIGHT_COL] = pd.NA
        df[INSULATOR_HORIZONTAL_COL] = pd.NA
        df = df.dropna(subset=[SIZE_COL, HEAD_COL, CODE_COL])
        self.base_df = df
        self.base_path = Path(path)
        self.summary = []
        return {
            "file": self.base_path.name,
            "rows": int(len(df)),
            "sizes": self.get_sizes(),
        }

    def load_set(self, path: str | Path) -> dict[str, Any]:
        df = read_set_sheet(path)
        require_columns(df, [SET_COL, CODE_COL, SET_DESC_COL, SET_INSTALL_COL], "SET")
        df = df[[SET_COL, CODE_COL, SET_DESC_COL, SET_INSTALL_COL]].copy()
        df = df.dropna(subset=[SET_COL, CODE_COL])
        self.set_df = df
        self.set_path = Path(path)
        return {
            "file": self.set_path.name,
            "rows": int(len(df)),
            "sets": int(df[SET_COL].astype(str).str.strip().str.lower().nunique()),
        }

    def load_keycode_catalog(self, path: str | Path, department: str, code_overrides: dict[str, str] | None = None) -> int:
        raw = pd.read_excel(path, header=None)
        code_overrides = code_overrides or {}
        rows = []
        for _, source in raw.iloc[7:].iterrows():
            keycode = clean_text(source.iloc[1] if len(source) > 1 else "")
            description = clean_text(source.iloc[2] if len(source) > 2 else "")
            set_code = clean_text(source.iloc[5] if len(source) > 5 else "")
            material_code = clean_text(source.iloc[8] if len(source) > 8 else "")
            code = code_overrides.get(keycode, set_code or material_code)
            if not keycode or not description or not code:
                continue
            rows.append({
                SIZE_COL: keycode, HEAD_COL: description, MATERIAL_COL: description,
                CODE_COL: code, QTY_COL: 1.0, DEPARTMENT_COL: department,
                INSULATOR_UPRIGHT_COL: pd.NA, INSULATOR_HORIZONTAL_COL: pd.NA,
            })
        if rows:
            self.base_df = pd.concat([self.base_df, pd.DataFrame(rows)], ignore_index=True)
        return len(rows)

    def load_department_base(self, path: str | Path, department: str, size_label: str) -> int:
        """Append a simple department workbook: รายการ, รายการวัสดุ, รหัสพัสดุ, จำนวน."""
        source = pd.read_excel(path, dtype={"รหัสพัสดุ": str})
        require_columns(source, ["รายการ", MATERIAL_COL, CODE_COL, QTY_COL], department)
        rows = []
        for _, item in source.iterrows():
            head = clean_text(item["รายการ"])
            material = clean_text(item[MATERIAL_COL])
            code = clean_text(item[CODE_COL]).replace("-", "")
            if not head or not material or not code:
                continue
            rows.append({
                SIZE_COL: size_label, HEAD_COL: head, MATERIAL_COL: material,
                CODE_COL: code, QTY_COL: parse_number(item[QTY_COL]), DEPARTMENT_COL: department,
                INSULATOR_UPRIGHT_COL: pd.NA, INSULATOR_HORIZONTAL_COL: pd.NA,
            })
        if rows:
            self.base_df = pd.concat([self.base_df, pd.DataFrame(rows)], ignore_index=True)
        return len(rows)

    def get_departments(self) -> list[str]:
        configured = [DEFAULT_DEPARTMENT, "แผนกแรงสูง TAC", "แผนกหม้อแปลง", "แผนกสายส่ง"]
        return configured

    def get_sizes(self, department: str = DEFAULT_DEPARTMENT) -> list[str]:
        if self.base_df is None:
            return []
        matches = self.base_df[self._department_mask(department)]
        sizes = matches[SIZE_COL].dropna().astype(str).str.strip().unique().tolist()
        return sorted(sizes, key=natural_key)

    def get_heads(self, size: str, department: str = DEFAULT_DEPARTMENT) -> list[str]:
        if self.base_df is None:
            raise ValueError("ยังไม่ได้โหลดไฟล์ BaseData")
        selected = str(size).strip()
        matches = self.base_df[
            (self.base_df[SIZE_COL].astype(str).str.strip() == selected)
            & self._department_mask(department)
        ]
        heads = matches[HEAD_COL].dropna().astype(str).str.strip().unique().tolist()
        return sorted(heads, key=natural_key)

    def get_insulator_rate(self, size: str, head: str, department: str = DEFAULT_DEPARTMENT) -> tuple[float, float]:
        if self.base_df is not None and INSULATOR_UPRIGHT_COL in self.base_df.columns:
            matches = self.base_df[
                (self.base_df[SIZE_COL].astype(str).str.strip() == str(size).strip())
                & (self.base_df[HEAD_COL].astype(str).str.strip() == str(head).strip())
                & self._department_mask(department)
            ]
            for _, row in matches.iterrows():
                upright = row.get(INSULATOR_UPRIGHT_COL)
                horizontal = row.get(INSULATOR_HORIZONTAL_COL)
                if not pd.isna(upright) or not pd.isna(horizontal):
                    return parse_number(upright), parse_number(horizontal)
        return insulator_rate(head)

    def get_status(self) -> dict[str, Any]:
        base = None
        if self.base_df is not None and self.base_path is not None:
            base = {
                "file": self.base_path.name,
                "rows": int(len(self.base_df)),
                "sizes": self.get_sizes(),
                "departments": self.get_departments(),
            }

        set_data = None
        if self.set_df is not None and self.set_path is not None:
            set_data = {
                "file": self.set_path.name,
                "rows": int(len(self.set_df)),
                "sets": int(self.set_df[SET_COL].astype(str).str.strip().str.lower().nunique()),
            }

        return {"base": base, "set": set_data}

    def calculate(self, pages: list[list[dict[str, Any]]]) -> dict[str, Any]:
        if self.base_df is None:
            raise ValueError("ยังไม่ได้โหลดไฟล์ BaseData")

        totals: dict[tuple[str, str], dict[str, Any]] = {}
        input_count = 0
        matched_rows = 0

        for page_number, page in enumerate(pages, start=1):
            for row_number, item in enumerate(page, start=1):
                size = str(item.get("size", "")).strip()
                head = str(item.get("head", "")).strip()
                department = clean_text(item.get("department")) or DEFAULT_DEPARTMENT
                count = parse_number(item.get("count"))
                if not size or not head or count <= 0:
                    continue

                input_count += 1
                high_voltage = department in {DEFAULT_DEPARTMENT, "แผนกแรงสูง TAC"}
                wire_kind = classify_wire_head(head) if high_voltage else None
                wire1 = clean_text(item.get("wire1"))
                wire2 = clean_text(item.get("wire2"))
                validate_wire_selection(wire_kind, wire1, wire2, page_number, row_number, head)
                lat_wire = clean_text(item.get("latWire"))
                if has_combined_lat(head):
                    validate_wire_selection("de", lat_wire, "", page_number, row_number, f"{head} — LAT.SLK")
                matches = self.base_df[
                    (self.base_df[SIZE_COL].astype(str).str.strip() == size)
                    & (self.base_df[HEAD_COL].astype(str).str.strip() == head)
                    & self._department_mask(department)
                ]

                for _, row in matches.iterrows():
                    material = clean_text(row[MATERIAL_COL])
                    code = clean_text(row[CODE_COL])
                    amount = parse_number(row[QTY_COL]) * count
                    if not code or amount == 0:
                        continue
                    add_material(totals, material, code, amount)
                    matched_rows += 1

                matched_rows += add_wire_materials(totals, wire_kind, wire1, wire2, count * wire_head_multiplier(head))
                if high_voltage and has_combined_lat(head):
                    matched_rows += add_wire_materials(totals, "de", lat_wire, "", count)

        self.summary = sorted(totals.values(), key=lambda r: (str(r[CODE_COL]).lower(), str(r[MATERIAL_COL]).lower()))
        return {
            "items": self.summary,
            "inputRows": input_count,
            "matchedRows": matched_rows,
            "summaryRows": len(self.summary),
        }

    def _department_mask(self, department: str) -> pd.Series:
        """Keep older in-memory/test BaseData compatible with the new department column."""
        if self.base_df is None:
            return pd.Series(dtype=bool)
        if DEPARTMENT_COL not in self.base_df.columns:
            return pd.Series(department == DEFAULT_DEPARTMENT, index=self.base_df.index)
        return self.base_df[DEPARTMENT_COL].astype(str).str.strip() == department

    def expand_set(self) -> dict[str, Any]:
        if self.set_df is None:
            raise ValueError("ยังไม่ได้โหลดไฟล์ SET")
        if not self.summary:
            raise ValueError("ยังไม่มีผลคำนวณให้แตก SET")

        expanded: list[dict[str, Any]] = []
        set_found = 0
        set_missing: list[str] = []
        expanded_lines = 0

        set_lookup = self.set_df.copy()
        set_lookup["_set_key"] = set_lookup[SET_COL].astype(str).str.strip().str.lower()

        for row in self.summary:
            code = clean_text(row[CODE_COL])
            qty = parse_number(row[TOTAL_COL])
            key = code.lower()
            if key.startswith("set"):
                matches = set_lookup[set_lookup["_set_key"] == key]
                if matches.empty:
                    set_missing.append(code)
                    continue
                set_found += 1
                for _, item in matches.iterrows():
                    expanded.append(
                        {
                            MATERIAL_COL: clean_text(item[SET_DESC_COL]),
                            CODE_COL: clean_text(item[CODE_COL]),
                            TOTAL_COL: parse_number(item[SET_INSTALL_COL]) * qty,
                        }
                    )
                    expanded_lines += 1
            else:
                expanded.append(
                    {
                        MATERIAL_COL: clean_text(row[MATERIAL_COL]),
                        CODE_COL: code,
                        TOTAL_COL: qty,
                    }
                )

        self.summary = group_summary(expanded)
        return {
            "items": self.summary,
            "setFound": set_found,
            "setMissing": set_missing,
            "expandedLines": expanded_lines,
            "summaryRows": len(self.summary),
        }

    def export_summary(self) -> bytes:
        if not self.summary:
            raise ValueError("ยังไม่มีผลลัพธ์สำหรับ export")

        output = BytesIO()
        df = pd.DataFrame(self.summary, columns=[MATERIAL_COL, CODE_COL, TOTAL_COL])
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Summary")
        return output.getvalue()

    def export_page_summary(self, pages: list[list[dict[str, Any]]]) -> bytes:
        totals: dict[tuple[str, str], dict[int, float]] = {}
        for page_number, page in enumerate(pages, start=1):
            for item in page:
                size = clean_text(item.get("size"))
                head = clean_text(item.get("head"))
                department = clean_text(item.get("department")) or DEFAULT_DEPARTMENT
                count = parse_number(item.get("count"))
                if not size or not head or count <= 0:
                    continue
                key = (head, size)
                if key not in totals:
                    totals[key] = {}
                totals[key][page_number] = totals[key].get(page_number, 0.0) + count

        if not totals:
            raise ValueError("ยังไม่มีข้อมูลแต่ละหน้าสำหรับ export")

        page_columns = [f"หน้า {page_number}" for page_number in range(1, len(pages) + 1)]
        rows = []
        for (head, size), page_totals in sorted(totals.items(), key=lambda item: (natural_key(item[0][1]), natural_key(item[0][0]))):
            row: dict[str, Any] = {HEAD_COL: head, "เสา": size}
            for page_number, column in enumerate(page_columns, start=1):
                row[column] = page_totals.get(page_number)
            rows.append(row)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            columns = [HEAD_COL, "เสา", *page_columns]
            pd.DataFrame(rows, columns=columns).to_excel(writer, index=False, sheet_name="สรุปแต่ละหน้า", startrow=2)
            sheet = writer.sheets["สรุปแต่ละหน้า"]
            last_column = get_column_letter(len(columns))
            sheet.merge_cells(f"A1:{last_column}1")
            title = sheet["A1"]
            title.value = "สรุปรายการหัวเสาแยกตามหน้า"
            title.font = Font(name="Tahoma", size=16, bold=True, color="FFFFFF")
            title.fill = PatternFill("solid", fgColor="0F766E")
            title.alignment = Alignment(horizontal="center", vertical="center")
            sheet.row_dimensions[1].height = 28

            header_fill = PatternFill("solid", fgColor="DDEDEA")
            border = Border(bottom=Side(style="thin", color="AAB7C4"))
            for cell in sheet[3]:
                cell.font = Font(name="Tahoma", bold=True, color="1D242D")
                cell.fill = header_fill
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for row in sheet.iter_rows(min_row=4, max_row=sheet.max_row, max_col=len(columns)):
                for cell in row:
                    cell.font = Font(name="Tahoma", size=10)
                for cell in row[2:]:
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "#,##0.###"

            sheet.column_dimensions["A"].width = 32
            sheet.column_dimensions["B"].width = 14
            for column_index in range(3, len(columns) + 1):
                sheet.column_dimensions[get_column_letter(column_index)].width = 13
            sheet.freeze_panes = "C4"
            sheet.auto_filter.ref = f"A3:{last_column}{sheet.max_row}"
            sheet.sheet_view.showGridLines = False

        return output.getvalue()

    def export_page_hardware(self, pages: list[list[dict[str, Any]]]) -> bytes:
        totals: dict[tuple[str, str, str], dict[int, float]] = {}
        details: list[dict[str, Any]] = []
        hardware_codes = {
            *(wire_details[1] for wire_details in WIRE_MATERIALS.values()),
            *(code for _, code in TENSIONLESS_MATERIALS.values()),
            *(code for _, code in TENSIONLESS_TAPES),
            PG3_MATERIAL[1], HOTLINE_CLAMP_MATERIAL[1], BAIL_CLAMP_MATERIAL[1], CLEVIS_MATERIAL[1],
        }
        set_lookup = None
        if self.set_df is not None:
            set_lookup = self.set_df.copy()
            set_lookup["_set_key"] = set_lookup[SET_COL].astype(str).str.strip().str.lower()

        def add_page_item(group: str, material: str, code: str, page_number: int, amount: float) -> None:
            if amount == 0:
                return
            key = (group, material, code)
            totals.setdefault(key, {})[page_number] = totals.setdefault(key, {}).get(page_number, 0.0) + amount

        for page_number, page in enumerate(pages, start=1):
            for row_number, item in enumerate(page, start=1):
                size = clean_text(item.get("size"))
                head = clean_text(item.get("head"))
                department = clean_text(item.get("department")) or DEFAULT_DEPARTMENT
                count = parse_number(item.get("count"))
                if not size or not head or count <= 0:
                    continue

                upright, horizontal = self.get_insulator_rate(size, head, department)
                add_page_item("ลูกถ้วย", "ลูกถ้วยตั้ง", "", page_number, upright * count)
                add_page_item("ลูกถ้วย", "ลูกถ้วยนอน", "", page_number, horizontal * count)
                if upright:
                    details.append({"หน้า": page_number, HEAD_COL: head, "ที่มา": "เกณฑ์ลูกถ้วย", MATERIAL_COL: "ลูกถ้วยตั้ง", CODE_COL: "", TOTAL_COL: upright * count})
                if horizontal:
                    details.append({"หน้า": page_number, HEAD_COL: head, "ที่มา": "เกณฑ์ลูกถ้วย", MATERIAL_COL: "ลูกถ้วยนอน", CODE_COL: "", TOTAL_COL: horizontal * count})

                if self.base_df is not None and set_lookup is not None:
                    matches = self.base_df[
                        (self.base_df[SIZE_COL].astype(str).str.strip() == size)
                        & (self.base_df[HEAD_COL].astype(str).str.strip() == head)
                        & self._department_mask(department)
                    ]
                    for _, base_row in matches.iterrows():
                        set_code = clean_text(base_row[CODE_COL])
                        if not set_code.lower().startswith("set"):
                            continue
                        set_quantity = parse_number(base_row[QTY_COL]) * count
                        for _, set_row in set_lookup[set_lookup["_set_key"] == set_code.lower()].iterrows():
                            code = clean_text(set_row[CODE_COL])
                            if code not in hardware_codes:
                                continue
                            material = clean_text(set_row[SET_DESC_COL])
                            amount = parse_number(set_row[SET_INSTALL_COL]) * set_quantity
                            add_page_item("อุปกรณ์ยึดสาย", material, code, page_number, amount)
                            details.append({"หน้า": page_number, HEAD_COL: head, "ที่มา": set_code, MATERIAL_COL: material, CODE_COL: code, TOTAL_COL: amount})

                high_voltage = department in {DEFAULT_DEPARTMENT, "แผนกแรงสูง TAC"}
                wire_kind = classify_wire_head(head) if high_voltage else None
                wire1 = clean_text(item.get("wire1"))
                wire2 = clean_text(item.get("wire2"))
                validate_wire_selection(wire_kind, wire1, wire2, page_number, row_number, head)
                wire_totals: dict[tuple[str, str], dict[str, Any]] = {}
                add_wire_materials(wire_totals, wire_kind, wire1, wire2, count * wire_head_multiplier(head))

                lat_wire = clean_text(item.get("latWire"))
                if high_voltage and has_combined_lat(head):
                    validate_wire_selection("de", lat_wire, "", page_number, row_number, f"{head} — LAT.SLK")
                    add_wire_materials(wire_totals, "de", lat_wire, "", count)

                for material in wire_totals.values():
                    amount = parse_number(material[TOTAL_COL])
                    add_page_item(
                        "อุปกรณ์ยึดสาย",
                        clean_text(material[MATERIAL_COL]),
                        clean_text(material[CODE_COL]),
                        page_number,
                        amount,
                    )
                    details.append({
                        "หน้า": page_number, HEAD_COL: head, "ที่มา": "ช่องเลือกสาย",
                        MATERIAL_COL: clean_text(material[MATERIAL_COL]), CODE_COL: clean_text(material[CODE_COL]), TOTAL_COL: amount,
                    })

        if not totals:
            raise ValueError("ยังไม่มีข้อมูลลูกถ้วยหรืออุปกรณ์ยึดสายสำหรับ export")

        page_columns = [f"หน้า {number}" for number in range(1, len(pages) + 1)]
        rows = []
        for (group, material, code), page_totals in sorted(totals.items(), key=lambda item: (item[0][0], item[0][2], item[0][1])):
            row: dict[str, Any] = {"ประเภท": group, MATERIAL_COL: material, CODE_COL: code}
            for page_number, column in enumerate(page_columns, start=1):
                row[column] = page_totals.get(page_number)
            rows.append(row)

        output = BytesIO()
        columns = ["ประเภท", MATERIAL_COL, CODE_COL, *page_columns]
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame(rows, columns=columns).to_excel(writer, index=False, sheet_name="ลูกถ้วยและอุปกรณ์", startrow=2)
            sheet = writer.sheets["ลูกถ้วยและอุปกรณ์"]
            last_column = get_column_letter(len(columns))
            sheet.merge_cells(f"A1:{last_column}1")
            title = sheet["A1"]
            title.value = "สรุปลูกถ้วยและอุปกรณ์ยึดสายแยกตามหน้า"
            title.font = Font(name="Tahoma", size=16, bold=True, color="FFFFFF")
            title.fill = PatternFill("solid", fgColor="63318A")
            title.alignment = Alignment(horizontal="center", vertical="center")
            sheet.row_dimensions[1].height = 28
            for cell in sheet[3]:
                cell.font = Font(name="Tahoma", bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="4B216E")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for row in sheet.iter_rows(min_row=4, max_row=sheet.max_row, max_col=len(columns)):
                for cell in row:
                    cell.font = Font(name="Tahoma", size=10)
                for cell in row[3:]:
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "#,##0.###"
            sheet.column_dimensions["A"].width = 18
            sheet.column_dimensions["B"].width = 64
            sheet.column_dimensions["C"].width = 18
            for column_index in range(4, len(columns) + 1):
                sheet.column_dimensions[get_column_letter(column_index)].width = 13
            sheet.freeze_panes = "D4"
            sheet.auto_filter.ref = f"A3:{last_column}{sheet.max_row}"
            sheet.sheet_view.showGridLines = False

            detail_columns = ["หน้า", HEAD_COL, "ที่มา", MATERIAL_COL, CODE_COL, TOTAL_COL]
            pd.DataFrame(details, columns=detail_columns).to_excel(writer, index=False, sheet_name="ที่มารายการ")
            detail_sheet = writer.sheets["ที่มารายการ"]
            for cell in detail_sheet[1]:
                cell.font = Font(name="Tahoma", bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="4B216E")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for row in detail_sheet.iter_rows(min_row=2, max_row=detail_sheet.max_row):
                for cell in row:
                    cell.font = Font(name="Tahoma", size=10)
                row[5].number_format = "#,##0.###"
            for column, width in {"A": 10, "B": 30, "C": 18, "D": 64, "E": 18, "F": 14}.items():
                detail_sheet.column_dimensions[column].width = width
            detail_sheet.freeze_panes = "A2"
            detail_sheet.auto_filter.ref = f"A1:F{detail_sheet.max_row}"
            detail_sheet.sheet_view.showGridLines = False
        return output.getvalue()

    def export_page_insulators(self, pages: list[list[dict[str, Any]]]) -> bytes:
        page_totals: dict[int, tuple[float, float]] = {}
        details: list[dict[str, Any]] = []

        for page_number, page in enumerate(pages, start=1):
            upright_total = 0.0
            horizontal_total = 0.0
            for item in page:
                size = clean_text(item.get("size"))
                head = clean_text(item.get("head"))
                department = clean_text(item.get("department")) or DEFAULT_DEPARTMENT
                if not size or not head:
                    continue
                count = parse_number(item.get("count"))
                if count == 0:
                    continue
                upright_rate, horizontal_rate = self.get_insulator_rate(size, head, department)
                upright = upright_rate * count
                horizontal = horizontal_rate * count
                if upright == 0 and horizontal == 0:
                    continue
                upright_total += upright
                horizontal_total += horizontal
                details.append({
                    "หน้า": page_number,
                    SIZE_COL: size,
                    HEAD_COL: head,
                    "จำนวนหัว": count,
                    "ตั้ง/หัว": upright_rate,
                    "นอน/หัว": horizontal_rate,
                    "ลูกถ้วยตั้ง": upright,
                    "ลูกถ้วยนอน": horizontal,
                })
            page_totals[page_number] = (upright_total, horizontal_total)

        if not details:
            raise ValueError("ยังไม่มีข้อมูลหัวเสาที่มีลูกถ้วยสำหรับ export")

        summary_rows = []
        for page_number, (upright, horizontal) in page_totals.items():
            summary_rows.append({"หน้า": page_number, "ลูกถ้วยตั้ง": upright, "ลูกถ้วยนอน": horizontal, "รวมลูกถ้วย": upright + horizontal})
        summary_rows.append({
            "หน้า": "รวมทุกหน้า",
            "ลูกถ้วยตั้ง": sum(value[0] for value in page_totals.values()),
            "ลูกถ้วยนอน": sum(value[1] for value in page_totals.values()),
            "รวมลูกถ้วย": sum(sum(value) for value in page_totals.values()),
        })

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="สรุปลูกถ้วยแยกหน้า")
            pd.DataFrame(details).to_excel(writer, index=False, sheet_name="ที่มาลูกถ้วย")
            for sheet_name, widths in {
                "สรุปลูกถ้วยแยกหน้า": [14, 18, 18, 18],
                "ที่มาลูกถ้วย": [10, 18, 34, 14, 14, 14, 16, 16],
            }.items():
                sheet = writer.sheets[sheet_name]
                for cell in sheet[1]:
                    cell.font = Font(name="Tahoma", bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="4B216E")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                    for cell in row:
                        cell.font = Font(name="Tahoma", size=10)
                    for cell in row[1:]:
                        cell.number_format = "#,##0.###"
                for index, width in enumerate(widths, start=1):
                    sheet.column_dimensions[get_column_letter(index)].width = width
                sheet.freeze_panes = "A2"
                sheet.auto_filter.ref = sheet.dimensions
                sheet.sheet_view.showGridLines = False
            total_row = writer.sheets["สรุปลูกถ้วยแยกหน้า"].max_row
            for cell in writer.sheets["สรุปลูกถ้วยแยกหน้า"][total_row]:
                cell.font = Font(name="Tahoma", bold=True)
                cell.fill = PatternFill("solid", fgColor="E8DDF2")
        return output.getvalue()

    def export_page_hardware_combined(self, pages: list[list[dict[str, Any]]]) -> bytes:
        """Combine hardware/preform and insulator-audit exports into separate sheets."""
        hardware = self.export_page_hardware(pages)
        insulators = self.export_page_insulators(pages)
        return merge_excel_workbooks(hardware, insulators)

    def export_page_crossarms(self, pages: list[list[dict[str, Any]]]) -> bytes:
        if self.base_df is None or self.set_df is None:
            raise ValueError("กรุณาโหลด BaseData และ SET ก่อน export คอน")

        set_lookup = self.set_df.copy()
        set_lookup["_set_key"] = set_lookup[SET_COL].astype(str).str.strip().str.lower()
        detail_totals: dict[tuple[int, str, str, str], dict[str, Any]] = {}

        def add_detail(page_number: int, size: str, head: str, head_count: float, material: str, code: str, amount: float) -> None:
            is_crossarm = code.startswith("10001") or bool(re.match(r"^STEEL\s*CHANNEL", material, re.IGNORECASE))
            if not is_crossarm or amount == 0:
                return
            key = (page_number, size, head, code)
            if key not in detail_totals:
                detail_totals[key] = {
                    "หน้า": page_number, SIZE_COL: size, HEAD_COL: head, "จำนวนหัว": head_count,
                    MATERIAL_COL: material, CODE_COL: code, TOTAL_COL: 0.0,
                }
            detail_totals[key][TOTAL_COL] += amount

        for page_number, page in enumerate(pages, start=1):
            for item in page:
                size = clean_text(item.get("size"))
                head = clean_text(item.get("head"))
                department = clean_text(item.get("department")) or DEFAULT_DEPARTMENT
                if not size or not head:
                    continue
                count = parse_number(item.get("count"))
                if count == 0:
                    continue
                matches = self.base_df[
                    (self.base_df[SIZE_COL].astype(str).str.strip() == size)
                    & (self.base_df[HEAD_COL].astype(str).str.strip() == head)
                    & self._department_mask(department)
                ]
                for _, base_row in matches.iterrows():
                    base_code = clean_text(base_row[CODE_COL])
                    base_quantity = parse_number(base_row[QTY_COL]) * count
                    if base_code.lower().startswith("set"):
                        set_rows = set_lookup[set_lookup["_set_key"] == base_code.lower()]
                        for _, set_row in set_rows.iterrows():
                            code = clean_text(set_row[CODE_COL])
                            add_detail(
                                page_number, size, head, count,
                                clean_text(set_row[SET_DESC_COL]), code,
                                base_quantity * parse_number(set_row[SET_INSTALL_COL]),
                            )
                    else:
                        add_detail(
                            page_number, size, head, count,
                            clean_text(base_row[MATERIAL_COL]), base_code, base_quantity,
                        )

        details = [row for row in detail_totals.values() if row[TOTAL_COL] != 0]
        if not details:
            raise ValueError("ไม่พบรายการคอนจากหัวเสาที่เลือก")

        page_count = max(len(pages), 1)
        summary_totals: dict[str, dict[str, Any]] = {}
        for row in details:
            code = row[CODE_COL]
            page_number = int(row["หน้า"])
            summary = summary_totals.setdefault(code, {MATERIAL_COL: row[MATERIAL_COL], "pages": {}})
            summary["pages"][page_number] = summary["pages"].get(page_number, 0.0) + parse_number(row[TOTAL_COL])
        summary_rows = []
        for code, summary in sorted(summary_totals.items()):
            amounts = summary["pages"]
            row: dict[str, Any] = {MATERIAL_COL: summary[MATERIAL_COL], CODE_COL: code}
            for page_number in range(1, page_count + 1):
                row[f"หน้า {page_number}"] = amounts.get(page_number, 0.0)
            row[TOTAL_COL] = sum(amounts.values())
            summary_rows.append(row)

        detail_columns = ["หน้า", SIZE_COL, HEAD_COL, "จำนวนหัว", MATERIAL_COL, CODE_COL, TOTAL_COL]
        summary_columns = [MATERIAL_COL, CODE_COL, *(f"หน้า {number}" for number in range(1, page_count + 1)), TOTAL_COL]
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame(summary_rows, columns=summary_columns).to_excel(writer, index=False, sheet_name="สรุปคอนแยกหน้า")
            pd.DataFrame(details, columns=detail_columns).sort_values(["หน้า", HEAD_COL, CODE_COL]).to_excel(
                writer, index=False, sheet_name="ที่มาคอน"
            )
            for sheet_name in ("สรุปคอนแยกหน้า", "ที่มาคอน"):
                sheet = writer.sheets[sheet_name]
                for cell in sheet[1]:
                    cell.font = Font(name="Tahoma", bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="4B216E")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                    for cell in row:
                        cell.font = Font(name="Tahoma", size=10)
                    for cell in row[2:]:
                        cell.number_format = "#,##0.###"
                sheet.freeze_panes = "C2" if sheet_name == "สรุปคอนแยกหน้า" else "A2"
                sheet.auto_filter.ref = sheet.dimensions
                sheet.sheet_view.showGridLines = False
            summary_sheet = writer.sheets["สรุปคอนแยกหน้า"]
            summary_sheet.column_dimensions["A"].width = 64
            summary_sheet.column_dimensions["B"].width = 18
            for column in range(3, summary_sheet.max_column + 1):
                summary_sheet.column_dimensions[get_column_letter(column)].width = 14
            detail_sheet = writer.sheets["ที่มาคอน"]
            for column, width in enumerate([10, 18, 34, 14, 64, 18, 16], start=1):
                detail_sheet.column_dimensions[get_column_letter(column)].width = width
        return output.getvalue()


def classify_wire_head(head: str) -> str | None:
    normalized = clean_text(head).upper()
    compact = re.sub(r"\s+", "", normalized)
    combined_rules = {
        "SP,DDE.BLST.4.5M": "dde_bl",
        "DP,DDE.BLST.4.5M": "dde_bl",
        "DP,DDEST.4.5M": "dde",
        "DP,DEST.4.5M": "de",
    }
    if compact in combined_rules:
        return combined_rules[compact]
    if re.match(r"^LAT\.SLK(?=$|\s)", normalized):
        return "de"
    if normalized.startswith("2"):
        # Only the insulator totals of +DE.CON assemblies have been confirmed.
        if "+" in normalized:
            return None
        normalized = normalized[1:]
    if normalized.startswith("DDE.BL"):
        return "dde_bl"
    if normalized.startswith("DDE"):
        return "dde"
    if normalized.startswith("DE"):
        return "de"
    if normalized.startswith("BA"):
        return "ba"
    return None


def insulator_rate(head: str) -> tuple[float, float]:
    name = re.sub(r"\s*,\s*", ",", clean_text(head).upper())
    compact = re.sub(r"\s+", "", name)
    exact = {
        "DP,DEST.4.5M": (6, 12), "DP,DDEST.4.5M": (12, 24),
        "DP,DDE.BLST.4.5M": (6, 24), "SP,DDE.BLST.4.5M": (3, 24),
        "2BAST.4.5M": (12, 24), "2BA.ST4.5M+DE.CON": (12, 36),
        "2DE.ST4.5+DE.CON": (6, 36), "DDE,DP.ST3.0M": (12, 24),
        "DDE.ST3M,LAT.SLK": (12, 36), "CCB,CCB": (6, 0),
    }
    if compact in exact:
        return exact[compact]
    if re.match(r"^(CTB|CSC)(?=$|[.\s])", name): return (0, 0)
    if re.match(r"^LAT\.SLK(?=$|[.\s])", name): return (6, 12)
    if re.match(r"^BA\.SLK(?=$|[.\s])", name): return (6, 0)
    if re.search(r"1\s*-?\s*P\b", name):
        for prefix, rate in (("DE.CON", (4, 8)), ("DDE.BL", (0, 16)), ("DDE", (4, 16)), ("BA", (2, 8)), ("SP", (2, 0))):
            if name.startswith(prefix): return rate
        return (0, 0)
    if re.match(r"^2(?:BA|DE|DDE|SP|DP)(?=$|[.\s])", name):
        upright, horizontal = insulator_rate(name[1:])
        return upright * 2, horizontal * 2
    if name == "SP บน,ล่าง": return (3, 0)
    if re.match(r"^CCB(?:\s|$)", name): return (6, 0) if "ประกบ" in name else (3, 0)
    for prefix, rate in (("DDE.BL", (0, 24)), ("DDE", (6, 24)), ("DE", (0, 12)), ("BA", (4, 12)), ("SP", (3, 0)), ("DP", (6, 0))):
        if name == prefix or name.startswith(prefix + ".") or name.startswith(prefix + " "):
            return rate
    return (0, 0)


def has_combined_lat(head: str) -> bool:
    return re.sub(r"\s+", "", clean_text(head).upper()) == "DDE.ST3M,LAT.SLK"


def wire_head_multiplier(head: str) -> float:
    normalized = clean_text(head).upper()
    if re.search(r"1\s*-?\s*P\b", normalized):
        return 2 / 3
    return 2 if normalized.startswith("2") else 1


def validate_wire_selection(
    wire_kind: str | None,
    wire1: str,
    wire2: str,
    page_number: int,
    row_number: int,
    head: str,
) -> None:
    if wire_kind is None:
        return
    if wire1 not in WIRE_MATERIALS:
        raise ValueError(f"หน้า {page_number} แถว {row_number} ({head}): กรุณาเลือกชนิดสายช่องแรก")
    if wire_kind in {"dde", "dde_bl", "ba"} and wire2 not in WIRE_MATERIALS:
        raise ValueError(f"หน้า {page_number} แถว {row_number} ({head}): กรุณาเลือกชนิดสายช่องที่สอง")
    if wire_kind == "dde" and conductor_group(wire1) != conductor_group(wire2):
        raise ValueError(f"หน้า {page_number} แถว {row_number} ({head}): สายซ้ายและขวาต้องมีขนาดเดียวกันสำหรับ Tensionless")
    if wire_kind == "dde" and conductor_group(wire1) not in TENSIONLESS_MATERIALS:
        raise ValueError(f"หน้า {page_number} แถว {row_number} ({head}): ยังไม่มีรหัส Tensionless สำหรับสายขนาด {conductor_group(wire1)}")


def conductor_group(wire: str) -> str:
    details = WIRE_MATERIALS.get(clean_text(wire))
    return details[2] if details else ""


def add_material(
    totals: dict[tuple[str, str], dict[str, Any]],
    material: str,
    code: str,
    amount: float,
) -> None:
    if not code or amount == 0:
        return
    key = (material, code)
    if key not in totals:
        totals[key] = {MATERIAL_COL: material, CODE_COL: code, TOTAL_COL: 0.0}
    totals[key][TOTAL_COL] += amount


def add_wire_materials(
    totals: dict[tuple[str, str], dict[str, Any]],
    wire_kind: str | None,
    wire1: str,
    wire2: str,
    count: float,
) -> int:
    if wire_kind is None:
        return 0

    selected_wires = [wire2] if wire_kind == "ba" else [wire1]
    if wire_kind in {"dde", "dde_bl"}:
        selected_wires.append(wire2)

    added = 0
    for wire in selected_wires:
        material, code, _, needs_clevis = WIRE_MATERIALS[wire]
        add_material(totals, material, code, 3 * count)
        added += 1
        if needs_clevis:
            clevis_material, clevis_code = CLEVIS_MATERIAL
            add_material(totals, clevis_material, clevis_code, 3 * count)
            added += 1

    if wire_kind == "dde":
        material, code = TENSIONLESS_MATERIALS[conductor_group(wire1)]
        tensionless_quantity = 3 * count
        add_material(totals, material, code, tensionless_quantity)
        added += 1
        for tape_material, tape_code in TENSIONLESS_TAPES:
            add_material(totals, tape_material, tape_code, tensionless_quantity)
            added += 1

    if wire_kind == "ba":
        if conductor_group(wire2) == "185":
            material, code = PG3_MATERIAL
            add_material(totals, material, code, 6 * count)
            added += 1
        else:
            for material, code in (HOTLINE_CLAMP_MATERIAL, BAIL_CLAMP_MATERIAL):
                add_material(totals, material, code, 3 * count)
                added += 1

    return added


def read_set_sheet(path: str | Path) -> pd.DataFrame:
    sheets = pd.read_excel(path, sheet_name=None)
    if SET_SHEET in sheets:
        selected = sheets[SET_SHEET]
        if {SET_COL, CODE_COL, SET_DESC_COL, SET_INSTALL_COL}.issubset(set(selected.columns)):
            return selected
        normalized = normalize_report_set(selected)
        if normalized is not None:
            return normalized
    for df in sheets.values():
        if {SET_COL, CODE_COL, SET_DESC_COL, SET_INSTALL_COL}.issubset(set(df.columns)):
            return df
        normalized = normalize_report_set(df)
        if normalized is not None:
            return normalized
    first = next(iter(sheets.values()))
    return first


def normalize_report_set(df: pd.DataFrame) -> pd.DataFrame | None:
    source_set_col = "รหัสอุปกรณ์ต่อชุด"
    required = {source_set_col, CODE_COL, SET_DESC_COL, SET_INSTALL_COL}
    if not required.issubset(set(df.columns)):
        return None

    set_values = df[source_set_col].ffill()
    item_rows = df[source_set_col].isna() & df[CODE_COL].notna()
    return pd.DataFrame(
        {
            SET_COL: set_values[item_rows],
            CODE_COL: df.loc[item_rows, CODE_COL],
            SET_DESC_COL: df.loc[item_rows, SET_DESC_COL],
            SET_INSTALL_COL: df.loc[item_rows, SET_INSTALL_COL],
        }
    )


def require_columns(df: pd.DataFrame, columns: list[str], label: str) -> None:
    missing = [col for col in columns if col not in df.columns]
    if missing:
        raise ValueError(f"ไฟล์ {label} ขาดคอลัมน์: {', '.join(missing)}")


def group_summary(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    totals: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        material = clean_text(row[MATERIAL_COL])
        code = clean_text(row[CODE_COL])
        amount = parse_number(row[TOTAL_COL])
        if not code or amount == 0:
            continue
        key = (material, code)
        if key not in totals:
            totals[key] = {MATERIAL_COL: material, CODE_COL: code, TOTAL_COL: 0.0}
        totals[key][TOTAL_COL] += amount
    return sorted(totals.values(), key=lambda r: (str(r[CODE_COL]).lower(), str(r[MATERIAL_COL]).lower()))


def clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def merge_excel_workbooks(*workbook_bytes: bytes) -> bytes:
    """Merge worksheets from generated workbooks while preserving their presentation."""
    target = Workbook()
    target.remove(target.active)
    used_names: set[str] = set()
    for data in workbook_bytes:
        source = load_workbook(BytesIO(data))
        for source_sheet in source.worksheets:
            title = source_sheet.title[:31]
            base, suffix = title, 2
            while title in used_names:
                marker = f" ({suffix})"
                title = f"{base[:31 - len(marker)]}{marker}"
                suffix += 1
            used_names.add(title)
            sheet = target.create_sheet(title)
            for row in source_sheet.iter_rows():
                for source_cell in row:
                    cell = sheet.cell(source_cell.row, source_cell.column, source_cell.value)
                    if source_cell.has_style:
                        cell.font = copy(source_cell.font)
                        cell.fill = copy(source_cell.fill)
                        cell.border = copy(source_cell.border)
                        cell.alignment = copy(source_cell.alignment)
                        cell.number_format = source_cell.number_format
                        cell.protection = copy(source_cell.protection)
            for merged_range in source_sheet.merged_cells.ranges:
                sheet.merge_cells(str(merged_range))
            for key, dimension in source_sheet.column_dimensions.items():
                sheet.column_dimensions[key].width = dimension.width
                sheet.column_dimensions[key].hidden = dimension.hidden
            for index, dimension in source_sheet.row_dimensions.items():
                sheet.row_dimensions[index].height = dimension.height
                sheet.row_dimensions[index].hidden = dimension.hidden
            sheet.freeze_panes = source_sheet.freeze_panes
            sheet.auto_filter.ref = source_sheet.auto_filter.ref
            sheet.sheet_view.showGridLines = source_sheet.sheet_view.showGridLines
    output = BytesIO()
    target.save(output)
    return output.getvalue()


def parse_number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, str):
        expression = value.strip()
        if not expression or len(expression) > 100:
            return 0.0
        try:
            return float(evaluate_add_sub(ast.parse(expression, mode="eval").body))
        except (SyntaxError, TypeError, ValueError):
            return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def evaluate_add_sub(node: ast.AST) -> float:
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
        return float(node.value)
    if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub)):
        left = evaluate_add_sub(node.left)
        right = evaluate_add_sub(node.right)
        return left + right if isinstance(node.op, ast.Add) else left - right
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
        value = evaluate_add_sub(node.operand)
        return value if isinstance(node.op, ast.UAdd) else -value
    raise ValueError("รองรับเฉพาะตัวเลข เครื่องหมาย +, - และวงเล็บ")


def natural_key(value: str) -> list[tuple[int, Any]]:
    # A punctuation dot in "st.4.5m" belongs to the text, not to 4.5.
    return [
        (0, float(part)) if index % 2 else (1, part.lower())
        for index, part in enumerate(re.split(r"([0-9]+(?:\.[0-9]+)?)", str(value)))
        if part
    ]
