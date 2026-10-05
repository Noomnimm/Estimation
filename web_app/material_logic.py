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
INSTALL_TOTAL_COL = "จำนวนติดตั้ง"
DEMOLITION_TOTAL_COL = "จำนวนรื้อถอน"
DEPARTMENT_COL = "แผนก"
DEFAULT_DEPARTMENT = "แผนกแรงสูง"
HIGH_VOLTAGE_TAC_DEPARTMENT = "แผนกแรงสูง TAC"
SURGE_FUSE_SIZE = "Surge & Fuse"
GROUND_SIZE = "Ground"
GROUND_ASSEMBLY = "(ASSEMBLY NO.9706, 9701B)"
GROUND_SET_CODES = {number: f"Set205{14 + number}" for number in range(1, 6)}
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
    "50 A": ("SLEEVE,TENSIONLESS AL 50 SQ.MM.", "1020410014"),
    "50 ACSR": ("SLEEVE,TENSIONLESS AL 50 SQ.MM.", "1020410014"),
    "50 PIC": ("SLEEVE,TENSIONLESS AL 50 SQ.MM.", "1020410014"),
    "185 A": ("SLEEVE,TENSIONLESS AL 185 SQ.MM.", "1020410017"),
    "185 ACSR": ("SLEEVE,TENSIONLESS AL 185 SQ.MM.", "1020410017"),
    "185 PIC": ("SLEEVE,TENSIONLESS AL 185 SQ.MM.", "1020410017"),
    "50 SAC": ("SLEEVE,TENSIONLESS COM.AL 50 SQ.MM.", "1020410022"),
    "185 SAC": ("SLEEVE,TENSIONLESS COM.AL 185 SQ.MM.", "1020410027"),
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

SURGE_ARRESTER_CROSSARM = {
    (False, False): ("LIGHTNING ARRESTER 20-21 kV. 5 kA", "1040000000"),
    (False, True): ("LIGHTNING ARRESTER,20-21 KV.10 KA.", "1040000001"),
    (True, False): ("LIGHTNING ARRESTER 24-26 KV.5 KA.", "1040000002"),
    (True, True): ("LIGHTNING ARRESTER 24-26 KV.10 KA.", "1040000003"),
}
SURGE_ARRESTER_TANK = {
    False: ("S.A., 21 KV.5 KA.,WITHOUT BRACKET", "1040000007"),
    True: ("S.A., 24 kV.5 kA.,WITHOUT BRACKET.", "1040000008"),
}
SURGE_ARRESTER_CODES = {code for _, code in (*SURGE_ARRESTER_CROSSARM.values(), *SURGE_ARRESTER_TANK.values())}

THREE_PHASE_TRANSFORMER_HEAD = "หม้อแปลง 3 เฟส 22 kV (Sealed)"
THREE_PHASE_TRANSFORMERS = {
    50: ("หม้อแปลงขนาด 50 เควีเอ", "1050010200", 3, "1040030002"),
    100: ("หม้อแปลงขนาด 100 เควีเอ", "1050010201", "5-6", "1040030003"),
    160: ("หม้อแปลงขนาด 160 เควีเอ", "1050010202", 8, "1040030004"),
    250: ("หม้อแปลงขนาด 250 เควีเอ", "1050010204", 15, "1040030006"),
    315: ("หม้อแปลงขนาด 315 เควีเอ", "1050010010", 15, "1040030006"),
    500: ("หม้อแปลงขนาด 500 เควีเอ", "1050010012", 20, "1040030007"),
}
THREE_PHASE_LT_FUSE_CHOICES = {
    (50, 1): [("80",)],
    (100, 1): [("150-160",)],
    (100, 2): [("50", "100"), ("80", "80"), ("80", "100")],
    (160, 1): [("200",)],
    (160, 2): [("80", "150-160"), ("80", "200"), ("100", "150-160")],
    (250, 2): [("100", "200"), ("150-160", "200"), ("200", "200")],
    (315, 2): [("200", "200")],
}
LT_FUSE_SETS = {
    "32-36": ("LT. FUSE, 20 KVA. 3-P, 4-WIRE (32 AMP.)", "Set14019"),
    "50": ("LT. FUSE, 30 KVA. 3-P, 4-WIRE (50 AMP.)", "Set14020"),
    "80": ("LT. FUSE, 50 KVA. 3-P, 4-WIRE (80 AMP.)", "Set14021"),
    "100": ("LT. FUSE, 70 KVA. 3-P, 4-WIRE (100 AMP.)", "Set14022"),
    "150-160": ("LT. FUSE, 100 KVA. 3-P, 4-WIRE (160 AMP.)", "Set14023"),
    "200": ("LT. FUSE, 140 KVA. 3-P, 4-WIRE (200 AMP.)", "Set14024"),
}


def three_phase_config_head(kva: int, circuits: int, choice_index: int) -> str:
    choices = THREE_PHASE_LT_FUSE_CHOICES.get((kva, circuits), [])
    if not 0 <= choice_index < len(choices):
        raise ValueError("ยังไม่มีข้อมูลฟิวส์ที่รองรับสำหรับขนาดและจำนวน LT ที่เลือก")
    fuse_label = "/".join(choices[choice_index])
    return f"3P Sealed 22 kV · {kva} kVA · {circuits} LT · Fuse {fuse_label} A"


def is_three_phase_config_head(head: str) -> bool:
    return bool(re.fullmatch(r"3P Sealed 22 kV · \d+ kVA · [12] LT · Fuse [\d/-]+ A", head.strip()))


def pages_have_work_types(pages: list[list[dict[str, Any]]]) -> bool:
    return any(str(item.get("workType", "")).strip() for page in pages for item in page)


def page_labels(pages: list[list[dict[str, Any]]]) -> list[str]:
    """Return stable Excel labels while keeping legacy untyped exports unchanged."""
    if not pages_have_work_types(pages):
        return [f"หน้า {number}" for number in range(1, len(pages) + 1)]
    counters = {"install": 0, "demolition": 0}
    labels: list[str] = []
    for page in pages:
        work_type = "demolition" if any(item.get("workType") == "demolition" for item in page) else "install"
        counters[work_type] += 1
        prefix = "รื้อถอน" if work_type == "demolition" else "ติดตั้ง"
        labels.append(f"{prefix} {counters[work_type]}")
    return labels


def add_work_type_group_headers(sheet, labels: list[str], start_column: int, row: int = 2) -> None:
    """Add merged installation/demolition headers above per-page columns."""
    groups: list[tuple[str, int, int]] = []
    for offset, label in enumerate(labels):
        if label.startswith("ติดตั้ง "):
            group = "งานติดตั้ง"
        elif label.startswith("รื้อถอน "):
            group = "งานรื้อถอน"
        else:
            continue
        column = start_column + offset
        if groups and groups[-1][0] == group and groups[-1][2] == column - 1:
            groups[-1] = (group, groups[-1][1], column)
        else:
            groups.append((group, column, column))
    colors = {"งานติดตั้ง": "4B216E", "งานรื้อถอน": "B36A00"}
    for group, first_column, last_column in groups:
        if last_column > first_column:
            sheet.merge_cells(start_row=row, start_column=first_column, end_row=row, end_column=last_column)
        cell = sheet.cell(row=row, column=first_column, value=group)
        cell.font = Font(name="Tahoma", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=colors[group])
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for column in range(first_column + 1, last_column + 1):
            sheet.cell(row=row, column=column).fill = PatternFill("solid", fgColor=colors[group])
    if groups:
        sheet.row_dimensions[row].height = 23


class MaterialWorkbook:
    def __init__(self) -> None:
        self.base_df: pd.DataFrame | None = None
        self.set_df: pd.DataFrame | None = None
        self.summary: list[dict[str, Any]] = []
        self.summary_by_work_type: dict[str, list[dict[str, Any]]] = {"install": [], "demolition": []}
        self.transformer_surge_by_work_type: dict[str, dict[str, float]] = {"install": {}, "demolition": {}}
        self.separate_work_types = False
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
        self.summary_by_work_type = {"install": [], "demolition": []}
        self.separate_work_types = False
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
        if department == "แผนกหม้อแปลง":
            transformer_rows = [row for row in rows if row[HEAD_COL].strip().casefold() == "30kva 1p"]
            if transformer_rows:
                editable_rows: list[dict[str, Any]] = []
                for (kva, circuits), fuse_choices in THREE_PHASE_LT_FUSE_CHOICES.items():
                    transformer_name, transformer_code, primary_amp, primary_fuse_code = THREE_PHASE_TRANSFORMERS[kva]
                    for choice_index, fuse_amps in enumerate(fuse_choices):
                        config_head = three_phase_config_head(kva, circuits, choice_index)
                        switch_set = "Set14143" if circuits == 1 else "Set14144"
                        switch_name = "X-ARM-C, INST. WITH 3-LT. SWITCH (3-P, 1-CCT) (12 M)" if circuits == 1 else "X-ARM-C, INST. WITH 6-LT. SWITCH (3-P, 2-CCT) (12 M)"
                        config_rows: list[dict[str, Any]] = []
                        for source in transformer_rows:
                            source_code = clean_text(source[CODE_COL])
                            if source_code == "1050000002":
                                config_rows.append({**source, HEAD_COL: config_head, MATERIAL_COL: transformer_name, CODE_COL: transformer_code, QTY_COL: 1.0})
                            elif source_code == "Set40201":
                                config_rows.append({**source, HEAD_COL: config_head, MATERIAL_COL: "ชุดหม้อแปลง 3 เฟส 22 kV บนคอนเดี่ยว", CODE_COL: "Set40203", QTY_COL: 1.0})
                            elif source_code == "Set14141":
                                config_rows.append({**source, HEAD_COL: config_head, MATERIAL_COL: switch_name, CODE_COL: switch_set, QTY_COL: 1.0})
                            elif source_code == "Set14013":
                                for amp in fuse_amps:
                                    fuse_name, fuse_set = LT_FUSE_SETS[amp]
                                    config_rows.append({**source, HEAD_COL: config_head, MATERIAL_COL: fuse_name, CODE_COL: fuse_set, QTY_COL: 1.0})
                            elif source_code == "1040030000":
                                config_rows.append({**source, HEAD_COL: config_head, MATERIAL_COL: f"FUSE LINK 22 kV. {primary_amp} A.EEI-NEMA TYPE K OR HIGH SURGE TYPE", CODE_COL: primary_fuse_code, QTY_COL: 3.0})
                            elif source_code in {"1040000000", "1040000007"}:
                                source_quantity = parse_number(source[QTY_COL])
                                arrester_quantity = -3.0 if source_quantity < 0 else 3.0 if source_quantity > 0 else 0.0
                                config_rows.append({**source, HEAD_COL: config_head, QTY_COL: arrester_quantity})
                            elif source_code == "Set40104":
                                wiring_set, wire_size = (
                                    ("Set40112", "50") if kva == 50 else ("Set40114", "95")
                                )
                                config_rows.append({
                                    **source,
                                    HEAD_COL: config_head,
                                    MATERIAL_COL: f"LT WIRING, {wire_size} SQ.MM., 3-P, 4-WIRE, TO LT SWITCH",
                                    CODE_COL: wiring_set,
                                    QTY_COL: float(circuits),
                                })
                            else:
                                config_rows.append({**source, HEAD_COL: config_head})
                        editable_rows.extend(config_rows)
                if editable_rows:
                    self.base_df = pd.concat([self.base_df, pd.DataFrame(editable_rows)], ignore_index=True)
        return len(rows)

    def get_departments(self) -> list[str]:
        configured = [DEFAULT_DEPARTMENT, "แผนกแรงสูง TAC", "แผนกแรงต่ำ", "แผนกหม้อแปลง", "แผนกสายส่ง"]
        return configured

    def ensure_high_voltage_auxiliary_catalog(
        self,
        suppressed_heads: set[tuple[str, str, str]] | None = None,
    ) -> None:
        """Copy the high-voltage Surge & Fuse catalog to TAC and ensure GR-1..GR-5."""
        if self.base_df is None or DEPARTMENT_COL not in self.base_df.columns:
            return
        suppressed_heads = suppressed_heads or set()

        def group_mask(department: str, size: str) -> pd.Series:
            return (
                self.base_df[DEPARTMENT_COL].astype(str).str.strip().eq(department)
                & self.base_df[SIZE_COL].astype(str).str.strip().eq(size)
            )

        source_surge = self.base_df.loc[
            group_mask(DEFAULT_DEPARTMENT, SURGE_FUSE_SIZE)
        ].copy()
        target_surge = self.base_df.loc[
            group_mask(HIGH_VOLTAGE_TAC_DEPARTMENT, SURGE_FUSE_SIZE)
        ]
        target_keys = {
            (str(row[HEAD_COL]).strip(), str(row[CODE_COL]).strip())
            for _, row in target_surge.iterrows()
        }
        surge_additions = []
        for _, source in source_surge.iterrows():
            key = (str(source[HEAD_COL]).strip(), str(source[CODE_COL]).strip())
            head = key[0]
            if key in target_keys or (
                HIGH_VOLTAGE_TAC_DEPARTMENT, SURGE_FUSE_SIZE, head
            ) in suppressed_heads:
                continue
            copied = source.copy()
            copied[DEPARTMENT_COL] = HIGH_VOLTAGE_TAC_DEPARTMENT
            surge_additions.append(copied)
            target_keys.add(key)
        if surge_additions:
            self.base_df = pd.concat(
                [self.base_df, pd.DataFrame(surge_additions)], ignore_index=True
            )

        ground_additions = []
        for department in (DEFAULT_DEPARTMENT, HIGH_VOLTAGE_TAC_DEPARTMENT):
            for number, set_code in GROUND_SET_CODES.items():
                short_name = f"HV.GROUNDING GR-{number}"
                canonical_head = f"{short_name} {GROUND_ASSEMBLY}"
                existing = self.base_df.loc[group_mask(department, GROUND_SIZE)]
                heads = existing[HEAD_COL].fillna("").astype(str).str.strip()
                aliases = heads.str.fullmatch(
                    rf"{re.escape(short_name)}(?:\s+\(ASSEMBLY NO\..*\))?",
                    case=False,
                )
                if (
                    (department, GROUND_SIZE, canonical_head) in suppressed_heads
                    or any(
                        (department, GROUND_SIZE, str(alias).strip()) in suppressed_heads
                        for alias in heads[aliases]
                    )
                ):
                    continue
                if canonical_head in set(heads):
                    continue

                if aliases.any():
                    alias_index = existing.index[aliases]
                    self.base_df.loc[alias_index, HEAD_COL] = canonical_head
                    self.base_df.loc[alias_index, MATERIAL_COL] = self.base_df.loc[
                        alias_index, MATERIAL_COL
                    ].where(
                        ~self.base_df.loc[alias_index, CODE_COL]
                        .astype(str).str.strip().str.lower().str.startswith("set"),
                        canonical_head,
                    )
                    self.base_df.loc[alias_index, CODE_COL] = self.base_df.loc[
                        alias_index, CODE_COL
                    ].where(
                        ~self.base_df.loc[alias_index, CODE_COL]
                        .astype(str).str.strip().str.lower().str.startswith("set"),
                        set_code,
                    )
                    continue

                ground_additions.append({
                    SIZE_COL: GROUND_SIZE,
                    HEAD_COL: canonical_head,
                    MATERIAL_COL: canonical_head,
                    CODE_COL: set_code,
                    QTY_COL: 1.0,
                    DEPARTMENT_COL: department,
                    INSULATOR_UPRIGHT_COL: pd.NA,
                    INSULATOR_HORIZONTAL_COL: pd.NA,
                })
        if ground_additions:
            self.base_df = pd.concat(
                [self.base_df, pd.DataFrame(ground_additions)], ignore_index=True
            )

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
        if department == "แผนกหม้อแปลง" and selected == "หม้อแปลง":
            heads.append(THREE_PHASE_TRANSFORMER_HEAD)
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

    def calculate(self, pages: list[list[dict[str, Any]]], structure_pages: list[list[dict[str, Any]]] | None = None) -> dict[str, Any]:
        if self.base_df is None:
            raise ValueError("ยังไม่ได้โหลดไฟล์ BaseData")

        self.separate_work_types = pages_have_work_types(pages)
        totals_by_work_type: dict[str, dict[str, dict[str, Any]]] = {"install": {}, "demolition": {}}
        transformer_surge_by_work_type: dict[str, dict[str, float]] = {"install": {}, "demolition": {}}
        input_count = 0
        matched_rows = 0

        for page_number, page in enumerate(pages, start=1):
            work_type = "demolition" if any(item.get("workType") == "demolition" for item in page) else "install"
            totals = totals_by_work_type[work_type]
            for row_number, item in enumerate(page, start=1):
                size = str(item.get("size", "")).strip()
                head = str(item.get("head", "")).strip()
                department = clean_text(item.get("department")) or DEFAULT_DEPARTMENT
                count = parse_number(item.get("count"))
                if not size or not head or count <= 0:
                    continue

                input_count += 1
                selected_surge_code = ""
                editable_three_phase_variant = department == "แผนกหม้อแปลง" and is_three_phase_config_head(head)
                if department == "แผนกหม้อแปลง" and head == THREE_PHASE_TRANSFORMER_HEAD:
                    head, selected_surge_code = self._three_phase_config(item, page_number, row_number)
                high_voltage = department in {DEFAULT_DEPARTMENT, "แผนกแรงสูง TAC"}
                wire_kind = classify_wire_head(head) if high_voltage else None
                wire1 = clean_text(item.get("wire1"))
                wire2 = clean_text(item.get("wire2"))
                wire3 = clean_text(item.get("wire3"))
                validate_wire_selection(wire_kind, wire1, wire2, page_number, row_number, head, wire3)
                lat_wire = clean_text(item.get("latWire"))
                if has_combined_lat(head):
                    validate_wire_selection("de", lat_wire, "", page_number, row_number, f"{head} — LAT.SLK")
                matches = self.base_df[
                    (self.base_df[SIZE_COL].astype(str).str.strip() == size)
                    & (self.base_df[HEAD_COL].astype(str).str.strip() == head)
                    & self._department_mask(department)
                ]
                if department == "แผนกหม้อแปลง" and selected_surge_code and matches.empty:
                    raise ValueError(f"หน้า {page_number} แถว {row_number}: ไม่พบข้อมูลตั้งต้นของ {head} ใน BaseData")
                has_explicit_surge_adjustments = (
                    department == "แผนกหม้อแปลง"
                    and matches[CODE_COL].astype(str).str.strip().isin(SURGE_ARRESTER_CODES).any()
                )

                for _, row in matches.iterrows():
                    material = clean_text(row[MATERIAL_COL])
                    code = clean_text(row[CODE_COL])
                    amount = parse_number(row[QTY_COL]) * count
                    if not code or amount == 0:
                        continue
                    if department == "แผนกหม้อแปลง" and selected_surge_code and not has_explicit_surge_adjustments and code.casefold() == "set40203":
                        surge_counts = transformer_surge_by_work_type[work_type]
                        surge_counts[selected_surge_code] = surge_counts.get(selected_surge_code, 0.0) + count
                    if department == "แผนกหม้อแปลง" and code in SURGE_ARRESTER_CODES:
                        ngr = bool(item.get("surgeNgr"))
                        within_3km = bool(item.get("surgeWithin3km"))
                        mounting = clean_text(item.get("surgeMounting")) or "crossarm"
                        if within_3km:
                            mounting = "crossarm"  # WITHOUT BRACKET is not available in 10 kA.
                        if editable_three_phase_variant and amount < 0:
                            # The copied negative row offsets the three default arresters inside Set40203.
                            material, code = SURGE_ARRESTER_TANK[False]
                        elif mounting == "tank":
                            if amount < 0:
                                continue
                            material, code = SURGE_ARRESTER_TANK[ngr]
                        else:
                            material, code = SURGE_ARRESTER_TANK[ngr] if amount < 0 else SURGE_ARRESTER_CROSSARM[(ngr, within_3km)]
                    add_material(totals, material, code, amount)
                    matched_rows += 1

                matched_rows += add_wire_materials(totals, wire_kind, wire1, wire2, count * wire_head_multiplier(head), wire3)
                if high_voltage and has_combined_lat(head):
                    matched_rows += add_wire_materials(totals, "de", lat_wire, "", count)

        for page_number, rows in enumerate(structure_pages or [], start=1):
            fallback_type = (
                "demolition" if page_number <= len(pages) and any(item.get("workType") == "demolition" for item in pages[page_number - 1]) else "install"
            )
            for row_number, item in enumerate(rows or [], start=1):
                code = clean_text(item.get("code"))
                material = clean_text(item.get("material"))
                count = parse_number(item.get("count"))
                if not code or not material or count == 0:
                    continue
                work_type = "demolition" if item.get("workType") == "demolition" else fallback_type
                add_material(totals_by_work_type[work_type], material, code, count)
                input_count += 1
                matched_rows += 1

        self.summary_by_work_type = {
            work_type: sorted(totals.values(), key=lambda r: (str(r[CODE_COL]).lower(), str(r[MATERIAL_COL]).lower()))
            for work_type, totals in totals_by_work_type.items()
        }
        self.transformer_surge_by_work_type = transformer_surge_by_work_type
        self.summary = combine_work_type_summaries(self.summary_by_work_type) if self.separate_work_types else self.summary_by_work_type["install"]
        return {
            "items": self.summary,
            "inputRows": input_count,
            "matchedRows": matched_rows,
            "summaryRows": len(self.summary),
        }

    def _three_phase_config(
        self,
        item: dict[str, Any],
        page_number: int,
        row_number: int,
    ) -> tuple[str, str]:
        try:
            kva = int(item.get("transformerKva", 0))
            circuits = int(item.get("transformerLtCircuits", 0))
        except (TypeError, ValueError):
            raise ValueError(f"หน้า {page_number} แถว {row_number}: กรุณาเลือกขนาดหม้อแปลงและจำนวน LT")
        choices = THREE_PHASE_LT_FUSE_CHOICES.get((kva, circuits), [])
        if kva not in THREE_PHASE_TRANSFORMERS or not choices:
            raise ValueError(f"หน้า {page_number} แถว {row_number}: ยังไม่มีข้อมูลฟิวส์ที่รองรับสำหรับ {kva} kVA / {circuits} LT")
        if clean_text(item.get("transformerType", "sealed")) != "sealed":
            raise ValueError(f"หน้า {page_number} แถว {row_number}: รุ่นนี้รองรับเฉพาะหม้อแปลง Sealed")
        if clean_text(item.get("transformerMounting", "single-pole-crossarm")) != "single-pole-crossarm":
            raise ValueError(f"หน้า {page_number} แถว {row_number}: รุ่นนี้รองรับเฉพาะติดตั้งบนคอนเดี่ยว")
        choice = item.get("transformerLtFuseChoice")
        if isinstance(choice, int):
            choice_index = choice
        else:
            try:
                choice_index = int(choice)
            except (TypeError, ValueError):
                choice_index = 0
        if not 0 <= choice_index < len(choices):
            raise ValueError(f"หน้า {page_number} แถว {row_number}: กรุณาเลือกชุดฟิวส์ LT ที่ถูกต้อง")

        return three_phase_config_head(kva, circuits, choice_index), self._selected_transformer_surge_code(item)

    @staticmethod
    def _selected_transformer_surge_code(item: dict[str, Any]) -> str:
        surge_ngr = bool(item.get("surgeNgr"))
        surge_within_3km = bool(item.get("surgeWithin3km"))
        surge_mounting = clean_text(item.get("surgeMounting")) or "crossarm"
        if surge_within_3km:
            surge_mounting = "crossarm"
        if surge_mounting == "tank":
            return SURGE_ARRESTER_TANK[surge_ngr][1]
        return SURGE_ARRESTER_CROSSARM[(surge_ngr, surge_within_3km)][1]

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

        set_found = 0
        set_missing: list[str] = []
        expanded_lines = 0

        set_lookup = self.set_df.copy()
        set_lookup["_set_key"] = set_lookup[SET_COL].astype(str).str.strip().str.lower()

        expanded_by_work_type: dict[str, list[dict[str, Any]]] = {"install": [], "demolition": []}
        for work_type, source_rows in self.summary_by_work_type.items():
            expanded = expanded_by_work_type[work_type]
            for row in source_rows:
                code = clean_text(row[CODE_COL])
                qty = parse_number(row[TOTAL_COL])
                key = code.lower()
                if key.startswith("set"):
                    matches = set_lookup[set_lookup["_set_key"] == key]
                    if matches.empty:
                        set_missing.append(code)
                        continue
                    set_found += 1
                    surge_choices = self.transformer_surge_by_work_type.get(work_type, {}) if key == "set40203" else {}
                    default_surge_code = SURGE_ARRESTER_TANK[False][1]
                    alternate_surge_count = sum(amount for surge_code, amount in surge_choices.items() if surge_code != default_surge_code)
                    for _, item in matches.iterrows():
                        child_code = clean_text(item[CODE_COL])
                        child_quantity = parse_number(item[SET_INSTALL_COL]) * qty
                        if key == "set40203" and child_code == default_surge_code:
                            child_quantity = max(0.0, child_quantity - 3 * alternate_surge_count)
                        if child_quantity:
                            expanded.append(
                                {
                                    MATERIAL_COL: clean_text(item[SET_DESC_COL]),
                                    CODE_COL: child_code,
                                    TOTAL_COL: child_quantity,
                                }
                            )
                            expanded_lines += 1

                    if key == "set40203":
                        surge_materials = {
                            surge_code: material
                            for material, surge_code in (*SURGE_ARRESTER_CROSSARM.values(), *SURGE_ARRESTER_TANK.values())
                        }
                        for surge_code, transformer_count in surge_choices.items():
                            if surge_code == default_surge_code:
                                continue
                            expanded.append(
                                {
                                    MATERIAL_COL: surge_materials[surge_code],
                                    CODE_COL: surge_code,
                                    TOTAL_COL: 3 * transformer_count,
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
            expanded_by_work_type[work_type] = group_summary(expanded)

        self.summary_by_work_type = expanded_by_work_type
        self.summary = combine_work_type_summaries(self.summary_by_work_type) if self.separate_work_types else self.summary_by_work_type["install"]
        return {
            "items": self.summary,
            "setFound": set_found,
            "setMissing": set_missing,
            "expandedLines": expanded_lines,
            "summaryRows": len(self.summary),
        }

    def get_set_components(self, code: str, quantity: Any = 1) -> dict[str, Any]:
        """Return the ten-digit material rows inside one SET for inline preview."""
        if self.set_df is None:
            raise ValueError("ยังไม่ได้โหลดไฟล์ SET")
        set_code = clean_text(code)
        multiplier = parse_number(quantity)
        if not set_code.lower().startswith("set"):
            raise ValueError("รหัสนี้ไม่ใช่ SET")
        set_keys = self.set_df[SET_COL].astype(str).str.strip().str.lower()
        matches = self.set_df[set_keys == set_code.lower()]
        if matches.empty:
            raise ValueError(f"ไม่พบรายละเอียด {set_code}")
        totals: dict[str, dict[str, Any]] = {}
        surge_choices = self.transformer_surge_by_work_type.get("install", {}) if set_code.lower() == "set40203" else {}
        default_surge_code = SURGE_ARRESTER_TANK[False][1]
        alternate_surge_count = sum(amount for surge_code, amount in surge_choices.items() if surge_code != default_surge_code)
        for _, item in matches.iterrows():
            material_code = clean_text(item[CODE_COL])
            if not re.fullmatch(r"\d{10}", material_code):
                continue
            quantity = parse_number(item[SET_INSTALL_COL]) * multiplier
            if set_code.lower() == "set40203" and material_code == default_surge_code:
                quantity = max(0.0, quantity - 3 * alternate_surge_count)
            add_material(
                totals,
                clean_text(item[SET_DESC_COL]),
                material_code,
                quantity,
            )
        if set_code.lower() == "set40203":
            surge_materials = {
                surge_code: material
                for material, surge_code in (*SURGE_ARRESTER_CROSSARM.values(), *SURGE_ARRESTER_TANK.values())
            }
            for surge_code, transformer_count in surge_choices.items():
                if surge_code != default_surge_code:
                    add_material(totals, surge_materials[surge_code], surge_code, 3 * transformer_count)
        items = sorted(totals.values(), key=lambda row: (str(row[CODE_COL]), str(row[MATERIAL_COL])))
        return {
            "setCode": set_code,
            "items": [{"material": row[MATERIAL_COL], "code": row[CODE_COL], "quantity": row[TOTAL_COL]} for row in items],
        }

    def export_summary(self) -> bytes:
        if not self.summary:
            raise ValueError("ยังไม่มีผลลัพธ์สำหรับ export")

        output = BytesIO()
        columns = [MATERIAL_COL, CODE_COL, INSTALL_TOTAL_COL, DEMOLITION_TOTAL_COL] if self.separate_work_types else [MATERIAL_COL, CODE_COL, TOTAL_COL]
        df = pd.DataFrame(self.summary, columns=columns)
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

        page_columns = page_labels(pages)
        typed_pages = pages_have_work_types(pages)
        export_columns: list[str] = []
        display_headers: list[str] = []
        group_labels: list[str] = []
        page_export_keys: dict[int, str] = {}
        total_export_keys: dict[str, str] = {}
        if typed_pages:
            for work_type, prefix in (("install", "ติดตั้ง "), ("demolition", "รื้อถอน ")):
                matching_pages = [
                    (page_number, label)
                    for page_number, label in enumerate(page_columns, start=1)
                    if label.startswith(prefix)
                ]
                if not matching_pages:
                    continue
                for page_number, label in matching_pages:
                    key = f"__{work_type}_page_{page_number}"
                    export_columns.append(key)
                    display_headers.append(label.removeprefix(prefix))
                    group_labels.append(label)
                    page_export_keys[page_number] = key
                total_key = f"__{work_type}_total"
                export_columns.append(total_key)
                display_headers.append("รวม")
                group_labels.append(f"{prefix}รวม")
                total_export_keys[work_type] = total_key
        else:
            export_columns = page_columns.copy()
            display_headers = page_columns.copy()
        rows = []
        for (head, size), page_totals in sorted(totals.items(), key=lambda item: (natural_key(item[0][1]), natural_key(item[0][0]))):
            row: dict[str, Any] = {HEAD_COL: head, "เสา": size}
            if typed_pages:
                for page_number, key in page_export_keys.items():
                    row[key] = page_totals.get(page_number)
                for work_type, total_key in total_export_keys.items():
                    prefix = "รื้อถอน " if work_type == "demolition" else "ติดตั้ง "
                    row[total_key] = sum(
                        page_totals.get(page_number, 0.0)
                        for page_number, label in enumerate(page_columns, start=1)
                        if label.startswith(prefix)
                    )
            else:
                for page_number, column in enumerate(page_columns, start=1):
                    row[column] = page_totals.get(page_number)
            rows.append(row)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            columns = [HEAD_COL, "เสา", *export_columns]
            pd.DataFrame(rows, columns=columns).to_excel(writer, index=False, sheet_name="สรุปแต่ละหน้า", startrow=2)
            sheet = writer.sheets["สรุปแต่ละหน้า"]
            if typed_pages:
                for column_index, header in enumerate(display_headers, start=3):
                    sheet.cell(row=3, column=column_index, value=header)
            last_column = get_column_letter(len(columns))
            sheet.merge_cells(f"A1:{last_column}1")
            title = sheet["A1"]
            title.value = "สรุปรายการหัวเสาแยกตามหน้า"
            title.font = Font(name="Tahoma", size=16, bold=True, color="FFFFFF")
            title.fill = PatternFill("solid", fgColor="0F766E")
            title.alignment = Alignment(horizontal="center", vertical="center")
            sheet.row_dimensions[1].height = 28
            add_work_type_group_headers(sheet, group_labels if typed_pages else page_columns, start_column=3, row=2)

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

        labels = page_labels(pages)
        typed_pages = pages_have_work_types(pages)
        for page_number, page in enumerate(pages, start=1):
            page_label: Any = labels[page_number - 1] if typed_pages else page_number
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
                    details.append({"หน้า": page_label, HEAD_COL: head, "ที่มา": "เกณฑ์ลูกถ้วย", MATERIAL_COL: "ลูกถ้วยตั้ง", CODE_COL: "", TOTAL_COL: upright * count})
                if horizontal:
                    details.append({"หน้า": page_label, HEAD_COL: head, "ที่มา": "เกณฑ์ลูกถ้วย", MATERIAL_COL: "ลูกถ้วยนอน", CODE_COL: "", TOTAL_COL: horizontal * count})

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
                            details.append({"หน้า": page_label, HEAD_COL: head, "ที่มา": set_code, MATERIAL_COL: material, CODE_COL: code, TOTAL_COL: amount})

                high_voltage = department in {DEFAULT_DEPARTMENT, "แผนกแรงสูง TAC"}
                wire_kind = classify_wire_head(head) if high_voltage else None
                wire1 = clean_text(item.get("wire1"))
                wire2 = clean_text(item.get("wire2"))
                wire3 = clean_text(item.get("wire3"))
                validate_wire_selection(wire_kind, wire1, wire2, page_number, row_number, head, wire3)
                wire_totals: dict[str, dict[str, Any]] = {}
                add_wire_materials(wire_totals, wire_kind, wire1, wire2, count * wire_head_multiplier(head), wire3)

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
                        "หน้า": page_label, HEAD_COL: head, "ที่มา": "ช่องเลือกสาย",
                        MATERIAL_COL: clean_text(material[MATERIAL_COL]), CODE_COL: clean_text(material[CODE_COL]), TOTAL_COL: amount,
                    })

        if not totals:
            raise ValueError("ยังไม่มีข้อมูลลูกถ้วยหรืออุปกรณ์ยึดสายสำหรับ export")

        page_columns = labels
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
            add_work_type_group_headers(sheet, page_columns, start_column=4, row=2)
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
        page_totals: dict[str, tuple[float, float]] = {}
        details: list[dict[str, Any]] = []

        labels = page_labels(pages)
        typed_pages = pages_have_work_types(pages)
        for page_number, page in enumerate(pages, start=1):
            page_label: Any = labels[page_number - 1] if typed_pages else page_number
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
                    "หน้า": page_label,
                    SIZE_COL: size,
                    HEAD_COL: head,
                    "จำนวนหัว": count,
                    "ตั้ง/หัว": upright_rate,
                    "นอน/หัว": horizontal_rate,
                    "ลูกถ้วยตั้ง": upright,
                    "ลูกถ้วยนอน": horizontal,
                })
            page_totals[page_label] = (upright_total, horizontal_total)

        if not details:
            raise ValueError("ยังไม่มีข้อมูลหัวเสาที่มีลูกถ้วยสำหรับ export")

        summary_rows = []
        for page_label, (upright, horizontal) in page_totals.items():
            summary_rows.append({"หน้า": page_label, "ลูกถ้วยตั้ง": upright, "ลูกถ้วยนอน": horizontal, "รวมลูกถ้วย": upright + horizontal})
        if typed_pages:
            for prefix, total_label in (("ติดตั้ง ", "รวมงานติดตั้ง"), ("รื้อถอน ", "รวมงานรื้อถอน")):
                selected = [value for label, value in page_totals.items() if str(label).startswith(prefix)]
                summary_rows.append({
                    "หน้า": total_label,
                    "ลูกถ้วยตั้ง": sum(value[0] for value in selected),
                    "ลูกถ้วยนอน": sum(value[1] for value in selected),
                    "รวมลูกถ้วย": sum(sum(value) for value in selected),
                })
        else:
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

        def add_detail(page_number: int, page_label: str, size: str, head: str, head_count: float, material: str, code: str, amount: float) -> None:
            is_crossarm = code.startswith("10001") or bool(re.match(r"^STEEL\s*CHANNEL", material, re.IGNORECASE))
            if not is_crossarm or amount == 0:
                return
            key = (page_number, size, head, code)
            if key not in detail_totals:
                detail_totals[key] = {
                    "หน้า": page_label, "_page_number": page_number, SIZE_COL: size, HEAD_COL: head, "จำนวนหัว": head_count,
                    MATERIAL_COL: material, CODE_COL: code, TOTAL_COL: 0.0,
                }
            detail_totals[key][TOTAL_COL] += amount

        labels = page_labels(pages)
        typed_pages = pages_have_work_types(pages)
        for page_number, page in enumerate(pages, start=1):
            page_label: Any = labels[page_number - 1] if typed_pages else page_number
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
                                page_number, page_label, size, head, count,
                                clean_text(set_row[SET_DESC_COL]), code,
                                base_quantity * parse_number(set_row[SET_INSTALL_COL]),
                            )
                    else:
                        add_detail(
                            page_number, page_label, size, head, count,
                            clean_text(base_row[MATERIAL_COL]), base_code, base_quantity,
                        )

        details = [row for row in detail_totals.values() if row[TOTAL_COL] != 0]
        if not details:
            raise ValueError("ไม่พบรายการคอนจากหัวเสาที่เลือก")

        summary_totals: dict[str, dict[str, Any]] = {}
        for row in details:
            code = row[CODE_COL]
            page_number = int(row["_page_number"])
            summary = summary_totals.setdefault(code, {MATERIAL_COL: row[MATERIAL_COL], "pages": {}})
            summary["pages"][page_number] = summary["pages"].get(page_number, 0.0) + parse_number(row[TOTAL_COL])
        summary_rows = []
        for code, summary in sorted(summary_totals.items()):
            amounts = summary["pages"]
            row: dict[str, Any] = {MATERIAL_COL: summary[MATERIAL_COL], CODE_COL: code}
            for page_number, page_label in enumerate(labels, start=1):
                row[page_label] = amounts.get(page_number, 0.0)
            if typed_pages:
                row[INSTALL_TOTAL_COL] = sum(amounts.get(number, 0.0) for number, label in enumerate(labels, start=1) if label.startswith("ติดตั้ง "))
                row[DEMOLITION_TOTAL_COL] = sum(amounts.get(number, 0.0) for number, label in enumerate(labels, start=1) if label.startswith("รื้อถอน "))
            else:
                row[TOTAL_COL] = sum(amounts.values())
            summary_rows.append(row)

        detail_columns = ["หน้า", SIZE_COL, HEAD_COL, "จำนวนหัว", MATERIAL_COL, CODE_COL, TOTAL_COL]
        summary_columns = [
            MATERIAL_COL, CODE_COL, *labels,
            *([INSTALL_TOTAL_COL, DEMOLITION_TOTAL_COL] if typed_pages else [TOTAL_COL]),
        ]
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary_startrow = 1 if typed_pages else 0
            pd.DataFrame(summary_rows, columns=summary_columns).to_excel(
                writer, index=False, sheet_name="สรุปคอนแยกหน้า", startrow=summary_startrow
            )
            pd.DataFrame(details, columns=detail_columns).sort_values(["หน้า", HEAD_COL, CODE_COL]).to_excel(
                writer, index=False, sheet_name="ที่มาคอน"
            )
            for sheet_name in ("สรุปคอนแยกหน้า", "ที่มาคอน"):
                sheet = writer.sheets[sheet_name]
                header_row = 2 if sheet_name == "สรุปคอนแยกหน้า" and typed_pages else 1
                for cell in sheet[header_row]:
                    cell.font = Font(name="Tahoma", bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="4B216E")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for row in sheet.iter_rows(min_row=header_row + 1, max_row=sheet.max_row):
                    for cell in row:
                        cell.font = Font(name="Tahoma", size=10)
                    for cell in row[2:]:
                        cell.number_format = "#,##0.###"
                sheet.freeze_panes = f"C{header_row + 1}" if sheet_name == "สรุปคอนแยกหน้า" else "A2"
                sheet.auto_filter.ref = f"A{header_row}:{get_column_letter(sheet.max_column)}{sheet.max_row}"
                sheet.sheet_view.showGridLines = False
            summary_sheet = writer.sheets["สรุปคอนแยกหน้า"]
            if typed_pages:
                add_work_type_group_headers(summary_sheet, labels, start_column=3, row=1)
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
        "DDE,DE(ST.4.5M)": "dde_de",
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
        "DDE,DE(ST.4.5M)": (6, 36),
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
    wire3: str = "",
) -> None:
    if wire_kind is None:
        return
    if wire1 not in WIRE_MATERIALS:
        raise ValueError(f"หน้า {page_number} แถว {row_number} ({head}): กรุณาเลือกชนิดสายช่องแรก")
    if wire_kind in {"dde", "dde_bl", "ba"} and wire2 not in WIRE_MATERIALS:
        raise ValueError(f"หน้า {page_number} แถว {row_number} ({head}): กรุณาเลือกชนิดสายช่องที่สอง")
    if wire_kind == "dde_de" and (wire2 not in WIRE_MATERIALS or wire3 not in WIRE_MATERIALS):
        raise ValueError(f"หน้า {page_number} แถว {row_number} ({head}): กรุณาเลือกชนิดสาย DDE สองด้านและสาย DE ให้ครบ")
    if wire_kind in {"dde", "dde_de"} and conductor_group(wire1) != conductor_group(wire2):
        raise ValueError(f"หน้า {page_number} แถว {row_number} ({head}): สายซ้ายและขวาต้องมีขนาดเดียวกันสำหรับ Tensionless")
    if wire_kind in {"dde", "dde_de"} and not is_dde_sac_acsr_transition(wire1, wire2) and wire1 not in TENSIONLESS_MATERIALS:
        raise ValueError(f"หน้า {page_number} แถว {row_number} ({head}): ยังไม่มีรหัสหลอดต่อสำหรับสาย {wire1}")


def conductor_group(wire: str) -> str:
    details = WIRE_MATERIALS.get(clean_text(wire))
    return details[2] if details else ""


def is_dde_sac_acsr_transition(wire1: str, wire2: str) -> bool:
    """DDE transition between 185 SAC and 185 ACSR uses PG3 instead of tensionless."""
    return {clean_text(wire1), clean_text(wire2)} == {"185 SAC", "185 ACSR"}


def add_material(
    totals: dict[str, dict[str, Any]],
    material: str,
    code: str,
    amount: float,
) -> None:
    if not code or amount == 0:
        return
    # Material descriptions may differ only by spacing or minor typing variations.
    # The material code is the identity; keep the first non-empty description found.
    key = code.casefold()
    if key not in totals:
        totals[key] = {MATERIAL_COL: material, CODE_COL: code, TOTAL_COL: 0.0}
    elif not totals[key][MATERIAL_COL] and material:
        totals[key][MATERIAL_COL] = material
    totals[key][TOTAL_COL] += amount


def add_wire_materials(
    totals: dict[str, dict[str, Any]],
    wire_kind: str | None,
    wire1: str,
    wire2: str,
    count: float,
    wire3: str = "",
) -> int:
    if wire_kind is None:
        return 0

    selected_wires = [wire2] if wire_kind == "ba" else [wire1]
    if wire_kind in {"dde", "dde_bl"}:
        selected_wires.append(wire2)
    if wire_kind == "dde_de":
        selected_wires.extend([wire2, wire3])

    added = 0
    for wire in selected_wires:
        material, code, _, needs_clevis = WIRE_MATERIALS[wire]
        add_material(totals, material, code, 3 * count)
        added += 1
        if needs_clevis:
            clevis_material, clevis_code = CLEVIS_MATERIAL
            add_material(totals, clevis_material, clevis_code, 3 * count)
            added += 1

    if wire_kind in {"dde", "dde_de"}:
        if is_dde_sac_acsr_transition(wire1, wire2):
            material, code = PG3_MATERIAL
            add_material(totals, material, code, 6 * count)
            added += 1
        else:
            material, code = TENSIONLESS_MATERIALS[wire1]
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
    totals: dict[str, dict[str, Any]] = {}
    for row in rows:
        material = clean_text(row[MATERIAL_COL])
        code = clean_text(row[CODE_COL])
        amount = parse_number(row[TOTAL_COL])
        if not code or amount == 0:
            continue
        key = code.casefold()
        if key not in totals:
            totals[key] = {MATERIAL_COL: material, CODE_COL: code, TOTAL_COL: 0.0}
        elif not totals[key][MATERIAL_COL] and material:
            totals[key][MATERIAL_COL] = material
        totals[key][TOTAL_COL] += amount
    return sorted(totals.values(), key=lambda r: (str(r[CODE_COL]).lower(), str(r[MATERIAL_COL]).lower()))


def combine_work_type_summaries(summary_by_work_type: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    combined: dict[str, dict[str, Any]] = {}
    for work_type, rows in summary_by_work_type.items():
        quantity_column = DEMOLITION_TOTAL_COL if work_type == "demolition" else INSTALL_TOTAL_COL
        for row in rows:
            material = clean_text(row[MATERIAL_COL])
            code = clean_text(row[CODE_COL])
            key = code.casefold()
            target = combined.setdefault(key, {
                MATERIAL_COL: material, CODE_COL: code,
                INSTALL_TOTAL_COL: 0.0, DEMOLITION_TOTAL_COL: 0.0, TOTAL_COL: 0.0,
            })
            if not target[MATERIAL_COL] and material:
                target[MATERIAL_COL] = material
            amount = parse_number(row[TOTAL_COL])
            target[quantity_column] += amount
            target[TOTAL_COL] += amount  # Compatibility for older callers; UI/export use the separated columns.
    return sorted(combined.values(), key=lambda row: (str(row[CODE_COL]).lower(), str(row[MATERIAL_COL]).lower()))


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
        except (SyntaxError, TypeError, ValueError, ArithmeticError):
            return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def evaluate_add_sub(node: ast.AST) -> float:
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
        return float(node.value)
    if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)):
        left = evaluate_add_sub(node.left)
        right = evaluate_add_sub(node.right)
        if isinstance(node.op, ast.Add):
            return left + right
        if isinstance(node.op, ast.Sub):
            return left - right
        if isinstance(node.op, ast.Mult):
            return left * right
        return left / right
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
        value = evaluate_add_sub(node.operand)
        return value if isinstance(node.op, ast.UAdd) else -value
    raise ValueError("รองรับเฉพาะตัวเลข เครื่องหมาย +, -, *, / และวงเล็บ")


def natural_key(value: str) -> list[tuple[int, Any]]:
    # A punctuation dot in "st.4.5m" belongs to the text, not to 4.5.
    return [
        (0, float(part)) if index % 2 else (1, part.lower())
        for index, part in enumerate(re.split(r"([0-9]+(?:\.[0-9]+)?)", str(value)))
        if part
    ]
