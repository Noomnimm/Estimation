import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTableWidget,
    QTableWidgetItem, QLabel, QComboBox, QLineEdit, QGridLayout, QHBoxLayout
)
from functools import partial


class MaterialCalculator(QWidget):
    def __init__(self):
        super().__init__()

        self.df = None
        self.df_set = None

        # ใช้ key แบบ (material, code) กันชื่อซ้ำ
        self.material_summary = {}

        self.input_rows = 2
        self.unique_sizes = []

        self.total_pages = 1
        self.current_page = 1
        self.page_data = {i: [] for i in range(1, self.total_pages + 1)}

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Material Calculator")
        self.setGeometry(100, 100, 1100, 650)
        self.layout = QVBoxLayout()

        # ---------- File Selection ----------
        file_layout = QHBoxLayout()
        self.file_label = QLabel("Please select an Excel file (BaseData)")
        file_layout.addWidget(self.file_label)

        self.loadButton = QPushButton("Load Excel File (BaseData)")
        self.loadButton.clicked.connect(self.load_excel)
        file_layout.addWidget(self.loadButton)
        self.layout.addLayout(file_layout)

        # ---------- Row Control ----------
        row_control_layout = QHBoxLayout()
        row_control_layout.addWidget(QLabel("จำนวนช่องกรอกข้อมูล:"))

        self.add_row_btn = QPushButton("➕ เพิ่มแถว")
        self.add_row_btn.clicked.connect(self.add_row)
        row_control_layout.addWidget(self.add_row_btn)

        self.remove_row_btn = QPushButton("➖ ลบแถว")
        self.remove_row_btn.clicked.connect(self.remove_row)
        row_control_layout.addWidget(self.remove_row_btn)

        self.layout.addLayout(row_control_layout)

        # ---------- Page Control ----------
        page_control_layout = QHBoxLayout()
        page_control_layout.addWidget(QLabel("จำนวนหน้าทั้งหมด:"))

        self.page_input = QLineEdit(str(self.total_pages))
        self.page_input.setFixedWidth(60)
        self.page_input.setPlaceholderText("หน้า")
        page_control_layout.addWidget(self.page_input)

        set_page_btn = QPushButton("กำหนดจำนวนหน้า")
        set_page_btn.clicked.connect(self.set_total_pages)
        page_control_layout.addWidget(set_page_btn)

        self.layout.addLayout(page_control_layout)

        # ---------- Navigation ----------
        nav_layout = QHBoxLayout()
        self.prevButton = QPushButton("◀ Previous Page")
        self.prevButton.clicked.connect(self.prev_page)
        nav_layout.addWidget(self.prevButton)

        self.page_label = QLabel(f"หน้า {self.current_page}/{self.total_pages}")
        nav_layout.addWidget(self.page_label)

        self.nextButton = QPushButton("Next Page ▶")
        self.nextButton.clicked.connect(self.next_page)
        nav_layout.addWidget(self.nextButton)

        self.layout.addLayout(nav_layout)

        # ---------- Clear Button ----------
        self.clearButton = QPushButton("ล้างข้อมูล (เฉพาะหน้านี้)")
        self.clearButton.clicked.connect(self.clear_current_page_data)
        self.layout.addWidget(self.clearButton)

        # ---------- Input Grid ----------
        self.input_grid = QGridLayout()
        self.inputs = []
        self.create_input_grid()
        self.layout.addLayout(self.input_grid)

        # ---------- Action Buttons ----------
        btn_layout = QHBoxLayout()

        self.calculateButton = QPushButton("คำนวณ")
        self.calculateButton.clicked.connect(self.calculate_materials)
        btn_layout.addWidget(self.calculateButton)

        self.exportButton = QPushButton("บันทึกไฟล์ Excel")
        self.exportButton.clicked.connect(self.export_to_excel)
        btn_layout.addWidget(self.exportButton)

        self.loadSetButton = QPushButton("แนบไฟล์ชุด SET")
        self.loadSetButton.clicked.connect(self.load_set_file)
        btn_layout.addWidget(self.loadSetButton)

        self.expandSetButton = QPushButton("แปลง SET เป็น 10 หลัก")
        self.expandSetButton.clicked.connect(self.expand_set_items)
        btn_layout.addWidget(self.expandSetButton)

        self.layout.addLayout(btn_layout)

        # ---------- Log Label ----------
        self.log_label = QLabel("Log: -")
        self.layout.addWidget(self.log_label)

        # ---------- Table ----------
        self.tableWidget = QTableWidget()
        self.layout.addWidget(self.tableWidget)

        self.setLayout(self.layout)

    # -------------------- Grid -------------------- #
    def create_input_grid(self):
        self.clear_grid()
        self.inputs = []

        for row in range(self.input_rows):
            size_input = QComboBox()
            size_input.setMinimumWidth(150)
            if self.df is not None:
                size_input.addItems(["  "] + self.unique_sizes)
            size_input.currentIndexChanged.connect(partial(self.update_head_options, row))
            self.input_grid.addWidget(size_input, row, 0)

            head_input = QComboBox()
            head_input.setMinimumWidth(250)
            self.input_grid.addWidget(head_input, row, 1)

            count_input = QLineEdit()
            count_input.setPlaceholderText("กรอกจำนวน")
            count_input.setMinimumWidth(120)
            self.input_grid.addWidget(count_input, row, 2)

            self.inputs.append((size_input, head_input, count_input))

        # ✅ Restore data correctly (populate head options first)
        if self.page_data.get(self.current_page):
            for i, (size, head, count) in enumerate(self.page_data[self.current_page]):
                if i >= len(self.inputs):
                    break
                size_cb, head_cb, count_le = self.inputs[i]
                size_cb.setCurrentText(size)
                self.update_head_options(i)      # สำคัญ
                head_cb.setCurrentText(head)
                count_le.setText(str(count))

    def clear_grid(self):
        while self.input_grid.count():
            item = self.input_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    # -------------------- Row Control -------------------- #
    def add_row(self):
        self.save_current_page_data()
        self.input_rows += 1
        self.create_input_grid()

    def remove_row(self):
        if self.input_rows > 1:
            self.save_current_page_data()
            self.input_rows -= 1
            self.create_input_grid()

    # -------------------- Page Control -------------------- #
    def set_total_pages(self):
        try:
            value = int(self.page_input.text())
            if value < 1:
                raise ValueError

            self.save_current_page_data()
            old_data = self.page_data.copy()

            self.total_pages = value
            self.page_data = {i: old_data.get(i, []) for i in range(1, self.total_pages + 1)}

            if self.current_page > self.total_pages:
                self.current_page = self.total_pages

            self.page_label.setText(f"หน้า {self.current_page}/{self.total_pages}")
            self.create_input_grid()
            self.file_label.setText("✅ กำหนดจำนวนหน้าเรียบร้อย")

        except Exception:
            self.file_label.setText("❌ ใส่จำนวนหน้าไม่ถูกต้อง (ต้องเป็นเลขจำนวนเต็มบวก)")

    def prev_page(self):
        if self.current_page > 1:
            self.save_current_page_data()
            self.current_page -= 1
            self.page_label.setText(f"หน้า {self.current_page}/{self.total_pages}")
            self.create_input_grid()

    def next_page(self):
        if self.current_page < self.total_pages:
            self.save_current_page_data()
            self.current_page += 1
            self.page_label.setText(f"หน้า {self.current_page}/{self.total_pages}")
            self.create_input_grid()

    def save_current_page_data(self):
        self.page_data[self.current_page] = [
            (size.currentText(), head.currentText(), count.text())
            for size, head, count in self.inputs
        ]

    def clear_current_page_data(self):
        for size_input, head_input, count_input in self.inputs:
            size_input.setCurrentIndex(0)
            head_input.clear()
            count_input.clear()
        self.page_data[self.current_page] = []
        self.file_label.setText("✅ ล้างข้อมูลหน้านี้แล้ว")

    # -------------------- File Load -------------------- #
    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)")
        if path:
            try:
                self.df = pd.read_excel(path, sheet_name='BaseData')
                self.unique_sizes = sorted(self.df['ขนาดเสา (m)'].astype(str).unique())
                self.create_input_grid()
                self.file_label.setText(f"✅ โหลด BaseData สำเร็จ: {path}")
            except Exception as e:
                self.df = None
                self.file_label.setText(f"❌ โหลด BaseData ไม่สำเร็จ: {e}")

    def load_set_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "แนบไฟล์ชุดวัสดุ", "", "Excel Files (*.xlsx)")
        if path:
            try:
                self.df_set = pd.read_excel(path, sheet_name='วัสดุทั้งหมด')
                self.file_label.setText(f"✅ โหลดฐานข้อมูลชุด SET แล้ว: {path}")
            except Exception as e:
                self.df_set = None
                self.file_label.setText(f"❌ โหลดไฟล์ SET ไม่สำเร็จ: {e}")

    # -------------------- Dropdown Update -------------------- #
    def update_head_options(self, row):
        if self.df is None or row >= len(self.inputs):
            return

        selected_size = self.inputs[row][0].currentText()
        head_input = self.inputs[row][1]

        if selected_size.strip():
            heads = self.df[
                self.df['ขนาดเสา (m)'].astype(str) == str(selected_size)
            ]['รหัสหัวเสา'].astype(str).unique()

            head_input.clear()
            head_input.addItems(["  "] + list(heads))
        else:
            head_input.clear()

    # -------------------- Calculate -------------------- #
    def calculate_materials(self):
        self.save_current_page_data()

        if self.df is None:
            self.file_label.setText("❌ กรุณาโหลดไฟล์ BaseData ก่อน")
            return

        self.material_summary = {}
        total_lines = 0

        for page in range(1, self.total_pages + 1):
            for size, head, count in self.page_data.get(page, []):
                if not str(size).strip() or not str(head).strip() or not str(count).isdigit():
                    continue

                count = int(count)

                # ✅ FIX: เทียบแบบ string กัน type mismatch
                materials = self.df[
                    (self.df['ขนาดเสา (m)'].astype(str) == str(size)) &
                    (self.df['รหัสหัวเสา'].astype(str) == str(head))
                ]

                for _, r in materials.iterrows():
                    material = str(r['รายการวัสดุ'])
                    code = str(r['รหัสพัสดุ'])
                    amount = float(r['จำนวน']) * count

                    key = (material, code)
                    if key in self.material_summary:
                        self.material_summary[key]['จำนวนรวม'] += amount
                    else:
                        self.material_summary[key] = {
                            'รายการวัสดุ': material,
                            'รหัสพัสดุ': code,
                            'จำนวนรวม': amount
                        }
                    total_lines += 1

        self.display_table()
        self.log_label.setText(f"Log: คำนวณเสร็จ | พบรายการรวม {len(self.material_summary)} แถว (จากการรวม)")

    # -------------------- Expand SET -------------------- #
    def expand_set_items(self):
        # ✅ FIX: DataFrame ห้ามใช้เป็น bool
        if self.df_set is None or not self.material_summary:
            self.file_label.setText("กรุณาโหลดไฟล์ SET และคำนวณก่อน")
            return

        expanded = []
        set_found = 0
        set_missing = []
        expanded_lines = 0

        try:
            for _, data in self.material_summary.items():
                code = str(data['รหัสพัสดุ']).strip().lower()
                qty = float(data['จำนวนรวม'])

                if code.startswith("set"):
                    # ✅ FIX: normalize + กัน NaN
                    matches = self.df_set[
                        self.df_set['Set'].astype(str).str.strip().str.lower() == code
                    ]

                    if matches.empty:
                        set_missing.append(code)
                        continue

                    set_found += 1

                    for _, item in matches.iterrows():
                        amount = float(item['ติดตั้ง']) * qty
                        expanded.append({
                            'รายการวัสดุ': str(item['คำอธิบาย']),
                            'รหัสพัสดุ': str(item['รหัสพัสดุ']),
                            'จำนวนรวม': amount
                        })
                        expanded_lines += 1
                else:
                    expanded.append({
                        'รายการวัสดุ': str(data['รายการวัสดุ']),
                        'รหัสพัสดุ': str(data['รหัสพัสดุ']),
                        'จำนวนรวม': float(data['จำนวนรวม'])
                    })

            df = pd.DataFrame(expanded)
            df_summary = df.groupby(['รายการวัสดุ', 'รหัสพัสดุ'], as_index=False)['จำนวนรวม'].sum()

            # rebuild summary
            self.material_summary = {
                (str(row['รายการวัสดุ']), str(row['รหัสพัสดุ'])): {
                    'รายการวัสดุ': str(row['รายการวัสดุ']),
                    'รหัสพัสดุ': str(row['รหัสพัสดุ']),
                    'จำนวนรวม': float(row['จำนวนรวม'])
                }
                for _, row in df_summary.iterrows()
            }

            self.display_table()

            # Log
            msg = f"Log: แตก SET เสร็จ | พบ SET {set_found} รายการ | แตกเป็นวัสดุ {expanded_lines} แถว | ผลรวมใหม่ {len(self.material_summary)} แถว"
            if set_missing:
                msg += f" | ❗ไม่พบ SET: {', '.join(set_missing[:10])}" + (" ..." if len(set_missing) > 10 else "")
            self.log_label.setText(msg)
            self.file_label.setText("✅ แปลง SET เป็น 10 หลัก เรียบร้อย")

        except Exception as err:
            self.file_label.setText("❌ เกิดข้อผิดพลาดขณะประมวลผล SET")
            self.log_label.setText(f"Log: ERROR expand_set_items -> {err}")

    # -------------------- Table Display -------------------- #
    def display_table(self):
        self.tableWidget.clear()
        self.tableWidget.setColumnCount(3)

        items = list(self.material_summary.values())
        self.tableWidget.setRowCount(len(items))
        self.tableWidget.setHorizontalHeaderLabels(["รายการวัสดุ", "รหัสพัสดุ", "จำนวนรวม"])

        for row, data in enumerate(items):
            self.tableWidget.setItem(row, 0, QTableWidgetItem(str(data['รายการวัสดุ'])))
            self.tableWidget.setItem(row, 1, QTableWidgetItem(str(data['รหัสพัสดุ'])))
            self.tableWidget.setItem(row, 2, QTableWidgetItem(str(data['จำนวนรวม'])))

        self.tableWidget.resizeColumnsToContents()

    # -------------------- Export -------------------- #
    def export_to_excel(self):
        if not self.material_summary:
            self.file_label.setText("ยังไม่มีข้อมูลวัสดุ กรุณาคำนวณก่อน!")
            return

        path, _ = QFileDialog.getSaveFileName(self, "บันทึกไฟล์ Excel", "material_summary.xlsx", "Excel Files (*.xlsx)")
        if path:
            try:
                df_export = pd.DataFrame(list(self.material_summary.values()))
                df_export.to_excel(path, index=False)
                self.file_label.setText(f"✅ บันทึกไฟล์แล้ว: {path}")
            except Exception as e:
                self.file_label.setText(f"❌ บันทึกไฟล์ไม่สำเร็จ: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MaterialCalculator()
    window.show()
    sys.exit(app.exec_())