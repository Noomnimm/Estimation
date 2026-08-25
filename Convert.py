import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel,
    QTableWidget, QTableWidgetItem
)

class SetExpander(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SET Expander")
        self.setGeometry(300, 300, 900, 600)

        self.df_input = None
        self.df_set = None
        self.df_result = None

        self.layout = QVBoxLayout()

        self.label_status = QLabel("📂 กรุณาแนบไฟล์")
        self.layout.addWidget(self.label_status)

        self.btn_input = QPushButton("แนบไฟล์วัสดุ (สรุป)")
        self.btn_input.clicked.connect(self.load_input_file)
        self.layout.addWidget(self.btn_input)

        self.btn_set = QPushButton("แนบไฟล์ฐานชุดวัสดุ (sheet: วัสดุทั้งหมด)")
        self.btn_set.clicked.connect(self.load_set_file)
        self.layout.addWidget(self.btn_set)

        self.btn_expand = QPushButton("แปลง SET → รายการย่อย")
        self.btn_expand.clicked.connect(self.expand_set_items)
        self.layout.addWidget(self.btn_expand)

        self.btn_export = QPushButton("บันทึกผลลัพธ์เป็น Excel")
        self.btn_export.clicked.connect(self.export_result)
        self.layout.addWidget(self.btn_export)

        self.table = QTableWidget()
        self.layout.addWidget(self.table)

        self.setLayout(self.layout)

    def load_input_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "เลือกไฟล์วัสดุ", "", "Excel Files (*.xlsx)")
        if path:
            self.df_input = pd.read_excel(path)
            self.label_status.setText(f"✅ โหลดวัสดุ: {path}")

    def load_set_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "เลือกไฟล์ชุดวัสดุ", "", "Excel Files (*.xlsx)")
        if path:
            self.df_set = pd.read_excel(path, sheet_name='วัสดุทั้งหมด')
            self.label_status.setText(f"✅ โหลดชุดวัสดุ: {path}")

    def expand_set_items(self):
        if self.df_input is None or self.df_set is None:
            self.label_status.setText("❌ กรุณาแนบไฟล์ทั้งสองก่อน")
            return

        expanded = []
        for _, row in self.df_input.iterrows():
            code = str(row['รหัสพัสดุ']).strip().lower()
            qty = row['จำนวนรวม'] if pd.notna(row['จำนวนรวม']) else 0

            if code.startswith("set"):
                matches = self.df_set[self.df_set['Set'].str.lower() == code]
                if matches.empty:
                    continue
                for _, item in matches.iterrows():
                    try:
                        item_qty = float(item['ติดตั้ง']) * qty
                        expanded.append({
                            'รายการวัสดุ': str(item['คำอธิบาย']).strip(),
                            'รหัสพัสดุ': str(item['รหัสพัสดุ']).strip(),
                            'จำนวนรวม': item_qty
                        })
                    except:
                        continue
            else:
                expanded.append({
                    'รายการวัสดุ': str(row['รายการวัสดุ']).strip(),
                    'รหัสพัสดุ': str(row['รหัสพัสดุ']).strip(),
                    'จำนวนรวม': qty
                })

        df = pd.DataFrame(expanded)

        # 🔧 Clean รายการวัสดุให้เรียบร้อยก่อน group
        df['รายการวัสดุ'] = df['รายการวัสดุ'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        df['รหัสพัสดุ'] = df['รหัสพัสดุ'].astype(str).str.strip()

        # ✅ Group by รหัสพัสดุ และเอาชื่อวัสดุที่พบครั้งแรกมาใช้
        df_grouped = (
            df.groupby('รหัสพัสดุ', as_index=False)
              .agg({
                  'รายการวัสดุ': 'first',
                  'จำนวนรวม': 'sum'
              })
        )

        self.df_result = df_grouped
        self.display_result(df_grouped)
        self.label_status.setText("✅ แปลงข้อมูลเรียบร้อย")

    def display_result(self, df):
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['รายการวัสดุ', 'รหัสพัสดุ', 'จำนวนรวม'])
        for i, row in df.iterrows():
            self.table.setItem(i, 0, QTableWidgetItem(str(row['รายการวัสดุ'])))
            self.table.setItem(i, 1, QTableWidgetItem(str(row['รหัสพัสดุ'])))
            self.table.setItem(i, 2, QTableWidgetItem(str(row['จำนวนรวม'])))
        self.table.resizeColumnsToContents()

    def export_result(self):
        if self.df_result is None:
            self.label_status.setText("❌ ยังไม่มีข้อมูลให้บันทึก")
            return
        path, _ = QFileDialog.getSaveFileName(self, "บันทึกเป็น Excel", "expanded_result.xlsx", "Excel Files (*.xlsx)")
        if path:
            self.df_result.to_excel(path, index=False)
            self.label_status.setText(f"✅ บันทึกไฟล์แล้ว: {path}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = SetExpander()
    window.show()
    sys.exit(app.exec_())
