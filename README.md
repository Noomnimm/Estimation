# Estimation — Material Calculator

โปรแกรมคำนวณและสรุปรายการวัสดุจากฐานข้อมูล Excel รองรับทั้งโปรแกรมเดสก์ท็อปและหน้าเว็บ

## ความสามารถหลัก

- โหลดฐานข้อมูลจากชีต `BaseData`
- เลือกขนาดเสา รหัสหัวเสา และจำนวน
- บันทึกรายการอินพุตได้หลายหน้า
- รวมจำนวนวัสดุตามรายการและรหัสพัสดุ
- โหลดฐานข้อมูลชุด `SET` และแตกเป็นรหัสพัสดุรายรายการ
- ส่งออกผลสรุปเป็นไฟล์ Excel
- โหลด `Newdata.xlsx` และ `Report.xlsx` เป็นฐานข้อมูลเริ่มต้นโดยอัตโนมัติ

## การติดตั้ง

ต้องมี Python 3.10 ขึ้นไป จากนั้นติดตั้ง dependencies:

```powershell
python -m pip install -r requirements.txt
```

## โปรแกรมเดสก์ท็อป

```powershell
python MaterialCalculator.py
```

## เว็บแอป

```powershell
cd web_app
python server.py
```

แล้วเปิด <http://127.0.0.1:8000>

## Deploy บน Render

โปรเจกต์มีไฟล์ `render.yaml` สำหรับ deploy เว็บจาก GitHub โดยตรง:

1. เข้าสู่ระบบ [Render](https://render.com) ด้วย GitHub
2. เลือก **New > Blueprint**
3. เชื่อม repository `Noomnimm/Estimation`
4. ตรวจสอบชื่อบริการแล้วกด **Deploy Blueprint**

Render จะติดตั้ง dependencies จาก `requirements-web.txt` และ deploy ใหม่อัตโนมัติเมื่อ branch `main` มีการเปลี่ยนแปลง

## รูปแบบไฟล์ข้อมูล

ไฟล์ BaseData ต้องมีชีตชื่อ `BaseData` และคอลัมน์ต่อไปนี้:

- `ขนาดเสา (m)`
- `รหัสหัวเสา`
- `รายการวัสดุ`
- `รหัสพัสดุ`
- `จำนวน`

ไฟล์ SET ต้องมีชีตชื่อ `วัสดุทั้งหมด` และคอลัมน์ต่อไปนี้:

- `Set`
- `รหัสพัสดุ`
- `คำอธิบาย`
- `ติดตั้ง`

> ไฟล์ Excel และ Word ถูกยกเว้นจาก Git เพื่อป้องกันการเผยแพร่ข้อมูลจริงโดยไม่ตั้งใจ
