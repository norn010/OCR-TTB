# PDF OCR Web (Typhoon OCR)

เว็บแอปนี้ใช้สำหรับ:
- อัปโหลดไฟล์ PDF
- ใส่รหัสผ่าน PDF (ถ้ามี)
- ส่งไฟล์ไป OCR ผ่าน Typhoon OCR API
- จำ API Key ใน browser เพื่อลดการกรอกซ้ำ
- แสดงผล OCR แบบจัดรูปแบบ (ตัวหนา/ตาราง) พร้อมเวลาในการประมวลผล

## 1) ติดตั้ง

```powershell
cd "C:\Users\EAKSAHA\Downloads\หนอน\OCR TTB"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 2) ตั้งค่า API Key (เลือกแบบใดแบบหนึ่ง)

- ใส่ในหน้าเว็บตอนรัน
- หรือ set env ก่อนรัน

```powershell
$env:TYPHOON_API_KEY="YOUR_KEY"
```

## 3) รันเว็บ

```powershell
python app.py
```

จากนั้นเปิด `http://127.0.0.1:5000`

## ตัวอย่างไฟล์ทดสอบ

- PDF: `C:\Users\EAKSAHA\Downloads\16900394.pdf`
- Password: `002034`

## หมายเหตุ

- ถ้าไฟล์ PDF ไม่ได้ล็อกรหัสผ่าน สามารถปล่อยช่องรหัสผ่านว่างได้
- ช่อง `Pages` เป็น optional และใช้รูปแบบ `1,2,3`
