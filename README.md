# 📦 Lotus's Tool Suite

เครื่องมืออัตโนมัติที่ช่วยประมวลผลข้อมูลจากรูปภาพ, PDF และ Excel สำหรับใช้งานในงานเอกสาร งานพิมพ์ และงานจัดการข้อมูลในธุรกิจค้าปลีก เช่น Lotus's

---

## 📁 รายการโปรเจกต์ในนี้

| ชื่อไฟล์ | หน้าที่ |
|----------|----------|
| `Demo แปลงข้อความจากรูป V4.py` | แปลงเลขจากรูปภาพโดยใช้ Tesseract OCR พร้อม GUI |
| `Lotus's Auto Typer V5 15-6-68.py` | โปรแกรมพิมพ์อัตโนมัติจาก Excel พร้อมเลือก Logic และหน่วงเวลา |
| `POG2TextV2.py` | ดึงเลขรหัสสินค้า 9 หลักจากไฟล์ PDF และ export เป็น Excel |

---

## ✅ วิธีติดตั้งไลบรารี

1. ติดตั้ง Python เวอร์ชัน 3.10 ขึ้นไป
2. ติดตั้งไลบรารีทั้งหมดจาก `requirements.txt`

```bash
pip install -r requirements.txt
```

> หรือให้โปรแกรมติดตั้งให้อัตโนมัติในครั้งแรกที่เปิด

---

## 🖼 OCR ด้วย Tesseract (Demo V4)

- รองรับการวางภาพจาก Clipboard (Ctrl+V)
- ใช้ Tesseract OCR ตรวจจับเลขที่อยู่ในภาพ
- บันทึกเป็น `.txt` และสามารถ Export เป็น Excel ได้
- รองรับเลขที่เป็นแนวตั้ง เช่น 9 หลัก

📌 ต้องติดตั้ง Tesseract OCR  
[ดาวน์โหลด Tesseract OCR](https://github.com/tesseract-ocr/tesseract)

---

## ⌨️ Auto Typer V5

- พิมพ์ข้อมูลจาก Excel เข้าโปรแกรมอื่นแบบอัตโนมัติ
- รองรับ Logic พิมพ์ที่ปรับแต่งได้ในโฟลเดอร์ `Logic`
- มีระบบจำตำแหน่งล่าสุด (resume)
- เลือกหน่วงเวลาแบบ “ปรับเอง” หรือ “หน่วงอัจฉริยะ”

---

## 📄 ดึงรหัส 9 หลักจาก PDF

- ใช้ `pdfplumber` อ่านข้อมูลจาก PDF
- ตรวจจับเฉพาะรหัสสินค้า 9 หลักที่ไม่อยู่ใน blacklist
- มีฟิลเตอร์ซ้ำ, blacklist และ ML Model กรองตัวเลข
- บันทึกเป็น Excel พร้อมช่องจำนวน = 0

---

## 📚 ไลบรารีที่ใช้ (จาก `requirements.txt`)

```txt
pytesseract
Pillow
openpyxl
pyautogui
pdfplumber
pandas
joblib
```

---

## 📂 โครงสร้างโฟลเดอร์ที่แนะนำ

```
project-folder/
│
├── Logic/
│   └── [ไฟล์ logic สำหรับ Auto Typer].py
├── ข้อมูลถาวร/
│   └── [Excel ที่บันทึกไว้]
├── Demo แปลงข้อความจากรูป V4.py
├── Lotus's Auto Typer V5 15-6-68.py
├── POG2TextV2.py
├── requirements.txt
└── README.md
```


---

## 🛠 วิธีใช้งานเบื้องต้น

### 1. ดาวน์โหลดและติดตั้ง Git
🔗 https://git-scm.com/downloads/win

### 2. ดาวน์โหลดและติดตั้ง Python (สำหรับ Windows)
🔗 https://www.python.org/downloads/windows/

### 3. ดาวน์โหลดและติดตั้ง Tesseract OCR (เวอร์ชันแนะนำ)
🔗 https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.2.0.20220712.exe

### 4. ใช้คำสั่ง Git ด้านล่างเพื่อติดตั้งโปรเจกต์

```bash
git clone https://github.com/MIECNX01/lotuss-auto-typer.git
cd lotuss-auto-typer

# ✅ ติดตั้งไลบรารีที่จำเป็น
pip install -r requirements.txt

# ▶️ เริ่มโปรแกรมหลัก
python main.py           # หรือ
python pic2textv3.py
```
