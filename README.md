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


---

# 📘 คู่มือการติดตั้งและใช้งาน (ภาษาไทย)

## 🔧 การติดตั้งที่จำเป็น
1. **ติดตั้ง Python** (หากยังไม่มี)
   - ดาวน์โหลด: https://www.python.org/downloads/windows/
   - ✅ **สำคัญ**: ให้ติ๊ก "Add Python.exe to PATH" ก่อนกด Install

2. **ติดตั้ง Tesseract OCR** (ถ้าต้องการใช้ฟีเจอร์พิมพ์ Precount Tag)
   - ดาวน์โหลด: https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.2.0.20220712.exe

---

## 🧾 การใช้งานโปรแกรม

### @@@ 1. POG2Text
@ คือโปรแกรมสำหรับดึงรหัสสินค้าจาก Planogram เป็นไฟล์ Excel ที่พร้อมใช้งานกับ Auto Typer

1.1 กดปุ่ม "เรียกดู" เพื่อเลือกไฟล์ PDF Planogram  
1.2 กดปุ่ม "ดึงข้อมูล & แสดงตัวอย่าง"  
1.3 เลือกพื้นที่ที่จะเซฟไฟล์ Excel  
1.4 กด "บันทึกเป็น Excel"  

⚠️ **ตรวจสอบความถูกต้องของข้อมูลทุกครั้ง**

---

### @@@ 2. แปลงข้อความจากรูป
@ คือโปรแกรมสำหรับดึงรูปจาก StoreLine มาเป็น Excel ที่พร้อมใช้ใน Auto Typer

2.1 เปิดหน้า StoreLine และย่อให้เหลือครึ่งหน้าจอ  
2.2 เปิดโปรแกรมแปลงข้อความจากรูป และวางทางซ้ายของหน้าต่าง StoreLine  
2.3 เปิด Batch ที่ต้องการคีย์ Precount Tag  
2.4 ขยายหน้าจอ StoreLine เป็น 300%  
2.5 กดปุ่ม Screenshot → ลากเพื่อแคปหน้าจอที่มีเลข Item  
2.6 กดปุ่ม "วางจาก Clipboard" ในโปรแกรม  
2.7 ทำจนครบทุก Item แล้วกด "Export แนวตั้ง" เพื่อได้ Excel พร้อมใช้

⚠️ **ตรวจสอบความถูกต้องของข้อมูลทุกครั้ง**

---

### @@@ 3. Lotus's Auto Typer
@ คือโปรแกรมช่วยพิมพ์ข้อมูลจาก Excel โดยรูปแบบข้อมูลคือ:

- คอลัมน์ A: รหัสสินค้า 9 หลัก (แถวที่ 2 เป็นต้นไป)
- คอลัมน์ B: จำนวนที่ต้องการพิมพ์ (ใส่ 0 หากพิมพ์ป้ายราคา)

3.1 เปิดหน้า StoreLine ย่อขนาดไว้ด้านขวา  
3.2 เปิด Auto Typer วางไว้ด้านซ้าย  
3.3 กด "อัปโหลด Excel" เพื่อโหลดข้อมูล  
3.4 เลือก Logic การพิมพ์  
   - Logic 1 = คีย์ใส่แบช  
   - Logic 2 = พิมพ์ป้ายเล็ก  
   - Logic 3 = พิมพ์ Precount Tag  
3.5 ปรับ "หน่วงเวลาต่อรายการ" (ค่าเริ่มต้น 1 วินาที)  
   - ลดค่าถ้าต้องการให้เร็วขึ้น  
   - หากเร็วเกิน StoreLine อาจพิมพ์ไม่ทัน  
3.6 กด "เริ่มพิมพ์" → โปรแกรมจะนับถอยหลัง 3 วินาที  
   - รีบคลิกที่ช่องใน StoreLine เพื่อให้พิมพ์ในจุดที่ถูกต้อง  
   - **แนะนำให้ทดสอบ 1 รายการก่อน หากพิมพ์ Precount Tag**

⚠️ **ตรวจสอบความถูกต้องของข้อมูลทุกครั้ง**
