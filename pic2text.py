import subprocess
import sys
import os
import pytesseract
from tkinter import Tk, filedialog, Button, Text, Label, messagebox, Scrollbar, Frame
from PIL import Image, ImageGrab, ImageFilter, ImageOps, ImageEnhance
from openpyxl import Workbook
import re

# ===== ตรวจสอบและติดตั้งไลบรารี =====
required_modules = {
    "pytesseract": "pytesseract",
    "PIL": "Pillow",
    "openpyxl": "openpyxl"
}

missing = []
for module, pip_name in required_modules.items():
    try:
        __import__(module)
    except ImportError:
        missing.append(pip_name)

if missing:
    message = (
        "❌ ไลบรารียังไม่ครบ\n\n"
        "ขาด:\n" + "\n".join(f"- {lib}" for lib in missing) +
        "\n\n✅ ติดตั้งโดยใช้:\n\npip install " + " ".join(missing)
    )
    with open("ติดตั้งไลบรารี.txt", "w", encoding="utf-8") as f:
        f.write(message)
    os.startfile("ติดตั้งไลบรารี.txt")
    sys.exit()

# ===== ตรวจสอบ path ของ Tesseract OCR =====
tesseract_paths = [
    r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
    r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe",
    os.path.expanduser(r"~\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe")
]

tesseract_path = None
for path in tesseract_paths:
    if os.path.exists(path):
        tesseract_path = path
        break

if not tesseract_path:
    message = (
        "❌ ไม่พบ Tesseract OCR\n\n"
        "ติดตั้งได้ที่:\nhttps://github.com/tesseract-ocr/tesseract"
    )
    with open("ติดตั้ง-Tesseract.txt", "w", encoding="utf-8") as f:
        f.write(message)
    os.startfile("ติดตั้ง-Tesseract.txt")
    sys.exit()

pytesseract.pytesseract.tesseract_cmd = tesseract_path

# ===== เริ่ม GUI =====
root = Tk()
root.title("OCR ดึงเลขแนวตั้ง (แม่นยำสูงสุด)")
root.geometry("750x700")

Label(root, text="📸 เลือกรูป หรือ Ctrl+V เพื่อวางจาก Clipboard", font=("TH Sarabun New", 16)).pack(pady=5)

frame = Frame(root)
frame.pack(padx=10, pady=10, fill="both", expand=True)

scrollbar = Scrollbar(frame)
scrollbar.pack(side="right", fill="y")

text_box = Text(frame, wrap="word", font=("TH Sarabun New", 14), yscrollcommand=scrollbar.set)
text_box.pack(side="left", fill="both", expand=True)
scrollbar.config(command=text_box.yview)

all_results = []

def ocr_image(img, filename="Clipboard"):
    try:
        # ===== 1. ตรวจสอบภาพ =====
        if not isinstance(img, Image.Image):
            messagebox.showerror("Clipboard Error", "ไม่พบภาพจาก Clipboard หรือภาพมีปัญหา")
            return

        # ===== 2. ปรับภาพให้คมและชัด =====
        img = img.convert("L")  # ขาวดำ
        img = ImageOps.autocontrast(img)
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.5)
        img = img.filter(ImageFilter.SHARPEN)
        img = img.filter(ImageFilter.SHARPEN)
        img = img.resize((int(img.width * 2.5), int(img.height * 2.5)), Image.LANCZOS)

        # ===== 3. OCR เฉพาะเลข =====
        config = r'--psm 6 -c tessedit_char_whitelist=0123456789'
        text = pytesseract.image_to_string(img, config=config, lang="eng")

        # ===== 4. ดึงเลขตั้งแต่ 2–10 หลักเท่านั้น =====
        lines = text.strip().split("\n")
        clean_lines = [line.strip() for line in lines if re.fullmatch(r"\d{2,10}", line.strip())]

        # ===== 5. แสดงผล และบันทึก
        result = f"ไฟล์: {filename}\n" + "\n".join(clean_lines) + "\n" + "-"*40 + "\n"
        text_box.insert("end", result)
        all_results.append([filename, "\n".join(clean_lines)])

        safe_name = filename.replace("/", "_").replace("\\", "_")
        with open(f"{safe_name}.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(clean_lines))

        # ===== 6. แจ้งผล
        item_count = len(clean_lines)
        if item_count == 0:
            messagebox.showwarning("OCR ไม่พบข้อมูล", "❌ ไม่พบเลขที่ตรงเงื่อนไขในภาพ")
        elif item_count < 5:
            messagebox.showwarning("OCR อาจไม่สมบูรณ์", f"📉 พบเพียง {item_count} Items\nตรวจสอบความชัดของภาพ")
        else:
            messagebox.showinfo("วางสำเร็จ", f"✅ พบข้อมูลทั้งหมด {item_count} Items")

    except Exception as e:
        messagebox.showerror("OCR Error", str(e))

def select_image():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
    if file_path:
        img = Image.open(file_path)
        ocr_image(img, filename=os.path.basename(file_path))

def paste_from_clipboard(event=None):
    try:
        img = ImageGrab.grabclipboard()
        ocr_image(img, filename="clipboard_image")
    except Exception as e:
        messagebox.showerror("Clipboard Error", str(e))
    return "break"

def export_numbers_to_excel():
    if not all_results:
        messagebox.showwarning("ไม่มีข้อมูล", "ยังไม่มีข้อความให้บันทึก")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "เลขแนวตั้ง"

    ws.append([])

    for _, content in all_results:
        numbers = re.findall(r"\d{2,10}", content)
        for num in numbers:
            ws.append([num, 0])

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        wb.save(save_path)
        messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกไฟล์ไว้ที่:\n{save_path}")

def clear_all():
    text_box.delete("1.0", "end")
    all_results.clear()
    messagebox.showinfo("ล้างข้อมูลแล้ว", "ข้อมูลทั้งหมดถูกล้างเรียบร้อย")

# ===== ปุ่ม GUI =====
top_frame = Frame(root)
top_frame.pack(pady=10)

row1 = Frame(top_frame)
row1.pack(pady=2)
Button(row1, text="เลือกภาพ", command=select_image, font=("TH Sarabun New", 12), width=20).pack(side="left", padx=5)
Button(row1, text="วางจาก Clipboard", command=paste_from_clipboard, font=("TH Sarabun New", 12), width=20).pack(side="left", padx=5)

row2 = Frame(top_frame)
row2.pack(pady=2)
Button(row2, text="Export แนวตั้ง", command=export_numbers_to_excel, font=("TH Sarabun New", 12), width=20).pack(side="left", padx=5)
Button(row2, text="🗑 เคลียร์ข้อมูล", command=clear_all, font=("TH Sarabun New", 12), width=18).pack(side="left", padx=5)

text_box.bind('<Control-v>', paste_from_clipboard)
text_box.bind('<Control-V>', paste_from_clipboard)

root.mainloop()
# แก้ไขโค้ด OCR เดิมให้ใช้โมเดล YOLO แทน Tesseract
import os
from tkinter import Tk, filedialog, Button, Text, Label, messagebox, Scrollbar, Frame
from PIL import Image, ImageGrab
from ultralytics import YOLO
import re

# โหลดโมเดล YOLOv8 ที่เทรนเสร็จ
model = YOLO("best.pt")  # หรือพาธเต็ม เช่น r"C:\path\to\best.pt"

# เริ่ม GUI
root = Tk()
root.title("YOLO ดึงเลขแนวตั้ง (ใช้โมเดลเทรนเอง)")
root.geometry("750x700")

Label(root, text="📸 เลือกรูป หรือ Ctrl+V เพื่อวางจาก Clipboard", font=("TH Sarabun New", 16)).pack(pady=5)

frame = Frame(root)
frame.pack(padx=10, pady=10, fill="both", expand=True)

scrollbar = Scrollbar(frame)
scrollbar.pack(side="right", fill="y")

text_box = Text(frame, wrap="word", font=("TH Sarabun New", 14), yscrollcommand=scrollbar.set)
text_box.pack(side="left", fill="both", expand=True)
scrollbar.config(command=text_box.yview)

all_results = []

def detect_yolo(img, filename="Clipboard"):
    try:
        if not isinstance(img, Image.Image):
            messagebox.showerror("Image Error", "ไม่พบภาพหรือไม่สามารถอ่านได้")
            return

        img_path = "_temp.jpg"
        img.save(img_path)

        results = model.predict(img_path, conf=0.4, save=False, verbose=False)[0]
        numbers = []

        for box in results.boxes:
            cls = int(box.cls[0])
            numbers.append(cls)

        result = f"ไฟล์: {filename}\n" + "".join(str(n) for n in numbers) + "\n" + "-"*40 + "\n"
        text_box.insert("end", result)
        all_results.append([filename, "".join(str(n) for n in numbers)])

        os.remove(img_path)

        if not numbers:
            messagebox.showwarning("ไม่พบเลข", "❌ ไม่พบตัวเลขในภาพ")
        elif len(numbers) < 5:
            messagebox.showwarning("น้อยเกินไป", f"📉 พบเพียง {len(numbers)} ตัวเลข")
        else:
            messagebox.showinfo("สำเร็จ", f"✅ พบทั้งหมด {len(numbers)} ตัวเลข")

    except Exception as e:
        messagebox.showerror("YOLO Error", str(e))

def select_image():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
    if file_path:
        img = Image.open(file_path)
        detect_yolo(img, filename=os.path.basename(file_path))

def paste_from_clipboard(event=None):
    try:
        img = ImageGrab.grabclipboard()
        detect_yolo(img, filename="clipboard_image")
    except Exception as e:
        messagebox.showerror("Clipboard Error", str(e))
    return "break"

def export_numbers_to_excel():
    from openpyxl import Workbook
    if not all_results:
        messagebox.showwarning("ไม่มีข้อมูล", "ยังไม่มีข้อมูลให้บันทึก")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "เลขแนวตั้ง"

    for _, content in all_results:
        for char in content:
            if char.isdigit():
                ws.append([char, 0])

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        wb.save(save_path)
        messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกที่:\n{save_path}")

def clear_all():
    text_box.delete("1.0", "end")
    all_results.clear()
    messagebox.showinfo("ล้างข้อมูลแล้ว", "ข้อมูลทั้งหมดถูกล้างแล้ว")

# GUI ปุ่ม
row1 = Frame(root)
row1.pack(pady=2)
Button(row1, text="เลือกภาพ", command=select_image, font=("TH Sarabun New", 12), width=20).pack(side="left", padx=5)
Button(row1, text="วางจาก Clipboard", command=paste_from_clipboard, font=("TH Sarabun New", 12), width=20).pack(side="left", padx=5)

row2 = Frame(root)
row2.pack(pady=2)
Button(row2, text="Export แนวตั้ง", command=export_numbers_to_excel, font=("TH Sarabun New", 12), width=20).pack(side="left", padx=5)
Button(row2, text="🗑 เคลียร์ข้อมูล", command=clear_all, font=("TH Sarabun New", 12), width=18).pack(side="left", padx=5)

text_box.bind('<Control-v>', paste_from_clipboard)
text_box.bind('<Control-V>', paste_from_clipboard)

root.mainloop()
