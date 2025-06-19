import subprocess
import sys
import os
import pytesseract
from tkinter import Tk, filedialog, Button, Text, Label, messagebox, Scrollbar, Frame, StringVar, Toplevel
from PIL import Image, ImageGrab, ImageFilter, ImageOps, ImageEnhance
from openpyxl import Workbook
import re

def auto_message(title, message, duration=1500):
    popup = Toplevel()
    popup.title(title)
    popup.geometry("300x100")
    popup.configure(bg="white")
    popup.attributes("-topmost", True)

    Label(popup, text=message, font=("TH Sarabun New", 16), bg="white", fg="green").pack(expand=True)
    popup.after(duration, popup.destroy)

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
    os.startfile("\u0e15\u0e34\u0e14\u0e15\u0e31\u0e49\u0e07\u0e44\u0e25\u0e1a\u0e23\u0e32\u0e23\u0e35.txt")
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
    os.startfile("\u0e15\u0e34\u0e14\u0e15\u0e31\u0e49\u0e07-Tesseract.txt")
    sys.exit()

pytesseract.pytesseract.tesseract_cmd = tesseract_path

# ===== เริ่ม GUI =====
root = Tk()
root.title("📄 OCR ดึงเลขแนวตั้ง (UI แบบสวย)")
root.geometry("600x460")
root.configure(bg="#f4f4f4")

last_number_var = StringVar()
last_number_var.set("เลขล่าสุด: -")

Label(root, textvariable=last_number_var, font=("TH Sarabun New", 18, "bold"), fg="blue", bg="#f4f4f4").pack(pady=(5, 0))
Label(root, text="📸 เลือกรูป หรือ Ctrl+V เพื่อวางจาก Clipboard", font=("TH Sarabun New", 16), bg="#f4f4f4", fg="#333").pack(pady=(0, 10))

frame = Frame(root, bg="#f4f4f4")
frame.pack(padx=10, pady=5)

scrollbar = Scrollbar(frame)
scrollbar.pack(side="right", fill="y")

text_box = Text(
    frame,
    wrap="word",
    font=("TH Sarabun New", 14),
    yscrollcommand=scrollbar.set,
    height=5,
    width=60,
    bg="#ffffff",
    fg="#000000",
    relief="groove",
    bd=2
)
text_box.pack(side="left", fill="y")
scrollbar.config(command=text_box.yview)

all_results = []

def ocr_image(img, filename="Clipboard"):
    try:
        if not isinstance(img, Image.Image):
            messagebox.showerror("Clipboard Error", "ไม่พบภาพจาก Clipboard หรือภาพมีปัญหา")
            return

        img = img.convert("L")
        img = ImageOps.autocontrast(img)
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.5)
        img = img.filter(ImageFilter.SHARPEN)
        img = img.filter(ImageFilter.SHARPEN)
        img = img.resize((int(img.width * 2.5), int(img.height * 2.5)), Image.LANCZOS)

        config = r'--psm 6 -c tessedit_char_whitelist=0123456789'
        text = pytesseract.image_to_string(img, config=config, lang="eng")

        lines = text.strip().split("\n")
        clean_lines = [line.strip() for line in lines if re.fullmatch(r"\d{2,10}", line.strip())]

        result = f"\u0e44\u0e1f\u0e25\u0e4c: {filename}\n" + "\n".join(clean_lines) + "\n" + "-"*40 + "\n"
        text_box.insert("end", result)
        all_results.append([filename, "\n".join(clean_lines)])

        if clean_lines:
            last_number_var.set(f"เลขล่าสุด: {clean_lines[-1]}")

        safe_name = filename.replace("/", "_").replace("\\", "_")
        with open(f"{safe_name}.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(clean_lines))

        item_count = len(clean_lines)
        if item_count == 0:
            messagebox.showwarning("OCR ไม่พบข้อมูล", "❌ ไม่พบเลขที่ตรงเงื่อนไขในภาพ")
        elif item_count < 5:
            messagebox.showwarning("OCR อาจไม่สมบูรณ์", f"📉 พบเพียง {item_count} Items\nตรวจสอบความชัดของภาพ")
        else:
            auto_message("วางสำเร็จ", f"✅ พบทั้งหมด {item_count} ข้อมูล", 1500)

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

    row_count = 0
    for _, content in all_results:
        numbers = re.findall(r"\d{2,10}", content)
        for num in numbers:
            ws.append([num, 0])
            row_count += 1

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        wb.save(save_path)
        messagebox.showinfo("บันทึกสำเร็จ", f"📅 บันทึกไฟล์เรียบร้อยแล้ว\nมีข้อมูลทั้งหมด {row_count} แถว\n📍 ที่อยู่ไฟล์:\n{save_path}")

def clear_all():
    text_box.delete("1.0", "end")
    all_results.clear()
    last_number_var.set("เลขล่าสุด: -")
    messagebox.showinfo("ล้างข้อมูลแล้ว", "ข้อมูลทั้งหมดถูกล้างเรียบร้อย")

btn_style = {"font": ("TH Sarabun New", 14), "width": 25, "bg": "#008080", "fg": "white", "bd": 0}
btn_red = {"font": ("TH Sarabun New", 16, "bold"), "width": 30, "bg": "#cc0000", "fg": "white", "bd": 0}

row1 = Frame(root, bg="#f4f4f4")
row1.pack(pady=8)
Button(row1, text="📁 เลือกภาพ", command=select_image, **btn_style).pack(pady=4)
Button(row1, text="📋 วางจาก Clipboard", command=paste_from_clipboard, **btn_red).pack(pady=4)

row2 = Frame(root, bg="#f4f4f4")
row2.pack(pady=8)
Button(row2, text="📄 Export แนวตั้ง", command=export_numbers_to_excel, **btn_style).pack(pady=4)
Button(row2, text="🗑 เคลียร์ข้อมูล", command=clear_all, **btn_style).pack(pady=4)

text_box.bind('<Control-v>', paste_from_clipboard)
text_box.bind('<Control-V>', paste_from_clipboard)

root.mainloop()
