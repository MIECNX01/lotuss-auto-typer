import os
import sys
import joblib
import subprocess
import pdfplumber
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ติดตั้งไลบรารีที่จำเป็น
for lib in ["pdfplumber", "pandas", "joblib"]:
    try:
        __import__(lib)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

item_codes_global = []
pdf_filename_global = ""
pdf_name_for_ui = ""

MODEL_PATH = os.path.join(os.path.dirname(__file__), "filter_model.pkl")
model = joblib.load(MODEL_PATH) if os.path.exists(MODEL_PATH) else None

def is_sequential(num):
    digits = [int(d) for d in num]
    increasing = all((digits[i] + 1) % 10 == digits[i + 1] % 10 for i in range(len(digits) - 1))
    decreasing = all((digits[i] - 1) % 10 == digits[i + 1] % 10 for i in range(len(digits) - 1))
    return increasing or decreasing

def is_repeated_digits(num):
    return all(d == num[0] for d in num)

def has_low_digit_variance(num):
    return len(set(num)) <= 3

def use_model_filter(num):
    if not model:
        return False
    features = [[int(c) for c in num]]
    prediction = model.predict(features)
    return prediction[0] == 1

def extract_9digit_items(pdf_path):
    global pdf_filename_global, pdf_name_for_ui
    pdf_filename_global = os.path.splitext(os.path.basename(pdf_path))[0]
    pdf_name_for_ui = os.path.basename(pdf_path)

    item_codes = []
    seen = {}
    blacklist_patterns = {"123456789", "456789123", "789123456"}

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    for cell in row:
                        if cell:
                            cleaned_cell = cell.replace(" ", "").strip()
                            found = re.findall(r"\b\d{9}\b", cleaned_cell)
                            for num in found:
                                if (
                                    num.startswith("111") or
                                    num in blacklist_patterns or
                                    is_sequential(num) or
                                    is_repeated_digits(num) or
                                    has_low_digit_variance(num) or
                                    use_model_filter(num)
                                ):
                                    continue
                                if num in seen:
                                    seen[num].append(page_num + 1)
                                else:
                                    seen[num] = [page_num + 1]
                                    item_codes.append(num)
    return item_codes, seen

def save_to_excel(item_codes, output_path):
    df = pd.DataFrame(item_codes)
    df[1] = 0  # คอลัมน์ B = 0
    df.to_excel(output_path, index=False, header=False, startrow=1)

def browse_pdf(entry, name_label):
    path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)
        name_label.config(text=f"ชื่อไฟล์: {os.path.basename(path)}")

def browse_output(entry):
    if not pdf_filename_global:
        messagebox.showwarning("กรุณาเลือก PDF ก่อน", "กรุณาเลือกไฟล์ PDF ก่อนเพื่อใช้ชื่ออ้างอิง")
        return
    initialfile = f"{pdf_filename_global}.xlsx"
    path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=initialfile,
                                        filetypes=[("Excel files", "*.xlsx")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def extract_and_preview(pdf_entry, tree, name_label):
    global item_codes_global
    pdf_path = pdf_entry.get()
    if not os.path.exists(pdf_path):
        messagebox.showwarning("ไฟล์ไม่ถูกต้อง", "กรุณาเลือกไฟล์ PDF ที่ถูกต้อง")
        return

    item_codes, seen = extract_9digit_items(pdf_path)
    item_codes_global = item_codes

    for row in tree.get_children():
        tree.delete(row)

    duplicates = [code for code, pages in seen.items() if len(pages) > 1]
    if duplicates:
        messagebox.showwarning("พบรหัสซ้ำ", f"พบรหัสซ้ำจำนวน {len(duplicates)} รายการ")

    for idx, code in enumerate(item_codes, start=1):
        tag = "duplicate" if code in duplicates else "normal"
        tree.insert("", "end", values=(idx, code), tags=(tag,))
    tree.tag_configure("duplicate", background="yellow")

    name_label.config(text=f"ชื่อไฟล์: {pdf_name_for_ui} | พบ {len(item_codes)} รายการ")
    messagebox.showinfo("เสร็จสิ้น", f"พบทั้งหมด {len(item_codes)} รหัส")

def export_excel(output_entry):
    output_path = output_entry.get()
    if not output_path or not item_codes_global:
        messagebox.showwarning("ข้อมูลไม่ครบ", "โปรดตรวจสอบไฟล์ PDF และข้อมูลที่แสดงก่อนบันทึก")
        return
    save_to_excel(item_codes_global, output_path)
    messagebox.showinfo("บันทึกสำเร็จ", f"ไฟล์บันทึกแล้วที่:\n{output_path}")

def clear_data(tree, name_label):
    global item_codes_global
    item_codes_global = []
    for row in tree.get_children():
        tree.delete(row)
    name_label.config(text="เคลียร์ข้อมูลแล้ว")

def run_app():
    root = tk.Tk()
    root.title("Extract 9-digit Item Codes from PDF")

    tk.Label(root, text="เลือกไฟล์ PDF:").grid(row=0, column=0, sticky="w")
    pdf_entry = tk.Entry(root, width=50)
    pdf_entry.grid(row=0, column=1)
    name_label = tk.Label(root, text="")
    name_label.grid(row=3, column=0, columnspan=3)
    tk.Button(root, text="เรียกดู", command=lambda: browse_pdf(pdf_entry, name_label)).grid(row=0, column=2)

    tk.Label(root, text="บันทึกเป็นไฟล์ Excel:").grid(row=1, column=0, sticky="w")
    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=1, column=1)
    tk.Button(root, text="เรียกดู", command=lambda: browse_output(output_entry)).grid(row=1, column=2)

    tk.Button(root, text="ดึงข้อมูล & แสดงตัวอย่าง",
              command=lambda: extract_and_preview(pdf_entry, tree, name_label)).grid(row=2, column=1, pady=10)
    tk.Button(root, text="บันทึกเป็น Excel", command=lambda: export_excel(output_entry)).grid(row=2, column=2)
    tk.Button(root, text="🗑️ เคลียร์ข้อมูล", command=lambda: clear_data(tree, name_label)).grid(row=2, column=0, pady=10)

    # Scrollbar
    tree_frame = tk.Frame(root)
    tree_frame.grid(row=4, column=0, columnspan=3, pady=10)

    tree_scrollbar = tk.Scrollbar(tree_frame)
    tree_scrollbar.pack(side="right", fill="y")

    tree = ttk.Treeview(tree_frame, columns=("ลำดับ", "รหัสสินค้า"), show="headings", height=15, yscrollcommand=tree_scrollbar.set)
    tree.heading("ลำดับ", text="ลำดับ")
    tree.heading("รหัสสินค้า", text="Item Code (9-digit)")
    tree.pack(side="left")
    tree_scrollbar.config(command=tree.yview)

    root.mainloop()

if __name__ == "__main__":
    run_app()
