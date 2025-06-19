import customtkinter as ctk
import tkinter as tk
import subprocess
import sys
import os
import urllib.request
import json
import time as systime
import threading
import pyautogui
import openpyxl
import winsound
from tkinter import filedialog, messagebox

# ตรวจสอบและติดตั้งไลบรารีเพิ่มเติมหากจำเป็น
required_packages = ["customtkinter", "pyautogui", "openpyxl", "pillow"]
def ensure_package(package):
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

for pkg in required_packages:
    ensure_package(pkg)

# ตรวจสอบ pip เวอร์ชันล่าสุด
try:
    current_version = subprocess.check_output([sys.executable, "-m", "pip", "--version"]).decode()
    response = urllib.request.urlopen("https://pypi.org/pypi/pip/json")
    latest_version = json.load(response)["info"]["version"]
    if latest_version not in current_version:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
except Exception as e:
    print(f"ไม่สามารถอัปเดต pip ได้: {e}")

# กำหนดโฟลเดอร์สำหรับข้อมูลและ logic
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FOLDER = os.path.join(BASE_DIR, "ข้อมูลถาวร")
LOGIC_FOLDER = os.path.join(BASE_DIR, "Logic")
os.makedirs(DATA_FOLDER, exist_ok=True)
os.makedirs(LOGIC_FOLDER, exist_ok=True)

# โหลด logic ที่พร้อมใช้งาน
logic_descriptions, available_logics = {}, []
for file in os.listdir(LOGIC_FOLDER):
    if file.endswith(".py"):
        name = file[:-3]
        with open(os.path.join(LOGIC_FOLDER, file), encoding="utf-8") as f:
            desc = f.readline().strip()
            logic_descriptions[name] = desc[1:].strip() if desc.startswith("#") else "ไม่มีคำอธิบาย"
        available_logics.append(name)

# เริ่มต้นสร้าง GUI
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("green")
app = ctk.CTk()
app.title("Lotus's Auto Typer V5 Pro")
app.geometry("800x600")

# ตัวแปรหลัก
data_list = []
typing_mode = tk.StringVar(value=available_logics[0] if available_logics else "")
delay_mode = tk.StringVar(value="ปรับเอง")
logic_description_var = tk.StringVar(value=logic_descriptions.get(typing_mode.get(), ""))
saved_file_var = tk.StringVar()
stop_typing_flag = False
delay_var = tk.DoubleVar(value=1.0)
progress_var = tk.DoubleVar()

# โหลด logic function
def load_typing_logic(mode):
    path = os.path.join(LOGIC_FOLDER, f"{mode}.py")
    namespace = {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            exec(f.read(), namespace)
        return namespace.get("type_entry")
    except Exception as e:
        messagebox.showerror("โหลด Logic ล้มเหลว", str(e))
        return None

# ฟังก์ชันอัปโหลดไฟล์ Excel
def upload_file():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            sheet = wb.active
            data_list.clear()
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    data_list.append((str(row[0]), str(row[1]) if row[1] else "0"))
            refresh_display()
            messagebox.showinfo("อัปโหลด", f"โหลด {os.path.basename(path)} เรียบร้อย")
        except Exception as e:
            messagebox.showerror("โหลดล้มเหลว", str(e))

# ฟังก์ชันแสดงผลข้อมูลในหน้าจอ
text_box = None
def refresh_display():
    if text_box:
        text_box.configure(state="normal")
        text_box.delete("1.0", tk.END)
        for i, (t, a) in enumerate(data_list, 1):
            text_box.insert(tk.END, f"{i}. {t} → {a}\n")
        text_box.configure(state="disabled")

# ฟังก์ชันนับถอยหลังก่อนเริ่มพิมพ์
def countdown(seconds):
    for i in range(seconds, 0, -1):
        progress_label.configure(text=f"เริ่มพิมพ์ใน {i} วินาที...")
        winsound.Beep(1000, 500)
        app.update()
        systime.sleep(1)

# ฟังก์ชันเริ่มพิมพ์
def start_typing():
    global stop_typing_flag
    stop_typing_flag = False

    if not data_list:
        messagebox.showwarning("ไม่มีข้อมูล", "กรุณาอัปโหลดข้อมูลก่อน")
        return

    logic_func = load_typing_logic(typing_mode.get())
    if not logic_func:
        return

    def type_thread():
        countdown(3)
        total_items = len(data_list)

        auto_delay = delay_var.get()
        if delay_mode.get() == "หน่วงแบบอัจฉริยะ":
            if total_items > 1000:
                auto_delay = 5.0
            elif total_items > 500:
                auto_delay = 3.0

        start_time = systime.time()
        for idx, (text, amount) in enumerate(data_list):
            if stop_typing_flag:
                break

            logic_func(text, amount)
            pyautogui.press("enter")
            systime.sleep(auto_delay)

            progress = ((idx + 1) / total_items) * 100
            progress_var.set(progress)
            progress_label.configure(text=f"{progress:.1f}% เสร็จแล้ว")
            refresh_display()

        elapsed = systime.time() - start_time
        mins, secs = divmod(elapsed, 60)
        messagebox.showinfo("เสร็จสิ้น", f"ใช้เวลา {int(mins)} นาที {secs:.2f} วินาที")

    threading.Thread(target=type_thread).start()

# ฟังก์ชันหยุดการพิมพ์
def stop_typing():
    global stop_typing_flag
    stop_typing_flag = True

# รีเซ็ตโปรแกรม
def reset_index():
    data_list.clear()
    refresh_display()
    messagebox.showinfo("รีเซ็ต", "ข้อมูลถูกล้างแล้ว")

def reset_and_restart():
    reset_index()
    messagebox.showinfo("รีโหลด", "กำลังรีโหลดโปรแกรมใหม่...")
    app.destroy()
    os.execl(sys.executable, sys.executable, *sys.argv)

# ==== UI ELEMENTS ====
frame_top = ctk.CTkFrame(app)
frame_top.pack(pady=10, padx=10, fill="x")

ctk.CTkButton(frame_top, text="📂 อัปโหลด Excel", command=upload_file).pack(side="left", padx=5)
ctk.CTkButton(frame_top, text="✅ เริ่มพิมพ์", command=start_typing, fg_color="green").pack(side="left", padx=5)
ctk.CTkButton(frame_top, text="⛔ หยุดพิมพ์", command=stop_typing, fg_color="red").pack(side="left", padx=5)
ctk.CTkButton(frame_top, text="🔁 เริ่มใหม่", command=reset_index, fg_color="#FFC107", text_color="black").pack(side="left", padx=5)
ctk.CTkButton(frame_top, text="🔄 รีโหลด", command=reset_and_restart, fg_color="#FF5722").pack(side="left", padx=5)

frame_logic = ctk.CTkFrame(app)
frame_logic.pack(pady=5, fill="x", padx=10)
ctk.CTkLabel(frame_logic, text="🧠 เลือก Logic:").pack(side="left", padx=5)
logic_combo = ctk.CTkComboBox(frame_logic, values=available_logics, command=lambda _: logic_description_var.set(logic_descriptions.get(typing_mode.get(), "")))
logic_combo.set(typing_mode.get())
logic_combo.pack(side="left", padx=5, fill="x", expand=True)

logic_desc_label = ctk.CTkLabel(app, textvariable=logic_description_var, text_color="gray")
logic_desc_label.pack(pady=2)

frame_delay = ctk.CTkFrame(app)
frame_delay.pack(pady=5, padx=10, fill="x")
ctk.CTkLabel(frame_delay, text="⏱️ หน่วงเวลา (วิ):").pack(side="left", padx=5)
ctk.CTkEntry(frame_delay, textvariable=delay_var, width=80).pack(side="left", padx=5)
delay_combo = ctk.CTkComboBox(frame_delay, values=["ปรับเอง", "หน่วงแบบอัจฉริยะ"])
delay_combo.set(delay_mode.get())
delay_combo.pack(side="left", padx=5)

frame_display = ctk.CTkFrame(app)
frame_display.pack(padx=10, pady=10, fill="both", expand=True)
text_box = tk.Text(frame_display, state="disabled", wrap="none", bg="#f0f0f0")
text_box.pack(fill="both", expand=True)

progress_frame = ctk.CTkFrame(app)
progress_frame.pack(fill="x", padx=20, pady=5)
progress_bar = ctk.CTkProgressBar(progress_frame, variable=progress_var)
progress_bar.pack(fill="x", expand=True)
progress_label = ctk.CTkLabel(app, text="0.0%")
progress_label.pack()

app.mainloop()
