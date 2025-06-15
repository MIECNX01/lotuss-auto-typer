import tkinter as tk
from tkinter import ttk, filedialog
from pynput import keyboard
import threading

root = tk.Tk()
root.title("บันทึกขั้นตอนการพิมพ์/กดคีย์")
root.geometry("770x600")

step_widgets = []
recorded_keys = {}
current_record_step = None

actions = ["พิมพ์ข้อความ", "กดจากคีย์บอร์ด", "รอ (วินาที)"]

def create_step_row(index):
    frame = tk.Frame(step_container)
    frame.pack(pady=5, anchor="w")

    lbl = tk.Label(frame, text=f"ขั้นตอนที่ {index+1}:")
    lbl.pack(side="left")

    var_action = tk.StringVar()
    combo = ttk.Combobox(frame, textvariable=var_action, values=actions, width=20)
    combo.set("เลือกประเภท")
    combo.pack(side="left", padx=5)

    var_value = tk.StringVar()
    entry = tk.Entry(frame, textvariable=var_value, width=25)
    entry.pack(side="left", padx=5)

    record_btn = tk.Button(frame, text="🎧 เริ่มจับคีย์", command=lambda: start_recording(index))
    record_btn.pack(side="left", padx=5)

    reset_btn = tk.Button(frame, text="🔄 รีเซ็ต", command=lambda: reset_step(index))
    reset_btn.pack(side="left", padx=5)

    step_widgets.append((var_action, var_value, entry, record_btn, reset_btn, lbl))

    def on_select(event):
        action = var_action.get()
        if action == "พิมพ์ข้อความ":
            entry.config(state="normal")
            record_btn.config(state="disabled")
        elif action == "กดจากคีย์บอร์ด":
            entry.config(state="readonly")
            record_btn.config(state="normal")
            var_value.set("")
        elif action == "รอ (วินาที)":
            entry.config(state="normal")
            record_btn.config(state="disabled")
            var_value.set("1")

    combo.bind("<<ComboboxSelected>>", on_select)

def reset_step(index):
    var_action, var_value, entry, _, _, _ = step_widgets[index]
    var_value.set("")
    recorded_keys[index] = []
    if var_action.get() == "กดจากคีย์บอร์ด":
        entry.config(state="readonly")
    else:
        entry.config(state="normal")

def generate_steps():
    for widget in step_widgets:
        widget[2].master.destroy()
    step_widgets.clear()
    count = int(step_count_var.get())
    for i in range(count):
        create_step_row(i)
    update_step_labels()

def update_step_labels():
    for i, (_, _, _, _, _, lbl) in enumerate(step_widgets):
        lbl.config(text=f"ขั้นตอนที่ {i+1}:")

def on_key_press(key):
    global current_record_step
    if current_record_step is None:
        return

    try:
        key_name = key.char
    except AttributeError:
        key_name = str(key).replace("Key.", "")

    if current_record_step not in recorded_keys:
        recorded_keys[current_record_step] = []
    recorded_keys[current_record_step].append(key_name)
    step_widgets[current_record_step][1].set(" + ".join(recorded_keys[current_record_step]))

def start_recording(index):
    global current_record_step
    current_record_step = index
    step_widgets[index][1].set("")
    recorded_keys[index] = []
    threading.Thread(target=lambda: keyboard.Listener(on_press=on_key_press).run()).start()

def export_script():
    lines = [
        "import pyautogui",
        "import time",
        "",
        "def auto_typing():"
    ]

    for i, (action_var, value_var, _, _, _, _) in enumerate(step_widgets):
        action = action_var.get()
        value = value_var.get()

        if action == "พิมพ์ข้อความ":
            lines.append(f"    pyautogui.write({repr(value)}, interval=0.05)")
        elif action == "กดจากคีย์บอร์ด":
            keys = value.split(" + ")
            for key in keys:
                lines.append(f"    pyautogui.press({repr(key)})")
        elif action == "รอ (วินาที)":
            try:
                seconds = float(value)
                lines.append(f"    time.sleep({seconds})")
            except ValueError:
                lines.append("    # เวลาไม่ถูกต้อง")

    lines.append("")
    lines.append("auto_typing()")

    file_path = filedialog.asksaveasfilename(defaultextension=".py", filetypes=[("Python Files", "*.py")])
    if file_path:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

# --- ส่วน GUI ---
top_frame = tk.Frame(root)
top_frame.pack(pady=10)

tk.Label(top_frame, text="จำนวนขั้นตอน:").pack(side="left", padx=5)

step_count_var = tk.StringVar(value="3")
tk.Entry(top_frame, textvariable=step_count_var, width=5).pack(side="left", padx=5)

tk.Button(top_frame, text="🔧 สร้างขั้นตอน", command=generate_steps).pack(side="left", padx=5)

step_container = tk.Frame(root)
step_container.pack(fill="both", expand=True)

tk.Button(root, text="📦 ส่งออกสคริปต์", command=export_script, bg="green", fg="white").pack(pady=10)

root.mainloop()
