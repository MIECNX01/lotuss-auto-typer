import subprocess
import sys
import os
import shutil

# ไลบรารีที่ต้องติดตั้ง
required_packages = {
    "pyautogui": "pyautogui",
    "openpyxl": "openpyxl",
    "PIL": "Pillow",
    "pytesseract": "pytesseract",
    "pandas": "pandas",
    "joblib": "joblib",
    "pdfplumber": "pdfplumber",
    "scikit-learn": "scikit-learn",
    "pynput": "pynput"
}

def check_python_versions():
    print("\n🔎 ตรวจสอบ Python เวอร์ชันที่ติดตั้งอยู่ในเครื่อง...")
    found = {}
    candidates = ["python", "python3", "py", "python3.11", "python3.12", "python3.13"]
    for cmd in candidates:
        path = shutil.which(cmd)
        if path and path not in found:
            try:
                version = subprocess.check_output([cmd, "--version"], stderr=subprocess.STDOUT).decode().strip()
                found[path] = version
            except: pass

    possible_dirs = [
        r"C:\Python38", r"C:\Python39", r"C:\Python310", r"C:\Python311", r"C:\Python312", r"C:\Python313",
        os.path.expanduser(r"~\AppData\Local\Programs\Python\Python311"),
        os.path.expanduser(r"~\AppData\Local\Programs\Python\Python312"),
        os.path.expanduser(r"~\AppData\Local\Programs\Python\Python313"),
    ]

    for d in possible_dirs:
        exe = os.path.join(d, "python.exe")
        if os.path.exists(exe) and exe not in found:
            try:
                version = subprocess.check_output([exe, "--version"], stderr=subprocess.STDOUT).decode().strip()
                found[exe] = version
            except: pass

    if found:
        for path, ver in found.items():
            print(f"📍 {ver} → {path}")
        if len(found) > 1:
            print(f"\n⚠️ พบ Python หลายเวอร์ชัน ({len(found)} เวอร์ชัน)")
            print("🔧 แนะนำ: ควรลบ Python ที่ไม่ได้ใช้งานออก โดยเหลือไว้แค่ 1 เวอร์ชันที่ใช้งานจริงเท่านั้น")
            print("📌 เพื่อให้ระบบทำงานเสถียร และไม่สับสนเวลารันคำสั่ง python")
    else:
        print("❌ ไม่พบ Python")

def update_pip():
    print("\n📦 อัปเดต pip ให้เป็นเวอร์ชันล่าสุด...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
        print("✅ pip อัปเดตเรียบร้อยแล้ว")
    except:
        print("❌ อัปเดต pip ไม่สำเร็จ")

def check_and_offer_python_update():
    print("\n🔍 ตรวจสอบเวอร์ชัน Python ปัจจุบัน...")
    current_version = sys.version.split()[0]
    print(f"📌 คุณใช้ Python เวอร์ชัน: {current_version}")
    print("ℹ️ หากต้องการอัปเดตเวอร์ชันใหม่ แนะนำดาวน์โหลดจาก https://www.python.org/downloads/windows/")

def check_tkinter():
    print("\n🔍 ตรวจสอบว่าใช้งาน tkinter ได้หรือไม่...")
    try:
        import tkinter
        print("✅ พบ tkinter แล้ว (พร้อมใช้งาน GUI)")
    except ImportError:
        print("⚠️ ไม่พบ tkinter (GUI) กรุณาติดตั้งผ่านตัวติดตั้ง Python และเลือก 'tcl/tk'")

def update_packages(packages: dict):
    print("\n📦 ตรวจสอบและติดตั้งไลบรารีที่จำเป็น...")
    for module_name, pip_name in packages.items():
        if not pip_name:
            continue
        try:
            __import__(module_name)
            print(f"✅ มีไลบรารีแล้ว: {pip_name}")
        except ImportError:
            print(f"➕ ติดตั้งไลบรารี: {pip_name} ...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
                print(f"✅ ติดตั้งสำเร็จ: {pip_name}")
            except Exception as e:
                print(f"❌ ติดตั้งไม่สำเร็จ: {pip_name} ({e})")

if __name__ == "__main__":
    check_python_versions()
    update_pip()
    check_and_offer_python_update()
    check_tkinter()
    update_packages(required_packages)
    print("\n✅ เสร็จสิ้น! รีสตาร์ทเครื่องหากมีการติดตั้ง Python ใหม่")
    input("🔚 กด Enter เพื่อปิดหน้าต่าง...")
