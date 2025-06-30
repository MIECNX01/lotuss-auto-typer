import subprocess
import sys
import os
import platform
import urllib.request
import json
import shutil

# ==== STEP 0: ตรวจสอบและติดตั้ง pkg_resources (ผ่าน setuptools) ====
try:
    import pkg_resources
except ImportError:
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "setuptools"])
        import pkg_resources
    except Exception:
        print("❌ ไม่สามารถติดตั้ง setuptools ได้")
        sys.exit(1)

# ==== ตรวจสอบ tkinter (แยกต่างหาก) ====
def check_tkinter():
    try:
        import tkinter
        return True
    except ImportError:
        print("⚠️ ไม่พบ tkinter (GUI)\n👉 ให้เปิดตัวติดตั้ง Python และเลือก 'Modify' → ติ๊ก 'tcl/tk and IDLE' แล้วกด Install")
        return False

# ==== ฟังก์ชันถามผู้ใช้แบบ y/n ====
def ask_user(prompt):
    try:
        return input(f"{prompt} (y/n): ").strip().lower() == "y"
    except:
        return False

# ==== ตรวจสอบ Python หลายเวอร์ชัน ====
def check_python_versions():
    print("🔍 ตรวจสอบ Python ที่ติดตั้ง...\n")
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
    else:
        print("❌ ไม่พบ Python")

# ==== อัปเดต pip ====
def update_pip():
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
    except: pass

# ==== ตรวจสอบเวอร์ชัน Python ล่าสุด และถามว่าจะอัปเดตไหม ====
def check_and_offer_python_update():
    if platform.system() != "Windows":
        return
    try:
        with urllib.request.urlopen("https://www.python.org/doc/versions/") as res:
            html = res.read().decode()
        import re
        match = re.search(r'Python (\d+\.\d+\.\d+)', html)
        latest_version = match.group(1) if match else None
        current_version = platform.python_version()
        print(f"🧠 Python ปัจจุบัน: {current_version}")
        print(f"🌐 Python ล่าสุด: {latest_version}")
        if latest_version and current_version != latest_version:
            if ask_user(f"❓ อัปเดตเป็น {latest_version} ?"):
                url = f"https://www.python.org/ftp/python/{latest_version}/python-{latest_version}-amd64.exe"
                local = "python_installer.exe"
                urllib.request.urlretrieve(url, local)
                subprocess.run([local, "/quiet", "InstallAllUsers=1", "PrependPath=1"])
    except: pass

# ==== ไลบรารีจำเป็น (ไม่รวม tkinter) ====
required_packages = {
    "pyautogui": "pyautogui",
    "openpyxl": "openpyxl",
    "PIL": "Pillow",
    "pytesseract": "pytesseract",
    "pandas": "pandas",
    "joblib": "joblib",
    "pdfplumber": "pdfplumber",
    "scikit-learn": "scikit-learn"
}

# ==== ตรวจสอบและอัปเดตไลบรารี ====
def update_packages(packages):
    for pip_name in packages.values():
        if not pip_name:
            continue
        try:
            dist = pkg_resources.get_distribution(pip_name)
            installed_version = dist.version
            with urllib.request.urlopen(f"https://pypi.org/pypi/{pip_name}/json") as res:
                data = json.load(res)
            latest_version = data["info"]["version"]
            if installed_version != latest_version:
                if ask_user(f"🔄 {pip_name}: {installed_version} → {latest_version}, อัปเดต?"):
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", pip_name])
        except pkg_resources.DistributionNotFound:
            if ask_user(f"➕ ติดตั้ง {pip_name}?"):
                subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
        except: pass

# ==== MAIN ====
if __name__ == "__main__":
    check_python_versions()
    update_pip()
    check_and_offer_python_update()
    check_tkinter()
    update_packages(required_packages)
    print("\n✅ เสร็จสิ้น! รีสตาร์ทเครื่องหากมีการติดตั้ง Python ใหม่")
