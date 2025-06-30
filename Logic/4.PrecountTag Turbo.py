#Turbo Versionหน่วงเวลาไว้0.1วินาที
def type_entry(text, _):
    import pyautogui
    import time

    pyautogui.write(text, interval=0)      # พิมพ์คอลัมน์แรก
    pyautogui.press("tab")                 # ไปช่องถัดไป
    pyautogui.write("1", interval=0)       # พิมพ์เลข 1
    pyautogui.press("enter")               # กด Enter

    # ✅ คำสั่ง "พิมพ์ a"
    pyautogui.press('a')                   # <<<<< ตรงนี้คือการพิมพ์ 'a'

    time.sleep(1.4)                        # รอ 1.4 วินาที
    pyautogui.press("enter")               # กด Enter

