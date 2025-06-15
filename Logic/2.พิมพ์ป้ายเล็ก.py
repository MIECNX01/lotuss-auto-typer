# พิมพ์คอลัมน์แรก → Tab → 1 → Enter → S → รอ 1 วิ → Enter
def type_entry(text, _):
    import pyautogui
    import time

    pyautogui.write(text, interval=0.05)  # พิมพ์คอลัมน์แรก
    pyautogui.press("tab")               # กด Tab
    pyautogui.write("1")                 # พิมพ์ 1
    pyautogui.press("enter")            # Enter
    time.sleep(0.1)

    for _ in range(8):                  # กดลูกศรขึ้น 8 ครั้ง
        pyautogui.press("up")
        time.sleep(0.2)

    pyautogui.press("enter")            # Enter


