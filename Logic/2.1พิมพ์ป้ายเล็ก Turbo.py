#พิมพ์คอลัมน์แรก → Tab → 1 → Enter → พิมพ์ s → รอ 0.6 วิ → Enter
def type_entry(text, _):
    import pyautogui
    import time

    pyautogui.write(text, interval=0.05)  # พิมพ์คอลัมน์แรก
    pyautogui.press("tab")               # กด Tab
    pyautogui.write("1")                 # พิมพ์ 1
    pyautogui.press("enter")            # Enter
    pyautogui.write("s")                # พิมพ์ s
    time.sleep(0.6)                     # รอ 0.6 วินาที
    pyautogui.press("enter")            # Enter
