# พิมพ์คอลัมน์แรก → Tab → 1 → Enter → A4 → รอ 2 วิ → Enter
def type_entry(text, _):
    import pyautogui
    import time

    pyautogui.write(text, interval=0.05)   # พิมพ์คอลัมน์แรก
    pyautogui.press("tab")                 # กด Tab
    pyautogui.write("1")                   # พิมพ์ 1
    pyautogui.press("enter")              # Enter
    time.sleep(0.1)

    pyautogui.keyDown('shift')             # พิม A
    pyautogui.press('a')
    pyautogui.keyUp('shift')

    pyautogui.write("4")                   # พิม 4
    time.sleep(2)                          # รอ 2 วินาที
    pyautogui.press("enter")              # Enter
