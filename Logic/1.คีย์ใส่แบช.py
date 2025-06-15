# พิมพ์ข้อความแล้วกด tab, พิมพ์ตัวเลข แล้วกด enter
def type_entry(text, amount):
    import pyautogui
    pyautogui.write(text)
    pyautogui.press("tab")
    pyautogui.write(str(amount))