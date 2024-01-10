import pyautogui
import time

def p(s=5):
    print("You got 5 seconds to move the mouse to the target position, adjust the parameter if more time needed..")
    time.sleep(s)
    pos = pyautogui.position()
    print(pos)

p(8)