import os
import time
import pygetwindow as gw
import pyautogui
import re

# # Parameters for HP double screen in office
# pos_pagenum = (561, 129)
# pos_orgpage = (1753, 551)
# pos_extract = (919, 189)
# pos_sepfile = (809, 243)
# pos_extractbtn = (1184, 247)
# pos_extractpath = (957, 700)
# pos_close = (1889, 8)

# Parameters for LG wide screen at home
pos_pagenum = (1236, 150)
pos_orgpage = (3257, 675)
pos_extract = (1661, 222)
pos_sepfile = (1542, 294)
pos_extractbtn = (1982, 294)
pos_extractpath = (1685, 828)
pos_close = (3413, 13)

def export_1page(path_rawpdf, pgnum, folder_extractpdf):
    pattern = r'\d{3}_\d{3}'
    fileid = re.findall(pattern, path_rawpdf)
    if fileid:
        file_path = os.path.join(folder_extractpdf, fileid[0] + ' ' + pgnum + '.pdf')
        if os.path.exists(file_path):
            os.remove(file_path)
        else:
            pass
    else:
        print("No valid file id found in the path.")

    os.startfile(path_rawpdf)
    time.sleep(2)
    window_title = 'Adobe Acrobat Standard'
    try:
        window = gw.getWindowsWithTitle(window_title)[0]
        if not window.isMaximized:
            window.maximize()

        # Step1. Jump to page
        print('Jumping to page...')
        # pyautogui.click(pos_pagenum)
        # time.sleep(2)
        # pyautogui.hotkey('ctrl', 'a')
        # time.sleep(2)
        # pyautogui.press('backspace')
        # time.sleep(2)
        time.sleep(3)
        pyautogui.doubleClick(pos_pagenum)
        time.sleep(2)
        pyautogui.write(pgnum)
        pyautogui.press('enter')

        # Step2. Extract single-page pdf
        print('Extracting...')
        pyautogui.moveTo(pos_orgpage)
        time.sleep(2)
        pyautogui.click()
        time.sleep(2)
        pyautogui.click(pos_extract)
        time.sleep(2)
        pyautogui.click(pos_sepfile)
        time.sleep(2)
        pyautogui.click(pos_extractbtn)
        time.sleep(2)
        pyautogui.click(pos_extractpath)
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(2)
        pyautogui.press('backspace')
        time.sleep(2)
        pyautogui.write(folder_extractpdf)
        time.sleep(4)
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('enter')

        # Step3. Close pdf
        time.sleep(6)
        pyautogui.click(pos_close)
    except IndexError:
        print("No window find")

# if __name__ == "__main__":
#     export_1page(path_rawpdf, folder_extractpdf)