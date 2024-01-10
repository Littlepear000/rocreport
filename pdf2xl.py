import os
import time
import pygetwindow as gw
import pyautogui
import re

# # Parameters for HP double screen in office
# pos_file = (19, 42)
# pos_file_export = (201, 336)
# pos_file_export_spreadsheet = (576, 383)
# pos_file_export_spreadsheet_excel = (913, 385)
# pos_save_path = (417, 122)
# pos_close = (1889, 8)

# Parameters for LG wide screen at home
pos_file = (15, 45)
pos_file_export = (178, 398)
pos_file_export_spreadsheet = (644, 456)
pos_file_export_spreadsheet_excel = (1090, 462)
pos_save_path = (518, 150)
pos_close = (3413, 13)


def pdf2xl(extractpdf_path, savexl_path):
    pattern = r'\d{3}_\d{3} \d*'
    fileid = re.findall(pattern, extractpdf_path)
    if fileid:
        file_path = os.path.join(savexl_path, fileid[0] + '.xlsx')
        if os.path.exists(file_path):
            os.remove(file_path)
        else:
            print(f"The file {file_path} does not exist.")
    else:
        print("No valid file id found in the path.")

    os.startfile(extractpdf_path)
    time.sleep(2)
    window_title = 'Adobe Acrobat Standard'
    try:
        window = gw.getWindowsWithTitle(window_title)[0]
        if not window.isMaximized:
            window.maximize()

        pyautogui.click(pos_file)
        pyautogui.moveTo(pos_file_export)
        time.sleep(1)
        pyautogui.moveTo(pos_file_export_spreadsheet)
        time.sleep(1)
        pyautogui.click(pos_file_export_spreadsheet_excel)
        time.sleep(1)
        pyautogui.click(pos_save_path)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(1)
        pyautogui.press('backspace')
        pyautogui.write(savexl_path)
        time.sleep(3)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.click(pos_close)
    except IndexError:
        print("No window find")