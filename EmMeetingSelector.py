import datetime
import os
import time

import pyautogui
import win32gui

import AutoLogonRun

os.startfile(r"teams.lnk")
pyautogui.sleep(1)
AutoLogonRun.maximize_current_window()
# AutoLogonRun.bring_to_top(AutoLogonRun.get_pid_by_name("Teams.exe"))
AutoLogonRun.click_first_found(['ui/teams_all_ch_001.png', 'ui/teams_all_ch_002.png', 'ui/teams_all_ch_003.png'],
                               confidence=0.8, grayscale=True, wait=1)
pyautogui.sleep(1)
btn = AutoLogonRun.wait_for_element_appear('ui/ch_em.png', 6)
if btn is not None:
    AutoLogonRun.click_center(btn)
weekday = datetime.datetime.today().weekday()
image = ''
if weekday == 1:
    image = r'ui\em_lez_mar.png'
elif weekday == 2:
    image = r'ui\em_lez_mer.png'
elif weekday == 4:
    image = r'ui\em_lez_ven.png'
elif weekday == 3:  # just for debug
    image = r'ui\em_lez_mer.png'
else:
    exit(0)

btn = AutoLogonRun.wait_for_element_appear(image)

time.sleep(1)
if btn is not None:
    AutoLogonRun.click_center(btn)
else:
    AutoLogonRun.click_center(AutoLogonRun.wait_for_element_appear(r'ui\join_2.png.png'))
time.sleep(1)
