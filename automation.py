import subprocess
import pyautogui
import time
import pandas as pd
import xlrd
from datetime import datetime
from datetime import date
import calendar
import pyperclip
import win32clipboard


def sign_in(meetingId, pswd, topic):
    subprocess.Popen(
        ["C:\\Users\\Harshith\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe"])
    time.sleep(5)

    # Locates to home
    home_btn = pyautogui.locateCenterOnScreen("Assets\\home.png")
    pyautogui.moveTo(home_btn)
    pyautogui.click()
    time.sleep(0.5)

    # Clicks on join button
    join_btn = pyautogui.locateCenterOnScreen("Assets\\join_btn.png")
    pyautogui.moveTo(join_btn)
    pyautogui.click()

    # Enters meeting ID

    time.sleep(0.5)
    pyautogui.typewrite(meetingId)
    time.sleep(0.5)

    # Types my name
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.write(" 10A 9")
    time.sleep(0.5)
    pyautogui.press("enter")

    # Enters password
    time.sleep(7)
    pyautogui.typewrite(pswd)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(40)

    # Types the message
    # if datetime.now().strftime("%H:%M") > "12:00":
    #     win32clipboard.OpenClipboard()
    #     data = win32clipboard.GetClipboardData()
    #     if f"{name} {grade}, Good afternoon ma'am" in data:
            # pyperclip.copy(f"{name} {grade}, Good afternoon ma'am")
    #         win32clipboard.CloseClipboard()

    # else:
    #     win32clipboard.OpenClipboard()
    #     data = win32clipboard.GetClipboardData()
    #     if f"{name} {grade}, Good morning ma'am" in data:
    #         pyperclip.copy(f"{name} {grade}, Good morning ma'am")
    #         win32clipboard.CloseClipboard()

    time.sleep(10)
    print(f"{topic} done")


def main():
    todays_date = date.today().strftime("%d %m %Y")
    born = datetime.strptime(todays_date, "%d %m %Y").weekday()
    day = calendar.day_name[born]
    df = pd.ExcelFile("timings.xls")
    todays_df = pd.read_excel(df, day)
    name = input("Enter your name: ")
    grade = input("Enter your grade(with your section if you have one): ")
    pyperclip.copy(f"{name} {grade}, Good afternoon ma'am")
    pyperclip.copy(f"{name} {grade}, Good morning ma'am")
    print(todays_df)

    while True:
        current_time = datetime.now().strftime("%H-%M")
        if current_time in str(todays_df["Timing"]):
            details = todays_df.loc[todays_df["Timing"] == current_time]
            meetingId = str(details.iloc[0, 1])
            meetingPswd = str(details.iloc[0, 2])
            meetingTopic = str(details.iloc[0, 3])
            sign_in(meetingId, meetingPswd, meetingTopic)

if __name__ == '__main__':
    main()