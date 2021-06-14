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


def sign_in(meetingId, pswd, topic, name, grade, msgPreference):
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
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    print(data)

    pyautogui.hotkey("alt", "h")
    if datetime.now().strftime("%H:%M") > "12:00":
        pyperclip.copy('The text to be copied to the clipboard.')

        pyautogui.typewrite(f"{name} {grade}, Good afternoon ma'am")
    else:
        pyautogui.typewrite("Good morning ma'am, Harshith V 10A")
    time.sleep(10)
    print(f"{topic} done")


date = date.today().strftime("%d %m %Y")
born = datetime.strptime(date, "%d %m %Y").weekday()
day = calendar.day_name[born]
df = pd.ExcelFile("timings.xls")
todays_df = pd.read_excel(df, day)
print(todays_df)

while True:
    name = input("Enter your name: ")
    grade = input("Enter your grade(with your section if you have one): ")
    msgPreference = input(
        "Do you prefer to type your message on your own? (Y or N)")
    current_time = datetime.now().strftime("%H-%M")
    if current_time in str(todays_df["Timings"]):
        details = todays_df.loc[todays_df["Timings"] == current_time]
        meetingId = str(details.iloc[0, 1])
        meetingPswd = str(details.iloc[0, 2])
        meetingTopic = str(details.iloc[0, 3])
        sign_in(meetingId, meetingPswd, meetingTopic,
                name, grade, msgPreference)
