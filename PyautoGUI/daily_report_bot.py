import pyautogui
import pyperclip
import time
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import os


# --------------------------------------------------
# 1. Get today's date and time
# --------------------------------------------------

now = datetime.now()

date_time = now.strftime("%Y-%m-%d %H:%M:%S")
today = now.strftime("%Y-%m-%d")

print("Date and Time:", date_time)


# --------------------------------------------------
# 2. Open Chrome
# --------------------------------------------------

print("Opening Chrome...")

pyautogui.hotkey("win", "r")
time.sleep(1)

pyautogui.write("chrome")
pyautogui.press("enter")

time.sleep(4)


# --------------------------------------------------
# 3. Open weather search
# --------------------------------------------------

print("Opening weather page...")

pyautogui.hotkey("ctrl", "l")

pyautogui.write(
    "https://www.accuweather.com/en/in/chennai/206671/weather-forecast/206671",
    interval=0.01
)

pyautogui.press("enter")

time.sleep(5)


# --------------------------------------------------
# 4. Copy the important information
# --------------------------------------------------

print("Copying weather information...")

# Move to the approximate location of the temperature.
#
# IMPORTANT:
# You may need to adjust this coordinate depending
# on your screen resolution and browser layout.

pyautogui.click(500, 300)

time.sleep(1)

# Select visible page text
pyautogui.hotkey("ctrl", "a")

pyautogui.hotkey("ctrl", "c")

time.sleep(1)

page_text = pyperclip.paste()

print("Copied page text:")
print(page_text[:500])


# --------------------------------------------------
# 5. Extract temperature
# --------------------------------------------------

temperature = "Temperature not detected"

lines = page_text.splitlines()

for line in lines:

    line = line.strip()

    if "°C" in line or "°F" in line:

        temperature = line

        break


print("Fetched Data:", temperature)


# --------------------------------------------------
# 6. Create comment
# --------------------------------------------------

comment = "Good for outdoor activities"


# --------------------------------------------------
# 7. Create Excel workbook
# --------------------------------------------------

print("Creating Excel file...")

workbook = Workbook()

sheet = workbook.active

sheet.title = "Daily Report"


# Header row

sheet["A1"] = "Date & Time"
sheet["B1"] = "Fetched Data"
sheet["C1"] = "Comment"


# Data row

sheet["A2"] = date_time
sheet["B2"] = temperature
sheet["C2"] = comment


# --------------------------------------------------
# 8. Format Excel sheet
# --------------------------------------------------

for cell in sheet[1]:

    cell.font = Font(bold=True)

    cell.alignment = Alignment(horizontal="center")


for column in ["A", "B", "C"]:

    sheet.column_dimensions[column].width = 30


# --------------------------------------------------
# 9. Save Excel file
# --------------------------------------------------

filename = f"daily_report_{today}.xlsx"

output_folder = r"D:\Gen_AI_Class\PyautoGUI"

filepath = os.path.join(
    output_folder,
    filename
)

workbook.save(filepath)

print("Excel file saved:")
print(filepath)


# --------------------------------------------------
# 10. Open the Excel file
# --------------------------------------------------

print("Opening Excel file...")

os.startfile(filepath)

time.sleep(5)


# --------------------------------------------------
# 11. Take screenshot
# --------------------------------------------------

print("Taking screenshot...")

screenshot_filename = (
    f"daily_report_{today}_screenshot.png"
)

screenshot_path = os.path.join(
    output_folder,
    screenshot_filename
)

screenshot = pyautogui.screenshot()

screenshot.save(screenshot_path)


print("Screenshot saved:")
print(screenshot_path)

print()
print("====================================")
print("ASSIGNMENT COMPLETED")
print("====================================")
