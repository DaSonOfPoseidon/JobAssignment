import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# === CONFIGURATION ===
COLUMN_DATE = 'A'
COLUMN_WO = 'E'
COLUMN_DROPDOWN = 'H'
BASE_URL = "http://inside.sockettelecom.com/workorders/view.php?nCount="

def flexible_date_parser(date_str):
    try:
        return pd.to_datetime(str(date_str), errors='coerce')
    except:
        return None

def process_workorders(file_path):
    df = pd.read_excel(file_path)
    df[COLUMN_DATE] = df[COLUMN_DATE].apply(flexible_date_parser)
    df = df.dropna(subset=[COLUMN_DATE, COLUMN_WO, COLUMN_DROPDOWN])

    # Show unique available dates
    unique_dates = sorted(df[COLUMN_DATE].dt.date.unique())
    print("\nAvailable Dates:")
    for d in unique_dates:
        print(f" - {d}")

    # Get date range from user
    start_input = input("\nEnter start date (YYYY-MM-DD): ")
    end_input = input("Enter end date (YYYY-MM-DD): ")
    try:
        start_date = datetime.strptime(start_input, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_input, "%Y-%m-%d").date()
    except ValueError:
        print("Invalid date format.")
        return

    filtered_df = df[(df[COLUMN_DATE].dt.date >= start_date) & (df[COLUMN_DATE].dt.date <= end_date)]

    if filtered_df.empty:
        print("⚠️ No matching jobs found in that date range.")
        return

    print(f"\nFound {len(filtered_df)} work orders.")

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    for index, row in filtered_df.iterrows():
        raw_wo = row[COLUMN_WO]

        try:
            wo_number = str(int(raw_wo))
        except (ValueError, TypeError):
            excel_row = index + 2
            print(f"⚠️ Invalid WO number '{raw_wo}' on spreadsheet line {excel_row} — skipping.")
            continue

        dropdown_value = str(row[COLUMN_DROPDOWN]).lower().strip()
        url = BASE_URL + wo_number

        print(f"\n➡️ Opening WO #{wo_number} — Looking for dropdown match: '{dropdown_value}'")
        driver.get(url)

        try:
            dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "select"))
            )
            select = Select(dropdown)

            matched_option = None
            for option in select.options:
                full_text = option.text.lower().strip()
                if dropdown_value in full_text or full_text.startswith(dropdown_value.split()[0]):
                    matched_option = option.text
                    break

            if matched_option:
                select.select_by_visible_text(matched_option)
            else:
                print(f"⚠️ No dropdown match for '{dropdown_value}' — skipping WO #{wo_number}")
                continue

            enter_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' or @value='Enter' or @name='submit']"))
            )
            enter_button.click()

        except Exception as e:
            print(f"Error on WO #{wo_number}: {e}")

    print("\nJobs have been assigned. Please review any issues.")
    # driver.quit()

# === DRAG-AND-DROP GUI ===
def create_gui():
    app = TkinterDnD.Tk()
    app.title("Drop your Excel File")
    app.geometry("400x200")

    label = tk.Label(app, text="Drag and drop your Excel file here", width=40, height=10, bg="lightgray")
    label.pack(expand=True, fill="both", padx=10, pady=10)

    def drop(event):
        file_path = event.data.strip('{}')  # Remove curly braces if present
        if file_path.lower().endswith(('.xlsx', '.xls')):
            label.config(text="Processing...")
            app.update()
            process_workorders(file_path)
            label.config(text="Done! Drop another file or close.")
        else:
            label.config(text="Not a valid Excel file")

    label.drop_target_register(DND_FILES)
    label.dnd_bind('<<Drop>>', drop)

    app.mainloop()

if __name__ == "__main__":
    try:
        from tkinterdnd2 import TkinterDnD
    except ImportError:
        print("You need to install tkinterdnd2 with: pip install tkinterdnd2")
        exit()

    create_gui()
