import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), "python-embed", "lib"))

import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import webbrowser

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

# === CONFIGURATION ===
COLUMN_DATE = 0       # Column A
COLUMN_WO = 4         # Column E
COLUMN_DROPDOWN = 7   # Column H
BASE_URL = "http://inside.sockettelecom.com/workorders/view.php?nCount="
CHROMEDRIVER_PATH = os.path.join(os.path.dirname(__file__), "chromedriver.exe")

# === DATE PARSING FUNCTION ===
def flexible_date_parser(date_str):
    try:
        return pd.to_datetime(str(date_str), errors='coerce')
    except:
        return None

# === MAIN PROCESSING FUNCTION ===
def process_workorders(file_path):
    print(f"\nProcessing file: {file_path}")

    df_raw = pd.read_excel(file_path)

    # Pull from fixed columns (A, E, H)
    df = pd.DataFrame()
    df['Date'] = df_raw.iloc[:, COLUMN_DATE].apply(flexible_date_parser)
    df['WO'] = df_raw.iloc[:, COLUMN_WO]
    df['Name'] = df_raw.iloc[:, COLUMN_DROPDOWN]

    df = df.dropna(subset=['Date', 'WO', 'Name'])

    # Show available dates
    unique_dates = sorted(df['Date'].dt.date.unique())
    print("\nAvailable Dates:")
    for d in unique_dates:
        print(f" - {d}")

    # Ask for date range
    start_input = input("\nEnter start date (YYYY-MM-DD): ")
    end_input = input("Enter end date (YYYY-MM-DD): ")
    try:
        start_date = datetime.strptime(start_input, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_input, "%Y-%m-%d").date()
    except ValueError:
        print("Invalid date format. Use YYYY-MM-DD.")
        input("\nPress Enter to close...")
        return

    filtered_df = df[(df['Date'].dt.date >= start_date) & (df['Date'].dt.date <= end_date)]

    if filtered_df.empty:
        print("No matching jobs found for that date range.")
        input("\nPress Enter to close...")
        return

    print(f"\nProcessing {len(filtered_df)} work orders...")

    # Setup Selenium using local ChromeDriver
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)

    for index, row in filtered_df.iterrows():
        raw_wo = row['WO']
        dropdown_value = str(row['Name']).lower().strip()

        # Validate WO number
        try:
            wo_number = str(int(raw_wo))
        except (ValueError, TypeError):
            excel_row = index + 2
            print(f"\u274c Invalid WO number '{raw_wo}' on spreadsheet line {excel_row}.")
            continue

        url = BASE_URL + wo_number
        print(f"\n\U0001F517 Opening WO #{wo_number} — {url}")
        driver.get(url)

        # Wait until page is reachable or check for browser error
        try:
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "AssignmentsList"))
            )
        except:
            if "This site can’t be reached" in driver.page_source or "ERR_" in driver.page_source:
                print(f"\u274c Failed to load WO #{wo_number} — site unreachable.")
                continue

        try:
            dropdown = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "AssignEmpID"))
            )
            select = Select(dropdown)

            # Parse first name and last initial
            parts = dropdown_value.split()
            first_name = parts[0]
            last_initial = parts[1][0] if len(parts) > 1 else ""

            matched_option = None
            for option in select.options:
                full_text = option.text.lower().strip()
                if full_text.startswith(first_name) and f"{first_name} {last_initial}" in full_text:
                    matched_option = option.text
                    break

            if not matched_option:
                print(f"\u274c No dropdown match for '{dropdown_value}' — skipping WO #{wo_number}")
                continue

            # Check assignments
            assignments_div = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "AssignmentsList"))
            )
            assigned_names = assignments_div.text.lower()

            if matched_option.lower() in assigned_names:
                print(f"\U0001F7E1 WO #{wo_number}: '{matched_option}' already assigned — skipping.")
                continue
            elif assigned_names:
                print(f"\U0001F7E1 WO #{wo_number} has other people assigned — adding '{matched_option}'.")

            # Assign contractor
            select.select_by_visible_text(matched_option)
            add_button = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'Socket')]"))
            )
            add_button.click()

        except Exception as e:
            print(f"\u274c Error on WO #{wo_number}: {e}")

    print("\n\u2705 Done processing work orders.")
    input("\nPress Enter to close...")

# === DRAG & DROP GUI ===
def create_gui():
    app = TkinterDnD.Tk()
    app.title("Drop Excel File")
    app.geometry("400x200")

    label = tk.Label(app, text="Drag and drop your Excel file here", width=40, height=10, bg="lightgray")
    label.pack(expand=True, fill="both", padx=10, pady=10)

    def drop(event):
        file_path = event.data.strip('{}')
        if file_path.lower().endswith(('.xlsx', '.xls')):
            label.config(text="Processing...")
            app.update()
            process_workorders(file_path)
            label.config(text="Done. Drop another file or close.")
        else:
            label.config(text="Not a valid Excel file.")

    label.drop_target_register(DND_FILES)
    label.dnd_bind('<<Drop>>', drop)

    app.mainloop()

# === ENTRY POINT ===
if __name__ == "__main__":
    try:
        create_gui()
    except Exception as e:
        print(f"\n\u274c Fatal error: {e}")
        input("\nPress Enter to close...")
