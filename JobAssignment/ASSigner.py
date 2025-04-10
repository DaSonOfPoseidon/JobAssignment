import sys
import os
import time
sys.path.append(os.path.join(os.path.dirname(__file__), "embedded_python", "lib"))

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
        raw_name = str(row['Name']).strip()
        name_parts = raw_name.split()
        if len(name_parts) >= 2:
            dropdown_value = f"{name_parts[0].capitalize()} {name_parts[1][0].upper()}"
        else:
            dropdown_value = raw_name.capitalize()

        # Validate WO number
        try:
            wo_number = str(int(raw_wo))
        except (ValueError, TypeError):
            excel_row = index + 2
            print(f"❌ Invalid WO number '{raw_wo}' on spreadsheet line {excel_row}.")
            continue

        url = BASE_URL + wo_number
        print(f"\n🔗 Opening WO #{wo_number} — {url}")
        driver.get(url)

        # Retry logic to ensure correct WO page is loaded
        max_attempts = 3
        matched_wo = False

        for attempt in range(1, max_attempts + 1):
            try:
                # If current URL doesn't match the expected WO, reload it
                if not driver.current_url.strip().endswith(wo_number):
                    driver.get(url)

                displayed_wo_elem = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//td[contains(text(), 'Work Order #:')]/following-sibling::td"))
                )
                displayed_wo = displayed_wo_elem.text.strip()

                if displayed_wo == wo_number:
                    matched_wo = True
                    break
                else:
                    print(f"🟡 Attempt {attempt}: Page shows WO #{displayed_wo}, expected #{wo_number}. Retrying...")
            except Exception as e:
                print(f"🟡 Attempt {attempt}: Unable to find WO number on page. Retrying...")

            time.sleep(5)

        if not matched_wo:
            print(f"❌ Failed to verify correct WO after 3 attempts — skipping WO #{wo_number}")
            continue

        try:
            dropdown = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "AssignEmpID"))
            )
            select = Select(dropdown)

            # Parse first name and last initial with capitalization handling
            parts = dropdown_value.strip().lower().split()
            if not parts:
                print(f"❌ No valid name format for WO #{wo_number} — skipping.")
                continue
            first_name = parts[0]
            last_initial = parts[1][0] if len(parts) > 1 else ""

            matched_option = None
            for option in select.options:
                full_text = option.text.lower().strip()
                if full_text.startswith(first_name) and f"{first_name} {last_initial}" in full_text:
                    matched_option = option.text
                    break

            if not matched_option:
                print(f"❌ No dropdown match for '{dropdown_value}' — skipping WO #{wo_number}")
                continue

            assignments_div = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "AssignmentsList"))
            )
            assigned_names = assignments_div.text.lower()

            if matched_option.lower() in assigned_names:
                print(f"🟡 WO #{wo_number}: '{matched_option}' already assigned — skipping.")
                continue
            elif assigned_names:
                print(f"🟡 WO #{wo_number} has other people assigned — adding '{matched_option}'.")

            # Assign contractor
            select.select_by_visible_text(matched_option)
            add_button = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'Socket')]"))
            )
            add_button.click()

        except Exception as e:
            print(f"❌ Error on WO #{wo_number}: {e}")

    print("\n✅ Done processing work orders.")
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
        print(f"\n❌ Fatal error: {e}")
        input("\nPress Enter to close...")
