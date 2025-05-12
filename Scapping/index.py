import tkinter as tk
from tkinter import filedialog
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse
import re
from datetime import datetime
import time
import tempfile
import os

def browse_file(entry_widget, filetypes):
    filename = filedialog.askopenfilename(filetypes=filetypes)
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, filename)

def add_year_to_url(url, year):
    parsed = urlparse(url)
    filename = os.path.basename(parsed.path)  # judgeaa01.shtml
    player_id = filename.replace('.shtml', '')  # judgeaa01

    return f"https://www.baseball-reference.com/players/gl.fcgi?id={player_id}&t=b&year={year}"

# --- Function to Load Team Locations ---
def load_team_locations(txt_file):
    """Load team locations from a TXT file into dictionaries for lookup."""
    team_locations = {}
    team_name_to_abbr = {}

    with open(txt_file, 'r') as file:
        lines = [line.strip().lower() for line in file.readlines() if line.strip()]
        for i in range(2, len(lines), 4):
            if i + 3 < len(lines):
                abbr = lines[i].upper()  # Team abbreviation (ATH, BAL, etc.)
                abbr1 = lines[i + 1].upper()  # Team abbreviation (ATH, BAL, etc.)
                team_name = lines[i + 2]  # Full team name (athletics, orioles, etc.)
                location = lines[i + 3]  # Stadium location

                team_locations[abbr] = location
                if abbr != abbr1:
                    team_locations[abbr1] = location

                team_name_to_abbr[team_name] = abbr  # Map full name to abbreviation

    return team_locations, team_name_to_abbr

def parse_date_components(cell_text):
    try:
        first_part = cell_text.strip().split()[0]

        dt = datetime.strptime(first_part, '%Y-%m-%d')
        return dt.year, dt.month, dt.day
    except Exception:
        pass
    return None, None, None 

def process():
    location_file = file1_entry.get()
    stadium_file = file2_entry.get()
    years_back = year_entry.get()
    output_method_mode = output_mode.get()

    save_folder = os.path.dirname(location_file)
    save_folder = os.path.join(save_folder, "Output")
    os.makedirs(save_folder, exist_ok=True)

    if not (location_file and stadium_file and years_back):
        result_label.config(text="Please fill in all input fields.")
        return

    try:
        team_locations, team_name_to_abbr = load_team_locations(stadium_file)

        years_back = int(years_back)
        current_year = datetime.now().year
        years = [current_year - i for i in range(years_back)]

        df = pd.read_csv(location_file, header=None)

        # Load the text file even though it's not used
        with open(stadium_file, 'r', encoding='utf-8') as f:
            stadium_data = f.read()

        options = webdriver.ChromeOptions()
        options.page_load_strategy = 'none'
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        # options.add_argument("--headless")

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        for idx, row in df.iterrows():
            f_name = row[3].capitalize()
            base_url = row[1]  # Column B
            dob_name = row[3].replace("-", "")
            dob_date = row[5]
            dob_location = row[6]
            dob_month, dob_day, dob_year = dob_date.split("/")
            month_names = {
                '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun',
                '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec',
                '1': 'Jan', '2': 'Feb', '3': 'Mar', '4': 'Apr', '5': 'May', '6': 'Jun',
                '7': 'Jul', '8': 'Aug', '9': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
            }
            dob_month = month_names.get(dob_month, "Unknown")

            combined_yes = []
            combined_no = []

            for year in years:
                try:
                    url_with_year = add_year_to_url(base_url, year)

                    driver.get(url_with_year)
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "suppress_headers.row_summable"))
                    )
                    print("---------------")

                    table = driver.find_element(By.CLASS_NAME, "suppress_headers.row_summable")
                    tbody = table.find_element(By.TAG_NAME, "tbody")
                    rows = tbody.find_elements(By.TAG_NAME, "tr")
                    
                    print(f"Number of rows: {len(rows)}")
                    HomeRun_NO = []
                    HomeRun_YES = []

                    for tr in rows:
                        
                        cells = tr.find_elements(By.TAG_NAME, "td")
                        if not cells:
                            cells = tr.find_elements(By.TAG_NAME, "th")
                        
                        _date = cells[2].text.strip()
                        _year, _month, _day = parse_date_components(_date)
                        if not _year:
                            continue

                        home_team = cells[3].text.strip()

                        if cells[4].text.strip() == '@':
                            home_team = cells[5].text.strip()

                        _location = team_locations.get(home_team)

                        _HR = cells[14].text.strip()

                        _month = month_names.get(str(_month), "Unknown")

                        if _HR and float(_HR) > 0:
                            HomeRun_YES.append([
                                _day, _month, _year, _location,
                                "", "", "",  # Three empty columns as separators
                                dob_day, dob_month, dob_year,
                                dob_location
                            ])
                        else:
                            HomeRun_NO.append([
                                _day, _month, _year, _location,
                                "", "", "",  # Three empty columns as separators
                                dob_day, dob_month, dob_year,
                                dob_location
                            ])
                            
                    if output_method_mode == "yearly":
                        csv_path_yes = os.path.join(save_folder, f"{f_name}_HomeRun_{year}_YES.csv")
                        df_output_yes = pd.DataFrame(HomeRun_YES, columns=[
                            "Day", "Month", "Year", "Location",
                            "", "", "",
                            "DOB Day", "DOB Month", "DOB Year",
                            "Birth Location"
                        ])
                        df_output_yes.to_csv(csv_path_yes, index=False)

                        csv_path_no = os.path.join(save_folder, f"{f_name}_HomeRun_{year}_NO.csv")
                        df_output_no = pd.DataFrame(HomeRun_NO, columns=[
                            "Day", "Month", "Year", "Location",
                            "", "", "",
                            "DOB Day", "DOB Month", "DOB Year",
                            "Birth Location"
                        ])
                        df_output_no.to_csv(csv_path_no, index=False)
                    else:
                        combined_yes.extend(HomeRun_YES)
                        combined_no.extend(HomeRun_NO)

                    print(f"✔ Data written for {url_with_year}")
                except Exception as e:
                    print(f"✘ No table for {url_with_year}")

            if output_method_mode == "combined":
                if combined_yes:
                    csv_path_yes = os.path.join(save_folder, f"{f_name}_HomeRun_ALL_YES.csv")
                    df_output_yes = pd.DataFrame(combined_yes, columns=[
                        "Day", "Month", "Year", "Location",
                        "", "", "",
                        "DOB Day", "DOB Month", "DOB Year",
                        "Birth Location"
                    ])
                    df_output_yes.to_csv(csv_path_yes, index=False)

                if combined_no:
                    csv_path_no = os.path.join(save_folder, f"{f_name}_HomeRun_ALL_NO.csv")
                    df_output_no = pd.DataFrame(combined_no, columns=[
                        "Day", "Month", "Year", "Location",
                        "", "", "",
                        "DOB Day", "DOB Month", "DOB Year",
                        "Birth Location"
                    ])
                    df_output_no.to_csv(csv_path_no, index=False)
        driver.quit()

        result_label.config(text=f"Completed! Results saved in temporary file:\n{save_folder}")
    except Exception as e:
        result_label.config(text=f"An error occurred: {str(e)}")

# GUI Configuration
root = tk.Tk()
root.title("Yearly Table Scraper")
root.geometry("520x400")

tk.Label(root, text="Player_DOB_Location.csv file:").pack()
file1_entry = tk.Entry(root, width=60)
file1_entry.pack()
tk.Button(root, text="Browse", command=lambda: browse_file(file1_entry, [("CSV Files", "*.csv")])).pack(pady=2)

tk.Label(root, text="Stadium_Addresses.txt file:").pack()
file2_entry = tk.Entry(root, width=60)
file2_entry.pack()
tk.Button(root, text="Browse", command=lambda: browse_file(file2_entry, [("Text Files", "*.txt")])).pack(pady=2)

tk.Label(root, text="How many years back to fetch data? (e.g., 3)").pack()
year_entry = tk.Entry(root)
year_entry.pack(pady=5)

output_mode = tk.StringVar(value="yearly")

tk.Label(root, text="Output mode:").pack()
tk.Radiobutton(root, text="Save per year", variable=output_mode, value="yearly").pack()
tk.Radiobutton(root, text="Save combined", variable=output_mode, value="combined").pack(pady=(0, 10))

tk.Button(root, text="Start Scraping", command=process).pack(pady=10)

result_label = tk.Label(root, text="", wraplength=500)
result_label.pack()

root.mainloop()
