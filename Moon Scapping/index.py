import pandas as pd
import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import requests
from bs4 import BeautifulSoup

def convert_to_yyyymm(date_str):
    try:
        date_str = date_str.strip("()").strip()
        dt = datetime.datetime.strptime(date_str, "%m-%d-%Y")
        return dt.strftime("%Y-%m")
    except Exception as e:
        print(f"Date parse error: {date_str} -> {e}")
        return None

def scrape_moon_data(yyyymm_list):
    result = []
    headers = {"User-Agent": "Mozilla/5.0"}

    for yyyymm in yyyymm_list:
        url = f"https://www.almanac.com/astronomy/moon/calendar/AZ/Phoenix/{yyyymm}"
        try:
            print(f"🔍 Scraping: {url}")
            response = requests.get(url, headers=headers)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, "html.parser")
            tds = soup.find_all("td", class_="calday")
            year, month = map(int, yyyymm.split("-"))

            for td in tds:
                try:
                    day_tag = td.find("p", class_="daynumber")
                    if not day_tag or not day_tag.text.strip().isdigit():
                        continue

                    day = int(day_tag.text.strip())
                    phase = None
                    percent = None

                    # Check for 'phasename_minor' first
                    minor = td.find("p", class_="phasename_minor")
                    if minor:
                        parts = list(minor.stripped_strings)
                        if len(parts) == 2:
                            phase = parts[0]
                            percent = parts[1]
                    else:
                        # Check for 'phasename'
                        major = td.find("p", class_="phasename")
                        if major:
                            parts = list(major.stripped_strings)
                            if len(parts) >= 1:
                                phase = parts[0]
                                percent = "50%"  # Default percent if only time is provided

                    result.append({
                        "date": f"{year:04d}-{month:02d}-{day:02d}",
                        "phase": phase,
                        "percent": percent
                    })

                except Exception as td_error:
                    print(f"Error parsing td: {td_error}")

        except Exception as e:
            print(f"Failed to fetch {url} → {e}")

    return result

def normalize_date(date_str):
    try:
        date_str = date_str.strip("()")
        dt = datetime.datetime.strptime(date_str, "%m-%d-%Y")
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None


def process_file(filepath):
    df = pd.read_excel(filepath, header=None)

    # Fill in columns E (4) and F (5) if they don't exist
    while df.shape[1] <= 5:
        df[df.shape[1]] = ""

    yyyymm_set = set()

    for i in range(len(df)):
        cell = df.iat[i, 0]  # Column A (index 0)
        if isinstance(cell, str):
            yyyymm = convert_to_yyyymm(cell)
            if yyyymm:
                yyyymm_set.add(yyyymm)

    moon_map = scrape_moon_data(yyyymm_set) 

    for i in range(len(df)):
        cell = df.iloc[i, 0]

        date = normalize_date(cell)

        if date :
            for moon in moon_map:
                if (date == moon["date"]) :
                    df.iat[i, 4] =  moon["phase"]
                    df.iat[i, 5] =  moon["percent"]

    dir_path = os.path.dirname(filepath)
    base_name = os.path.splitext(os.path.basename(filepath))[0]
    output_path = os.path.join(dir_path, f"{base_name}_moon.xlsx")
    df.to_excel(output_path, index=False, header=False)
    return output_path

# GUI functions
def browse_file():
    path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
    if path:
        file_path_var.set(path)

def run_process():
    file_path = file_path_var.get()
    if not file_path:
        messagebox.showerror("Error", "Please select a file.")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        output_path = process_file(file_path)
        result_label.config(text=f"Complete!\nResult saved:\n{output_path}", fg="green")
    except Exception as e:
        print(str(e))
        result_label.config(text=f"Error occurred: {str(e)}", fg="red")

# GUI Setup
root = tk.Tk()
root.title("Moon Scrapping")
root.geometry("500x200")

file_path_var = tk.StringVar()

tk.Button(root, text="Select Excel File (.xlsx)", command=browse_file).pack(pady=(10, 0))
tk.Label(root, textvariable=file_path_var, wraplength=480).pack()

tk.Button(root, text="Execute", command=run_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
