import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import tkinter as tk
from tkinter import filedialog, messagebox

def format_date(csv_date):
    cleaned_date = str(csv_date).replace('"""', '').replace('"', '')
    date_parts = cleaned_date.split('-')
    if len(date_parts) < 2:
        return cleaned_date
    month, day = date_parts[0], date_parts[1]
    if day.startswith('0'):
        day = day[1:]
    if month.startswith('0'):
        month = month[1:]
    return f"{month}/{day}"

def fetch_result_for_player(driver, player_name, date, category):
    try:
        url = f"https://www.bettingpros.com/nba/props/{str(player_name).lower()}/{str(category).lower()}/"
        print(f"Accessing URL: {url}")
        driver.get(url)

        table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//section[4]//table"))
        )
        rows = table.find_elements(By.TAG_NAME, "tr")
        
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) > 0:
                table_date = cols[0].text.strip()
                formatted_date = format_date(date)
                if table_date == formatted_date:
                    stat_value = cols[6].text.strip().replace("O", "").replace("U", "").strip()
                    try:
                        stat_value = int(stat_value)
                    except ValueError:
                        return None
                    print(f"Found result for {player_name}: {stat_value}")
                    return stat_value
        print(f"No matching date found for {player_name}")
    except Exception as e:
        print(f"Error fetching data for {player_name}: {e}")
    return None

def browse_csv():
    csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if csv_path:
        csv_path_entry.delete(0, tk.END)
        csv_path_entry.insert(0, csv_path)

def scrape_results():
    csv_path = csv_path_entry.get()
    if not csv_path:
        messagebox.showerror("Error", "Please select a CSV file.")
        return

    try:
        df = pd.read_csv(csv_path)

        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        category_mapping = {
            "3Pts Made": "threes",
            "Points + Assists": "points-assists",
            "Points + Rebounds": "points-rebounds",
            "Rebounds + Assists": "rebounds-assists",
            "Pts + Ast + Reb": "points-assists-rebounds"
        }
        results = []
        for index, row in df.iterrows():
            player_name = row.get('Player Name', '')
            date = str(row.get('Date', '')).replace('""', '')
            category = row.get('Stat Category', '')

            if category in category_mapping:
                category = category_mapping[category]

            if not isinstance(player_name, str) or not isinstance(category, str):
                continue

            print(f"Fetching data for player: {player_name}, Date: {date}, Category: {category}")

            result = fetch_result_for_player(driver, player_name, date, category)
            if result is not None:
                row['Result'] = result
                row['H/A DIF'] = row.iloc[1] - row.iloc[5]
                row['H/A Results DIF'] = row.iloc[1] - row['Result']
                
                if row['H/A DIF'] != 0:
                    if (row['H/A DIF'] < 0 and row['H/A Results DIF'] < 0) or (row['H/A DIF'] > 0 and row['H/A Results DIF'] > 0):
                        row['H/A Spread Result'] = "Win"
                    else:
                        row['H/A Spread Result'] = "Lose"
                    results.append(row)

        driver.quit()

        if results:
            updated_df = pd.DataFrame(results)
            updated_df = updated_df.sort_values(by='H/A Results DIF', ascending=True)

            negative_df = updated_df[updated_df['H/A DIF'] < 0]
            positive_df = updated_df[updated_df['H/A DIF'] > 0]

            if not negative_df.empty and not positive_df.empty:
                blank_rows = pd.DataFrame([[""] * len(updated_df.columns)] * 2, columns=updated_df.columns)
                updated_df = pd.concat([negative_df, blank_rows, positive_df], ignore_index=True)

            category_name = df['Stat Category'].iloc[0].replace(" ", "_")
            output_path = f"{category_name}_player_props.csv"
            updated_df.to_csv(output_path, index=False)

            messagebox.showinfo("Success", f"Results saved to {output_path}")
        else:
            messagebox.showinfo("Info", "No matching results found. No file created.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

root = tk.Tk()
root.title("NBA Props Scraper")

csv_path_label = tk.Label(root, text="Select CSV File:")
csv_path_label.grid(row=0, column=0, padx=10, pady=10)
csv_path_entry = tk.Entry(root, width=50)
csv_path_entry.grid(row=0, column=1, padx=10, pady=10)
browse_button = tk.Button(root, text="Browse", command=browse_csv)
browse_button.grid(row=0, column=2, padx=10, pady=10)

scrape_button = tk.Button(root, text="Scrape Results!", command=scrape_results, bg="green", fg="white")
scrape_button.grid(row=1, column=1, pady=20)

root.mainloop()
