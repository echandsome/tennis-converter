import pandas as pd
import asyncio
import threading
from concurrent.futures import ThreadPoolExecutor
from playwright.sync_api import sync_playwright
import tkinter as tk
from tkinter import filedialog, messagebox
import time

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


def fetch_result_for_player(player_name, date, category, results, lock):
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()

            url = f"https://www.bettingpros.com/nba/props/{str(player_name).lower()}/{str(category).lower()}/"
            print(f"Accessing URL: {url}")

            for attempt in range(1, 5 + 1):
                try:
                    page.goto(url, timeout=25000)
                    page.wait_for_load_state("networkidle")
                    page.wait_for_selector("section:nth-of-type(4) table", timeout=15000)
                except Exception as e:
                    if attempt == 5:
                        return False
                    else:
                        time.sleep(2) 
           
            rows = page.query_selector_all("section:nth-of-type(4) table tr")
            stat_value = 0
            for row in rows:
                cols = row.query_selector_all("td")
                if len(cols) > 0:
                    table_date = cols[0].inner_text().strip()
                    formatted_date = format_date(date)
                    if table_date == formatted_date:
                        stat_text = cols[6].inner_text().strip().replace("O", "").replace("U", "").strip()
                        try:
                            stat_value = int(stat_text)
                            break
                        except ValueError:
                            continue

            results.append((player_name, date, category, stat_value))
            browser.close()
    except Exception as e:
        print(f"Error fetching data for {player_name}: {e}")


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
        category_mapping = {
            "3Pts Made": "threes",
            "Points + Assists": "points-assists",
            "Points + Rebounds": "points-rebounds",
            "Rebounds + Assists": "rebounds-assists",
            "Pts + Ast + Reb": "points-assists-rebounds"
        }

        results = []
        lock = threading.Lock()

        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = []
            for index, row in df.iterrows():
                player_name = row.get('Player Name', '')
                date = str(row.get('Date', '')).replace('""', '')
                category = row.get('Stat Category', '')

                if category in category_mapping:
                    category = category_mapping[category]

                if not isinstance(player_name, str) or not isinstance(category, str):
                    continue

                futures.append(executor.submit(fetch_result_for_player, player_name, date, category, results, lock))

            for future in futures:
                future.result()

        if results:
            result_df = pd.DataFrame(results, columns=['Player Name', 'Date', 'Stat Category', 'Result'])
            merged_df = pd.merge(df, result_df, on=['Player Name', 'Date', 'Stat Category'])
            merged_df['H/A DIF'] = merged_df.iloc[:, 1] - merged_df.iloc[:, 5]
            merged_df['H/A Results DIF'] = merged_df.iloc[:, 1] - merged_df['Result']

            def spread_result(row):
                if row['H/A DIF'] != 0:
                    if (row['H/A DIF'] < 0 and row['H/A Results DIF'] < 0) or (row['H/A DIF'] > 0 and row['H/A Results DIF'] > 0):
                        return "Win"
                    else:
                        return "Lose"
                return ""

            merged_df['H/A Spread Result'] = merged_df.apply(spread_result, axis=1)

            updated_df = merged_df.sort_values(by='H/A Results DIF', ascending=True)
            negative_df = updated_df[updated_df['H/A DIF'] < 0]
            positive_df = updated_df[updated_df['H/A DIF'] > 0]

            if not negative_df.empty and not positive_df.empty:
                blank_rows = pd.DataFrame([[''] * len(updated_df.columns)] * 2, columns=updated_df.columns)
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
root.title("NBA Props Scraper (Playwright Boosted)")

csv_path_label = tk.Label(root, text="Select CSV File:")
csv_path_label.grid(row=0, column=0, padx=10, pady=10)
csv_path_entry = tk.Entry(root, width=50)
csv_path_entry.grid(row=0, column=1, padx=10, pady=10)
browse_button = tk.Button(root, text="Browse", command=browse_csv)
browse_button.grid(row=0, column=2, padx=10, pady=10)

scrape_button = tk.Button(root, text="Scrape Results!", command=scrape_results, bg="green", fg="white")
scrape_button.grid(row=1, column=1, pady=20)

root.mainloop()
