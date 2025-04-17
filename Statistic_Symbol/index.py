import tkinter as tk
from tkinter import filedialog
import os
from openpyxl import load_workbook
from datetime import datetime, timedelta

def parse_mmdd_to_date(mmdd_str):
    # "(01-07)" → "01-07"
    mmdd_clean = mmdd_str.strip("()")
    try:
        return datetime.strptime(f"{datetime.today().year}-{mmdd_clean}", "%Y-%m-%d")
    except:
        return None

def get_column_n_values_by_name(daily_path, name):
    wb = load_workbook(daily_path, data_only=True)
    ws = wb.active

    results = []

    for row in ws.iter_rows(min_row=1):  # Iterate through all rows
        a_cell = row[0]  # Column A (index 0)
        if a_cell.value == name:
            n_cell = row[13]  # Column N (index 13)
            results.append(n_cell.value)

    return results

def get_filtered_rows(historical_path, symbol, days, count):
    wb = load_workbook(historical_path, data_only=True)
    ws = wb.active

    found = False
    today = datetime.today()
    date_limit = today - timedelta(days=int(days))

    c_count = 0

    for i, row in enumerate(ws.iter_rows(min_row=1), start=1):
        b_cell = row[1]  # Column B

        if found:
            if all(cell.value is None for cell in row):
                break  # Stop when reaching an empty row

            try:
                raw_date = row[0].value  # Column A: e.g., "(01-07)"
                if not isinstance(raw_date, str):
                    continue
                date_cell = parse_mmdd_to_date(raw_date)
                if date_cell is None:
                    continue

                if not (date_limit <= date_cell <= today):
                    continue  # Outside date range

                b_val = row[1].value  # Column B (duplicate with symbol, can be omitted)
                c_val = row[2].value  # Column C

                
                if isinstance(b_val, (int, float)) and isinstance(c_val, (int, float)):
                    diff = int(b_val) - int(c_val)
                    if int(diff) >= int(count):
                        c_count += 1
            except Exception as e:
                print(f"Error parsing row {i}: {e}")
                continue

        elif b_cell.value == symbol:
            found = True  # Start from the matched symbol

    return c_count

def process_daily_format(daily_format_path, daily_path_path, historical_path, date_val, count_val):
    wb = load_workbook(daily_format_path)
    ws = wb.active

    for i, row in enumerate(ws.iter_rows(min_row=1), start=1):
        a_value = row[0].value  # Column A
        symbols = get_column_n_values_by_name(daily_path_path, a_value)
        
        t_count = 0

        for symbol in symbols:
            t_count += get_filtered_rows(historical_path, symbol, date_val, count_val)

        print(f"Row {i} - t_count: {t_count}")

        row[8].value = t_count  # I column (index 8)

    out_path = os.path.join(os.path.dirname(daily_format_path), "Modified_Daily_Format.xlsx")
    wb.save(out_path)

    result_label.config(text=f"Completed!\nSaved to {out_path}", fg="green")

def browse_daily_format():
    path = filedialog.askopenfilename(title="Select Daily_Format.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        daily_format_path_var.set(path)

def browse_daily():
    path = filedialog.askopenfilename(title="Select Daily.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        daily_path_var.set(path)

def browse_historical():
    path = filedialog.askopenfilename(title="Select Historical Stats.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        historical_path_var.set(path)

def run_conversion():
    daily_format_path = daily_format_path_var.get()
    daily_path_path = daily_path_var.get()
    historical_path = historical_path_var.get()
    date_val = date_entry.get()
    count_val = count_entry.get()

    if not daily_format_path or not date_val or not count_val:
        result_label.config(text="Please select all files and input values.", fg="red")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        process_daily_format(daily_format_path, daily_path_path, historical_path, date_val, count_val)
    except Exception as e:
        print(e)
        result_label.config(text=f"Error: {str(e)}", fg="red")

# --- GUI SETUP ---
root = tk.Tk()
root.title("Daily Format Processor")
root.geometry("600x400")

daily_format_path_var = tk.StringVar()
daily_path_var = tk.StringVar()
historical_path_var = tk.StringVar()

tk.Button(root, text="Select Daily_Format.xlsx", command=browse_daily_format).pack(pady=(10, 0))
tk.Label(root, textvariable=daily_format_path_var, wraplength=580).pack()

tk.Button(root, text="Select Daily.xlsx", command=browse_daily).pack(pady=(10, 0))
tk.Label(root, textvariable=daily_path_var, wraplength=580).pack()

tk.Button(root, text="Select Historical Stats.xlsx", command=browse_historical).pack(pady=(10, 0))
tk.Label(root, textvariable=historical_path_var, wraplength=580).pack()

tk.Label(root, text="Enter Date:").pack(pady=(10, 0))
date_entry = tk.Entry(root)
date_entry.pack()

tk.Label(root, text="Enter Count:").pack(pady=(10, 0))
count_entry = tk.Entry(root)
count_entry.pack()

tk.Button(root, text="Run", command=run_conversion, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
