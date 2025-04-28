import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re

def process_files(daily_path, historical_path, col_choice):
    daily_df = pd.read_excel(daily_path, header=None)
    hist_df = pd.read_excel(historical_path, header=None)

    # Add columns E, F, G (index 4, 5, 6) if they don't exist
    for col in [4, 5, 6]:
        if col >= daily_df.shape[1]:
            daily_df[col] = ""

    group_start = None
    group_data = None
    _is = True
    i = 0
    while _is:

        try:
            cell = daily_df.iat[i, 0]
        except Exception as e:
            print(e)
            _is = False
            cell = "(4_Mar_2025)"
        
        if pd.isna(cell) or (isinstance(cell, str) and cell.startswith("(") and cell.endswith(")") and group_start is not None):
            # End of group
            if group_start is not None:
                over_total = 0
                under_total = 0

                for j in range(group_start + 1, i):
                    player = daily_df.iat[j, 0]
                    match_value = daily_df.iat[j, 13 if col_choice == 'N' else 11]

                    matched = hist_df[(hist_df[0] == player) & (hist_df[13 if col_choice == 'N' else 11] == match_value)]
                    over_count = (matched[7] == "OVER").sum()
                    under_count = (matched[7] == "UNDER").sum()

                    total = over_count + under_count
                    percent = f"{round(over_count / total * 100)}%" if total > 0 else ""

                    daily_df.iat[j, 4] = over_count
                    daily_df.iat[j, 5] = under_count
                    daily_df.iat[j, 6] = percent

                    over_total += over_count
                    under_total += under_count

                # Write group summary (AVG) in the empty row
                total_all = over_total + under_total
                percent_all = f"{round(over_total / total_all * 100)}%" if total_all > 0 else ""

                if not pd.isna(cell) and not pd.isna(daily_df.iat[i-1, 0]):
                    empty_row = pd.DataFrame([[None] * len(daily_df.columns)], columns=daily_df.columns)
                    daily_df = pd.concat([daily_df.iloc[:i], empty_row, daily_df.iloc[i:]]).reset_index(drop=True)

                daily_df.iat[i, 4] = over_total
                daily_df.iat[i, 5] = under_total
                daily_df.iat[i, 6] = percent_all

                if group_data:
                    daily_df.iat[i, 7] = group_data[4]

                group_data = None
                group_start = None            
        else:
            # Group header
            if isinstance(cell, str) and cell.startswith("(") and cell.endswith(")"):
                group_start = i
                if i + 1 < len(daily_df):
                    
                    group_data = daily_df.iloc[i + 1, 14:18].tolist()  # O~R from the next row
                else:
                    group_data = ["", "", "", ""]
                group_data.append(daily_df.iat[i + 1, 7])

        i += 1
    
    # Save result
    output_path = os.path.join(os.path.dirname(daily_path), "Daily_with_stats.xlsx")
    daily_df.to_excel(output_path, index=False, header=False)
    return output_path

# GUI functions
def browse_daily():
    path = filedialog.askopenfilename(title="Select Daily.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        daily_path_var.set(path)

def browse_historical():
    path = filedialog.askopenfilename(title="Select Historical.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        historical_path_var.set(path)

def run_process():
    daily_path = daily_path_var.get()
    historical_path = historical_path_var.get()
    col_choice = col_var.get()

    if not daily_path or not historical_path:
        messagebox.showerror("Error", "Please select both Daily and Historical files.")
        return

    if col_choice not in ['N', 'L']:
        messagebox.showerror("Error", "Please select N or L column.")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        output_path = process_files(daily_path, historical_path, col_choice)
        result_label.config(text=f"Complete!\nSaved to:\n{output_path}", fg="green")
    except Exception as e:
        print(e)
        result_label.config(text=f"Error: {str(e)}", fg="red")

# GUI setup
root = tk.Tk()
root.title("Daily vs Historical Analyzer")
root.geometry("600x300")

daily_path_var = tk.StringVar()
historical_path_var = tk.StringVar()
col_var = tk.StringVar(value='N')

tk.Button(root, text="Select Daily.xlsx", command=browse_daily).pack(pady=(10, 0))
tk.Label(root, textvariable=daily_path_var, wraplength=580).pack()

tk.Button(root, text="Select Historical.xlsx", command=browse_historical).pack(pady=(10, 0))
tk.Label(root, textvariable=historical_path_var, wraplength=580).pack()

# Radio buttons for column selection
frame = tk.Frame(root)
frame.pack(pady=10)
tk.Label(frame, text="Choose column for matching:").pack(side=tk.LEFT)
tk.Radiobutton(frame, text="N", variable=col_var, value='N').pack(side=tk.LEFT, padx=5)
tk.Radiobutton(frame, text="L", variable=col_var, value='L').pack(side=tk.LEFT, padx=5)

tk.Button(root, text="Run", command=run_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()