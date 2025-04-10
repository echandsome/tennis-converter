import pandas as pd
import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Date format conversion (e.g., "(16_Jan_2025)" → "01-16")
def convert_date_format(date_str):
    try:
        date_str = date_str.strip("()")
        dt = datetime.datetime.strptime(date_str, "%d_%b_%Y")
        return dt.strftime("(%m-%d)")
    except Exception as e:
        print(f"Date conversion error: {date_str} -> {e}")
        return None

def process_files(tennis_path, moon_csv_path):
    tennis_df = pd.read_excel(tennis_path, header=None)
    moon_df = pd.read_csv(moon_csv_path)

    # Add columns Q, R, S if they don't exist
    for col in [16, 17, 18]:
        if col >= tennis_df.shape[1]:
            tennis_df[col] = ""

    moon_info = ["", "", ""]

    for i in range(len(tennis_df)):
        cell = tennis_df.iloc[i, 0]

        if pd.isna(cell):
            continue

        # If it's a group header, update moon info
        if pd.notna(cell) and isinstance(cell, str) and cell.startswith("(") and cell.endswith(")"):
            date_key = convert_date_format(cell)
            moon_row = moon_df[moon_df.iloc[:, 0] == date_key]

            if not moon_row.empty:
                moon_info = moon_row.iloc[0, 0:3].tolist()
                
            else:
                moon_info = ["", "", ""]

        # Insert moon info into columns Q, R, S
        tennis_df.iat[i, 16] = moon_info[0]
        tennis_df.iat[i, 17] = moon_info[1]
        tennis_df.iat[i, 18] = moon_info[2]
        

    # Save
    output_path = os.path.join(os.path.dirname(tennis_path), "tennis_with_full_moon_data.xlsx")
    tennis_df.to_excel(output_path, index=False, header=False)
    return output_path

# GUI functions
def browse_tennis():
    path = filedialog.askopenfilename(title="Select tennis_match.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        tennis_path_var.set(path)

def browse_moon():
    path = filedialog.askopenfilename(title="Select Moon_Phase.csv", filetypes=[("CSV files", "*.csv")])
    if path:
        moon_path_var.set(path)

def run_process():
    tennis_path = tennis_path_var.get()
    moon_path = moon_path_var.get()

    if not tennis_path or not moon_path:
        messagebox.showerror("Error", "Please select both files.")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        output_path = process_files(tennis_path, moon_path)
        result_label.config(text=f"Complete!\nResult saved:\n{output_path}", fg="green")
    except Exception as e:
        print(str(e))
        result_label.config(text=f"Error occurred: {str(e)}", fg="red")

# Create GUI
root = tk.Tk()
root.title("Tennis Match + Full Moon Info")
root.geometry("500x250")

tennis_path_var = tk.StringVar()
moon_path_var = tk.StringVar()

tk.Button(root, text="Select tennis_match.xlsx", command=browse_tennis).pack(pady=(10, 0))
tk.Label(root, textvariable=tennis_path_var, wraplength=480).pack()

tk.Button(root, text="Select Moon_Phase.csv", command=browse_moon).pack(pady=(10, 0))
tk.Label(root, textvariable=moon_path_var, wraplength=480).pack()

tk.Button(root, text="Execute", command=run_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
