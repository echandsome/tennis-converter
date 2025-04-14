import pandas as pd
import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Date format conversion function: (16_Jan_2025) → (01-16)
def convert_date_format(date_str):
    try:
        date_str = date_str.strip("()")
        dt = datetime.datetime.strptime(date_str, "%d_%b_%Y")
        return f"({dt.strftime('%m-%d')})"
    except Exception as e:
        print(f"Date conversion error: {date_str} -> {e}")
        return None

# File processing function
def process_files(tennis_path, _):
    tennis_df = pd.read_excel(tennis_path, header=None)

    # Ensure at least 16 columns (up to column P)
    current_cols = tennis_df.shape[1]
    if current_cols <= 15:
        for i in range(15 - current_cols + 1):
            tennis_df[current_cols + i] = ""

    current_date = None

    for i in range(len(tennis_df)):
        cell = tennis_df.iloc[i, 0]

        # If it's a group header with a date
        if pd.notna(cell) and isinstance(cell, str) and cell.startswith("(") and cell.endswith(")"):
            converted = convert_date_format(cell)
            if converted:
                current_date = converted
                tennis_df.iat[i, 15] = current_date  # <- Insert date into the header row itself
            continue

        # If it's an empty row, reset the group date
        if pd.isna(cell) or (isinstance(cell, float) and pd.isna(cell)):
            current_date = None
            continue

        # Insert the current group date into column P
        if current_date:
            tennis_df.iat[i, 15] = current_date

    # Save the result
    output_path = os.path.join(os.path.dirname(tennis_path), "tennis_with_group_dates.xlsx")
    tennis_df.to_excel(output_path, index=False, header=False)
    return output_path

# GUI-related functions
def browse_tennis():
    path = filedialog.askopenfilename(title="Select xlsx file", filetypes=[("Excel files", "*.xlsx")])
    if path:
        tennis_path_var.set(path)

def run_process():
    tennis_path = tennis_path_var.get()

    if not tennis_path:
        messagebox.showerror("Error", "Please select the xlsx file.")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        output_path = process_files(tennis_path, None)
        result_label.config(text=f"Complete!\nResult saved at:\n{output_path}", fg="green")
    except Exception as e:
        print(str(e))
        result_label.config(text=f"Error occurred: {str(e)}", fg="red")

# Build GUI
root = tk.Tk()
root.title("Tennis Group Date Inserter")
root.geometry("500x230")

tennis_path_var = tk.StringVar()

tk.Button(root, text="Select tennis.xlsx", command=browse_tennis).pack(pady=(10, 0))
tk.Label(root, textvariable=tennis_path_var, wraplength=480).pack()

tk.Button(root, text="Execute", command=run_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
