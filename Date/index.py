import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

# Month name mapping
MONTH_MAP = {
    '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr',
    '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug',
    '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
}

# Function to separate the date
def format_date(date_str):
    if pd.isna(date_str):
        return ['', '', '']
    try:
        date = pd.to_datetime(date_str)
        return [str(date.day).zfill(2), MONTH_MAP[str(date.month).zfill(2)], str(date.year)]
    except:
        return ['', '', '']

# File processing function
def process_file(input_path):
    df = pd.read_excel(input_path, header=0)

    output_rows = []

    for _, row in df.iterrows():
        if pd.isna(row[0]) and pd.isna(row[3]) and pd.isna(row[14]):
            continue  # Remove completely empty rows

        # Date
        date_parts = format_date(row[0])
        # Location
        location = row[3] if not pd.isna(row[3]) else ''
        # O~R columns (index 14~17)
        extra_cols = [
            row[14] if not pd.isna(row[14]) else '',
            row[15] if not pd.isna(row[15]) else '',
            row[16] if not pd.isna(row[16]) else '',
            row[17] if not pd.isna(row[17]) else ''
        ]

        row_data = date_parts + [location] + [''] * 3 + extra_cols
        output_rows.append(row_data)

    # Keep 2 header rows
    header1 = ['Partner A', 'unknown time ON', '', '', '', '', '', 'Partner B', 'unknown time ON', '', '']
    header2 = ['Day', 'Month', 'Year', 'Location', '', '', '', 'Day', 'Month', 'Year', 'Location']
    output_data = [header1, header2] + output_rows

    out_dir = os.path.dirname(input_path)
    output_path = os.path.join(out_dir, "Converted_Astro.xlsx")

    df_out = pd.DataFrame(output_data)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_out.to_excel(writer, index=False, header=False)

    result_label.config(text=f"Conversion complete!\nSaved to:\n{output_path}", fg="green")

# GUI - File selection
def browse_input():
    path = filedialog.askopenfilename(title="Select Input Excel File", filetypes=[("Excel files", "*.xlsx")])
    if path:
        input_path_var.set(path)

# GUI - Run button
def run_conversion():
    input_path = input_path_var.get()
    if not input_path:
        result_label.config(text="Please select the input file first.", fg="red")
        return
    result_label.config(text="Processing...", fg="blue")
    root.update()
    try:
        process_file(input_path)
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}", fg="red")

# --- GUI SETUP ---
root = tk.Tk()
root.title("Astro Excel Converter")
root.geometry("500x220")

input_path_var = tk.StringVar()

tk.Button(root, text="Select Excel File", command=browse_input).pack(pady=(10, 0))
tk.Label(root, textvariable=input_path_var, wraplength=480).pack()

tk.Button(root, text="Convert", command=run_conversion, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
