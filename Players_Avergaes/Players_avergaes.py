import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import openpyxl

def process_grouping(input_path, _unused=None):
    df = pd.read_excel(input_path)

    # Specify columns (A ~ I)
    name_col = df.columns[0]  # Column A
    d_col = df.columns[3]     # Column D
    e_col = df.columns[4]     # Column E
    f_col = df.columns[5]     # Column F
    g_col = df.columns[6]     # Column G
    h_col = df.columns[7]     # Column H
    i_col = df.columns[8]     # Column I

    # Group by name and perform aggregation
    grouped = df.groupby(name_col).agg({
        d_col: 'sum',
        e_col: 'sum',
        f_col: 'sum',
        g_col: 'mean',
        h_col: 'mean',
        i_col: 'mean'
    }).reset_index()

    # Insert empty columns B and C with unique names
    grouped.insert(1, 'Empty_Column_B', '')  # Empty column B
    grouped.insert(2, 'Empty_Column_C', '')  # Empty column C

    # Set output file path
    out_dir = os.path.dirname(input_path)
    out_path = os.path.join(out_dir, "Players_Avergaes.xlsx")

    # Save result to Excel
    grouped.to_excel(out_path, index=False)

     # Auto-adjust column width
    wb = openpyxl.load_workbook(out_path)
    ws = wb.active

    for col in ws.columns:
        header_value = col[0].value
        col_letter = col[0].column_letter

        if not header_value or str(header_value).strip() == "":
            ws.column_dimensions[col_letter].width = 10
            continue

        max_length = 0
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(out_path)

    result_label.config(text=f"Processing complete!\n{out_path}", fg="green")


# --- GUI logic ---
def browse_file():
    path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
    if path:
        file_path_var.set(path)

def run_conversion():
    path = file_path_var.get()
    if not path:
        result_label.config(text="Please select an Excel file first.", fg="red")
        return
    result_label.config(text="Processing...", fg="blue")
    root.update()
    try:
        process_grouping(path, None)
    except Exception as e:
        result_label.config(text=f"Error occurred: {str(e)}", fg="red")


# --- GUI setup ---
root = tk.Tk()
root.title("Excel Group Summary Program")
root.geometry("500x250")

file_path_var = tk.StringVar()

tk.Button(root, text="Select Excel file", command=browse_file).pack(pady=(10, 0))
tk.Label(root, textvariable=file_path_var, wraplength=480).pack()

tk.Button(root, text="Run Conversion", command=run_conversion, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
