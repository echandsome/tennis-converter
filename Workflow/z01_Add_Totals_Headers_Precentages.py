import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def process_file():
    file_path = file_entry.get()
    output_format = output_format_var.get()

    if not file_path:
        result_label.config(text="Please select a file.")
        return

    try:
        file_ext = os.path.splitext(file_path)[1].lower()

        # Read file (including header)
        if file_ext == ".xlsx":
            df = pd.read_excel(file_path, engine="openpyxl", header=0)
        elif file_ext == ".csv":
            df = pd.read_csv(file_path, header=0)
        else:
            result_label.config(text="Unsupported file format. Please use .xlsx or .csv.")
            return

        # Save first column name and split Symbol
        original_symbol_col = df.columns[0]
        df = df.rename(columns={original_symbol_col: "Symbol"})

        split_data = df["Symbol"].astype(str).str.split('-', n=1, expand=True)
        df.insert(1, "Symbol Part A", split_data[0])
        df.insert(2, "Symbol Part B", split_data[1])

        num_columns = len(df.columns)
        i = 0
        if num_columns == 7:
            i += 1
        elif num_columns == 8:
            i += 2

        # Check if columns C, D exist
        if df.shape[1] < 6:
            result_label.config(text="Input must have at least columns A to E.")
            return

        col_c = pd.to_numeric(df.iloc[:, 4 + i], errors='coerce')  # C
        col_d = pd.to_numeric(df.iloc[:, 5 + i], errors='coerce')  # D
        col_e = col_c + col_d

        # Column G: Total = C + D
        df["Total"] = col_c.fillna(0) + col_d.fillna(0)

        # Column H: WIN% OVER = C / E (2 decimals, prevent divide-by-zero)
        df["WIN% OVER"] = (col_c / col_e).round(2).fillna(0)

        # Column I: WIN% UNDER = D / E (2 decimals)
        df["WIN% UNDER"] = (col_d / col_e).round(2).fillna(0)

        # Save
        input_dir = os.path.dirname(file_path)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_file = os.path.join(input_dir, f"{base_name}_split_output.{output_format}")

        if output_format == "xlsx":
            df.to_excel(output_file, index=False)
        else:
            df.to_csv(output_file, index=False, encoding="utf-8-sig")

        result_label.config(text=f"File saved successfully:\n{output_file}")

    except Exception as e:
        result_label.config(text=f"Error: {str(e)}")

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel and CSV files", "*.xlsx *.csv")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, filename)

# GUI Setup
root = tk.Tk()
root.title("Symbol Splitter + Calculations")
root.geometry("500x340")

tk.Label(root, text="Select Excel or CSV File:").pack(pady=5)
file_entry = tk.Entry(root, width=60)
file_entry.pack()
tk.Button(root, text="Browse", command=browse_file).pack(pady=5)

tk.Label(root, text="Select Output Format:").pack(pady=5)
output_format_var = tk.StringVar(value="xlsx")
tk.Radiobutton(root, text="XLSX", variable=output_format_var, value="xlsx").pack()
tk.Radiobutton(root, text="CSV", variable=output_format_var, value="csv").pack()

tk.Button(root, text="Process", command=process_file, width=20).pack(pady=10)

result_label = tk.Label(root, text="", wraplength=480, fg="blue")
result_label.pack(pady=10)

root.mainloop()
