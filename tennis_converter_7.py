import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def process_file():
    file_path = file_entry.get()
    min_rows_input = min_rows_entry.get()
    min_percentage_input = percentage_entry.get()
    
    if not file_path:
        result_label.config(text="Please select an Excel file.")
        return
    
    try:
        min_rows_values = parse_min_rows_input(min_rows_input)
        min_percentage = float(min_percentage_input)
        
        df = pd.read_excel(file_path, engine="openpyxl", header=None)
        column_n = df.columns[13]  # Column N (index 13)
        column_q = df.columns[16]  # Column Q (index 16)
        column_h = df.columns[7]   # Column H (index 7)

        all_results = []

        for min_rows in min_rows_values:
            grouped = df.groupby([column_n, column_q]).size().reset_index(name="count")
            valid_groups = grouped[grouped['count'] >= min_rows][[column_n, column_q]]
            df_filtered = df.merge(valid_groups, on=[column_n, column_q], how='inner')
            
            for (n, q), group in df_filtered.groupby([column_n, column_q]):
                group = group.sort_index(ascending=False)
                latest_rows = group.head(min_rows)

                over_pct = (latest_rows[column_h] == "OVER").sum() / len(latest_rows) * 100
                under_pct = (latest_rows[column_h] == "UNDER").sum() / len(latest_rows) * 100

                if over_pct >= min_percentage or under_pct >= min_percentage:
                    all_results.append(latest_rows)

        if all_results:
            df_final = pd.concat(all_results).drop_duplicates()
            
            input_dir = os.path.dirname(file_path)
            output_file = os.path.join(input_dir, "filtered_output.xlsx")
            df_final.to_excel(output_file, index=False, header=False)
            
            result_label.config(text=f"Processing complete! Saved at:\n{output_file}")
        else:
            result_label.config(text="No valid data found.")
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}")

def parse_min_rows_input(input_str):
    if "-" in input_str:
        start, end = map(int, input_str.split("-"))
        return list(range(start, end + 1))
    return [int(input_str)]

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, filename)

# GUI setup
root = tk.Tk()
root.title("Excel OVER/UNDER Processor")
root.geometry("460x320")

tk.Label(root, text="Select Excel File:").pack()
file_entry = tk.Entry(root, width=55)
file_entry.pack()
tk.Button(root, text="Browse", command=browse_file).pack(pady=5)

tk.Label(root, text="Min Rows per Group (e.g., 10-40 or 15):").pack()
min_rows_entry = tk.Entry(root)
min_rows_entry.pack()

tk.Label(root, text="Min OVER or UNDER % (e.g., 60):").pack()
percentage_entry = tk.Entry(root)
percentage_entry.pack()

tk.Button(root, text="Process", command=process_file).pack(pady=10)

result_label = tk.Label(root, text="", wraplength=440)
result_label.pack()

root.mainloop()
