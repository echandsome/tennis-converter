import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def process_file():
    file_path = file_entry.get()
    if not file_path:
        result_label.config(text="Please select an Excel file.")
        return
    
    try:
        # Read Excel file
        df = pd.read_excel(file_path, engine="openpyxl", header=None)
        column_n = df.columns[4]  # Column N (index 13)
        column_q = df.columns[5]  # Column Q (index 16)
        column_h = df.columns[7]   # Column H (index 7)

        all_results = []

        # Generate A column 0~100, B column 0~50 repeatedly
        for a in range(101):
            for b in range(51):
                # Group and filter by (A, B) values
                group = df[(df[column_n] == a) & (df[column_q] == b)]
                
                # If no group, add empty values
                if group.empty:
                    all_results.append([a, b, "", "", "", ""])
                else:
                    over_count = (group[column_h] == "OVER").sum()
                    under_count = (group[column_h] == "UNDER").sum()
                    total_count = len(group)
                    over_percentage = (over_count / total_count) * 100 if total_count > 0 else 0

                    # Add results to the list
                    all_results.append([a, b, over_count, under_count, total_count, f"{over_percentage:.2f} %"])

        # If there are results, convert to DataFrame
        if all_results:
            result_df = pd.DataFrame(all_results, columns=["A", "B", "Over", "Under", "Total", "OVER% (c/e)"])

            # Save output file in the same directory as the input file
            input_dir = os.path.dirname(file_path)
            output_file = os.path.join(input_dir, "processed_output.xlsx")
            result_df.to_excel(output_file, index=False)
            
            result_label.config(text=f"Processing complete! Saved at:\n{output_file}")
        else:
            result_label.config(text="No valid data found.")
    
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}")

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

tk.Button(root, text="Process", command=process_file).pack(pady=10)

result_label = tk.Label(root, text="", wraplength=440)
result_label.pack()

root.mainloop()
