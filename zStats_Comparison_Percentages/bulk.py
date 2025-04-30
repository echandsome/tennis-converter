import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def process_files():
    folder_path = folder_entry.get()
    if not folder_path:
        result_label.config(text="Please select a folder containing Excel files.")
        return

    # Create result folder at the same level as input folder
    parent_dir = os.path.dirname(folder_path)
    output_folder = os.path.join(parent_dir, "Processed_Results")
    os.makedirs(output_folder, exist_ok=True)

    processed_count = 0
    error_files = []

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(folder_path, file_name)
            try:
                df = pd.read_excel(file_path, engine="openpyxl", header=None)
                column_n = df.columns[4]  # Column E
                column_q = df.columns[5]  # Column F
                column_h = df.columns[7]  # Column H

                all_results = []

                date = df.iloc[2, 15]

                for a in range(101):
                    for b in range(51):
                        group = df[(df[column_n] == a) & (df[column_q] == b)]

                        if group.empty:
                            all_results.append([date, a, b, "", "", "", ""])
                        else:
                            over_count = (group[column_h] == "OVER").sum()
                            under_count = (group[column_h] == "UNDER").sum()
                            total_count = len(group)
                            over_percentage = (over_count / total_count) * 100 if total_count > 0 else 0
                            all_results.append([date, a, b, over_count, under_count, total_count, f"{over_percentage:.2f} %"])

                if all_results:
                    result_df = pd.DataFrame(all_results, columns=["Date", "A", "B", "Over", "Under", "Total", "OVER% (c/e)"])
                    output_file = os.path.join(output_folder, file_name.replace(".xlsx", "_processed.xlsx"))
                    result_df.to_excel(output_file, index=False)
                    processed_count += 1

            except Exception as e:
                error_files.append(file_name)

    # Output results
    if processed_count > 0:
        result_label.config(text=f"{processed_count} files processed.\nSaved in: {output_folder}")
    else:
        result_label.config(text="No valid Excel files found.")

    if error_files:
        result_label.config(text=result_label.cget("text") + f"\n Errors in: {', '.join(error_files)}")

def browse_folder():
    folder_selected = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, folder_selected)

# GUI setup
root = tk.Tk()
root.title("Bulk Excel OVER/UNDER Processor")
root.geometry("500x350")

tk.Label(root, text="Select Folder Containing Excel Files:").pack(pady=(10, 0))
folder_entry = tk.Entry(root, width=60)
folder_entry.pack(pady=5)
tk.Button(root, text="Browse Folder", command=browse_folder).pack()

tk.Button(root, text="Start Bulk Processing", command=process_files).pack(pady=20)

result_label = tk.Label(root, text="", wraplength=480, justify="left", fg="blue")
result_label.pack()

root.mainloop()
