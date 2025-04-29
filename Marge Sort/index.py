import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import openpyxl

def shorten_and_sort(file_path):
    df = pd.read_excel(file_path)

    # Rename headers
    rename_map = {
        df.columns[0]: "A",
        df.columns[1]: "B",
        "Over": "O",
        "Under": "U",
        "Total": "Total",
        "OVER% (c/e)": "O% (c/e)"
    }
    df = df.rename(columns=rename_map)

    # Sort by 'O% (c/e)' in descending order
    if "O% (c/e)" in df.columns:
        df = df.sort_values(by="O% (c/e)", ascending=False)
    
    return df

def merge_files_columnwise(folder_path):
    all_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not all_files:
        raise Exception("No .xlsx files found in the selected folder.")

    merged_df = pd.DataFrame()

    for idx, file in enumerate(all_files):
        file_path = os.path.join(folder_path, file)
        sorted_df = shorten_and_sort(file_path)

        if idx > 0:
            # Add an empty column as separator
            merged_df = pd.concat([merged_df, pd.DataFrame({f"": [""] * max(len(merged_df), len(sorted_df))})], axis=1)

        # Reset index to handle merging column-wise
        sorted_df = sorted_df.reset_index(drop=True)
        merged_df = pd.concat([merged_df, sorted_df], axis=1)

    parent_dir = os.path.dirname(folder_path)
    output_path = os.path.join(parent_dir, "merged_sorted_output.xlsx")
    merged_df.to_excel(output_path, index=False)

    result_label.config(text=f"Done! Output saved at:\n{output_path}", fg="green")

# GUI handlers
def browse_folder():
    folder_path = filedialog.askdirectory(title="Select folder containing .xlsx files")
    if folder_path:
        folder_path_var.set(folder_path)

def run_merging_process():
    folder_path = folder_path_var.get()
    if not folder_path:
        result_label.config(text="Please select a folder.", fg="red")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        merge_files_columnwise(folder_path)
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}", fg="red")

# GUI layout
root = tk.Tk()
root.title("XLSX Files Columnwise Merger and Sorter")
root.geometry("500x250")

folder_path_var = tk.StringVar()

tk.Button(root, text="Select Folder with .xlsx Files", command=browse_folder).pack(pady=(10, 0))
tk.Label(root, textvariable=folder_path_var, wraplength=480).pack()

tk.Button(root, text="Start Merge and Sort", command=run_merging_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
