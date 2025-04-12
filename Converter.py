import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

def convert_files(folder_path, file_type):
    files = os.listdir(folder_path)

    output_folder = os.path.join(os.path.dirname(folder_path), f"{os.path.basename(folder_path)}_converted")
    os.makedirs(output_folder, exist_ok=True)

    converted = 0

    try:
        for file in files:
            file_path = os.path.join(folder_path, file)
            filename_wo_ext = os.path.splitext(file)[0]

            if file_type == 'xlsx' and file.lower().endswith('.xlsx'):
                df = pd.read_excel(file_path)
                df.to_csv(os.path.join(output_folder, filename_wo_ext + ".csv"), index=False)
                converted += 1

            elif file_type == 'csv' and file.lower().endswith('.csv'):
                df = pd.read_csv(file_path)
                df.to_excel(os.path.join(output_folder, filename_wo_ext + ".xlsx"), index=False, header=False)
                converted += 1

        result_label.config(
            text=f"Conversion complete! {converted} file(s) saved in:\n{output_folder}", fg="green"
        )

    except Exception as e:
        result_label.config(text=f"Error: {str(e)}", fg="red")

def browse_folder():
    path = filedialog.askdirectory(title="Select Folder Containing Files")
    if path:
        folder_path_var.set(path)

def run_bulk_conversion():
    folder_path = folder_path_var.get()
    file_type = file_type_var.get()
    if not folder_path:
        result_label.config(text="Please select a folder first.", fg="red")
        return
    if file_type not in ['xlsx', 'csv']:
        result_label.config(text="Please select the file type to convert from.", fg="red")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()
    convert_files(folder_path, file_type)

# --- GUI Setup ---
root = tk.Tk()
root.title("Bulk XLSX ⇄ CSV Converter")
root.geometry("500x300")

folder_path_var = tk.StringVar()
file_type_var = tk.StringVar(value='xlsx') 

tk.Label(root, text="1. Select Folder with Files:").pack(pady=(10, 0))
tk.Button(root, text="Browse Folder", command=browse_folder).pack()
tk.Label(root, textvariable=folder_path_var, wraplength=480).pack()

tk.Label(root, text="2. Select input file format:").pack(pady=(15, 5))
radio_frame = tk.Frame(root)
radio_frame.pack()
tk.Radiobutton(radio_frame, text="XLSX to CSV", variable=file_type_var, value='xlsx').pack(side=tk.LEFT, padx=10)
tk.Radiobutton(radio_frame, text="CSV to XLSX", variable=file_type_var, value='csv').pack(side=tk.LEFT, padx=10)

tk.Button(root, text="Convert", command=run_bulk_conversion, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
