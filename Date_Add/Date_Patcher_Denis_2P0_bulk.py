import pandas as pd
import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Convert date format: (16_Jan_2025) → (01-16)
def convert_date_format(date_str):
    try:
        date_str = date_str.strip("()")
        dt = datetime.datetime.strptime(date_str, "%d_%b_%Y")
        return f"({dt.strftime('%m-%d')})"
    except Exception as e:
        print(f"Date conversion error: {date_str} -> {e}")
        return None

# Process one file
def process_single_file(filepath, output_folder):
    df = pd.read_excel(filepath, header=None)

    # Make sure there are at least 16 columns
    if df.shape[1] <= 15:
        for _ in range(15 - df.shape[1] + 1):
            df[df.shape[1]] = ""

    current_date = None

    for i in range(len(df)):
        cell = df.iloc[i, 0]
        if pd.notna(cell) and isinstance(cell, str) and cell.startswith("(") and cell.endswith(")"):
            converted = convert_date_format(cell)
            if converted:
                current_date = converted
                df.iat[i, 15] = current_date
            continue

        if pd.isna(cell):
            current_date = None
            continue

        if current_date:
            df.iat[i, 15] = current_date

    filename = os.path.basename(filepath)
    output_path = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}_with_dates.xlsx")
    df.to_excel(output_path, index=False, header=False)
    return output_path

# Process all .xlsx files in folder
def process_all_in_folder(folder_path):
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not files:
        raise ValueError("No .xlsx files found in the folder.")

    output_folder = os.path.join(os.path.dirname(folder_path), "Outputs")
    os.makedirs(output_folder, exist_ok=True)

    results = []
    for file in files:
        full_path = os.path.join(folder_path, file)
        try:
            output = process_single_file(full_path, output_folder)
            results.append(f"✅ {file} processed")
        except Exception as e:
            results.append(f"❌ {file} failed: {str(e)}")
    return output_folder, results

# GUI Functions
def browse_folder():
    folder = filedialog.askdirectory(title="Select Register Folder")
    if folder:
        folder_path_var.set(folder)

def run_batch_process():
    folder_path = folder_path_var.get()
    if not folder_path:
        messagebox.showerror("Error", "Please select a folder.")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        output_dir, results = process_all_in_folder(folder_path)
        preview = "\n".join(results[-10:])  # Show last 10 results
        result_label.config(
            text=f"✅ Done!\nSaved to: {output_dir}\n\nRecent results:\n{preview}",
            fg="green"
        )
    except Exception as e:
        result_label.config(text=f"❌ Error occurred: {str(e)}", fg="red")

# GUI Setup
root = tk.Tk()
root.title("Tennis Group Date Batch Inserter")
root.geometry("600x300")

folder_path_var = tk.StringVar()

tk.Button(root, text="Select Register Folder", command=browse_folder).pack(pady=(15, 0))
tk.Label(root, textvariable=folder_path_var, wraplength=580).pack(pady=(5, 10))

tk.Button(root, text="Run Batch Process", command=run_batch_process, bg="green", fg="white", width=20).pack(pady=10)

result_label = tk.Label(root, text="", font=("Arial", 10), wraplength=580, justify="left")
result_label.pack(pady=10)

root.mainloop()
