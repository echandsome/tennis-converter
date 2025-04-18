import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def process_tennis_statistics(file_path, output_folder):
    df = pd.read_excel(file_path, header=None)

    result = {}

    for _, row in df.iterrows():
        date = row[15]  # Column P
        name = row[13]  # Column N
        result_val = row[7]  # Column H

        if pd.isna(date) or pd.isna(name) or pd.isna(result_val):
            continue

        date = str(date).strip()
        name = str(name).strip()
        result_val = str(result_val).strip().upper()

        result.setdefault(name, {}).setdefault(date, {"OVER": 0, "UNDER": 0})
        if result_val == "OVER":
            result[name][date]["OVER"] += 1
        elif result_val == "UNDER":
            result[name][date]["UNDER"] += 1

    # Generate block-stacked DataFrame
    final_rows = []
    final_rows.append(["", "Win/Over Count", "Lose/Under Count", "Total Win/Lose", "Win/Over %"])
    for name in sorted(result.keys()):
        final_rows.append(["", name, name, name, name])
        for date in sorted(result[name].keys()):
            stats = result[name][date]
            win = stats["OVER"]
            lose = stats["UNDER"]
            total = win + lose
            percent = f"{(win / total * 100):.0f}%" if total > 0 else ""
            final_rows.append([date, win, lose, total, percent])
        final_rows.extend([[""] * 5] * 3)

    output_df = pd.DataFrame(final_rows)
    base_name = os.path.basename(file_path)
    output_file_name = f"Historical_Block_{base_name}"
    out_path = os.path.join(output_folder, output_file_name)
    output_df.to_excel(out_path, index=False, header=False)
    return out_path

# GUI Functions
def browse_folder():
    path = filedialog.askdirectory(title="Select Folder with Tennis Excel Files")
    if path:
        folder_path_var.set(path)

def run_bulk_process():
    folder_path = folder_path_var.get()
    if not folder_path:
        messagebox.showerror("Error", "Please select a folder with Excel files.")
        return

    result_label.config(text="Processing all files...", fg="blue")
    root.update()

    try:
        files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
        if not files:
            raise Exception("No .xlsx files found in the folder.")

        # Create output folder next to input folder
        parent_folder = os.path.dirname(folder_path)
        output_folder = os.path.join(parent_folder, "Historical_Block")
        os.makedirs(output_folder, exist_ok=True)

        for file in files:
            full_path = os.path.join(folder_path, file)
            process_tennis_statistics(full_path, output_folder)

        result_label.config(text=f"Done!\nProcessed {len(files)} files.\nSaved to:\n{output_folder}", fg="green")
    except Exception as e:
        print(str(e))
        result_label.config(text=f"Error occurred: {str(e)}", fg="red")

# GUI Layout
root = tk.Tk()
root.title("Bulk Tennis Stat Generator")
root.geometry("500x220")

folder_path_var = tk.StringVar()

tk.Button(root, text="Select Folder with Tennis Files", command=browse_folder).pack(pady=(15, 0))
tk.Label(root, textvariable=folder_path_var, wraplength=480).pack()

tk.Button(root, text="Generate Stats for All Files", command=run_bulk_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
