import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def process_tennis_statistics(file_path):
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
    final_rows.append(["", "Win/Over Count", "Win/Over Count", "Total Win/Lose", "Win/Over %"])
    for name in sorted(result.keys()):
        # Header rows
        final_rows.append(["", name, name, name, name])

        # Each row = one date for the player
        for date in sorted(result[name].keys()):
            stats = result[name][date]
            win = stats["OVER"]
            lose = stats["UNDER"]
            total = win + lose
            percent = f"{(win / total * 100):.0f}%" if total > 0 else ""
            final_rows.append([date, win, lose, total, percent])

        # Empty row between players
        final_rows.append([""] * 5)
        final_rows.append([""] * 5)
        final_rows.append([""] * 5)

    # Convert to DataFrame and save
    output_df = pd.DataFrame(final_rows)
    out_path = os.path.join(os.path.dirname(file_path), "tennis_statistics_block.xlsx")
    output_df.to_excel(out_path, index=False, header=False)
    return out_path

# GUI Functions
def browse_tennis():
    path = filedialog.askopenfilename(title="Select Tennis_Men.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        tennis_path_var.set(path)

def run_process():
    tennis_path = tennis_path_var.get()
    if not tennis_path:
        messagebox.showerror("Error", "Please select the Tennis_Men.xlsx file.")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        output_path = process_tennis_statistics(tennis_path)
        result_label.config(text=f"Complete!\nResult saved:\n{output_path}", fg="green")
    except Exception as e:
        print(str(e))
        result_label.config(text=f"Error occurred: {str(e)}", fg="red")

# GUI Layout
root = tk.Tk()
root.title("Tennis Stat Generator")
root.geometry("500x220")

tennis_path_var = tk.StringVar()

tk.Button(root, text="Select Tennis_Men.xlsx", command=browse_tennis).pack(pady=(15, 0))
tk.Label(root, textvariable=tennis_path_var, wraplength=480).pack()

tk.Button(root, text="Generate Stats", command=run_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
