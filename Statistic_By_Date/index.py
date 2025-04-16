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

        result.setdefault(date, {}).setdefault(name, {"OVER": 0, "UNDER": 0})
        if result_val == "OVER":
            result[date][name]["OVER"] += 1
        elif result_val == "UNDER":
            result[date][name]["UNDER"] += 1

    # Sorted date and player name lists
    all_dates = sorted(result.keys())
    all_names = sorted({name for players in result.values() for name in players})

    # Create empty DataFrame with required dimensions
    rows = len(all_dates) + 3  # 2 header rows + data rows + "all dates" row
    cols = 1 + len(all_names) * 4
    output_df = pd.DataFrame("", index=range(rows), columns=range(cols))

    # Write headers
    for name_idx, name in enumerate(all_names):
        base_col = 1 + name_idx * 4
        output_df.iat[0, base_col] = "Win/Over Count"
        output_df.iat[1, base_col] = name

        output_df.iat[0, base_col + 1] = "Lose/Under Count"
        output_df.iat[1, base_col + 1] = name

        output_df.iat[0, base_col + 2] = "Total Win/Lose"
        output_df.iat[1, base_col + 2] = name

        output_df.iat[0, base_col + 3] = "Win/Over %"
        output_df.iat[1, base_col + 3] = name

    # Fill in statistics by date
    for idx, date in enumerate(all_dates):
        row_idx = idx + 2
        output_df.iat[row_idx, 0] = date
        for name_idx, name in enumerate(all_names):
            base_col = 1 + name_idx * 4
            stats = result[date].get(name, {"OVER": 0, "UNDER": 0})
            win = stats["OVER"]
            lose = stats["UNDER"]
            total = win + lose
            percent = f"{(win / total * 100):.0f}%" if total > 0 else ""

            output_df.iat[row_idx, base_col] = win
            output_df.iat[row_idx, base_col + 1] = lose
            output_df.iat[row_idx, base_col + 2] = total
            output_df.iat[row_idx, base_col + 3] = percent

    # Save to Excel
    out_path = os.path.join(os.path.dirname(file_path), "tennis_statistics.xlsx")
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
