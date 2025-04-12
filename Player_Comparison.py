import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from itertools import combinations
import openpyxl

def process_file(input_path):
    # Read Excel file without header
    df = pd.read_excel(input_path, header=None)

    try:
        df = df.iloc[:, [0, 7, 15]]  # Columns A, H, P
    except IndexError:
        raise ValueError("The Excel file does not contain columns A, H, P (0, 7, 15).")

    df.columns = ['Player', 'Result', 'Day']
    df.dropna(subset=['Player', 'Result', 'Day'], inplace=True)

    pair_stats = {}

    for day, group in df.groupby('Day'):
        players_today = group[['Player', 'Result']].values

        for (p1, r1), (p2, r2) in combinations(players_today, 2):
            if p1 == p2:
                continue

            key = tuple(sorted([p1, p2]))

            if key not in pair_stats:
                pair_stats[key] = {'days': 0, 'oo': 0, 'uu': 0, 'ou': 0}

            pair_stats[key]['days'] += 1
            if r1 == 'OVER' and r2 == 'OVER':
                pair_stats[key]['oo'] += 1
            elif r1 == 'UNDER' and r2 == 'UNDER':
                pair_stats[key]['uu'] += 1
            else:
                pair_stats[key]['ou'] += 1

    rows = []
    for (p1, p2), stats in pair_stats.items():
        days = stats['days']
        oo = stats['oo']
        uu = stats['uu']
        ou = stats['ou']
        row = [
            p1, p2, days, oo, uu, ou,
            round(oo / days, 2),
            round(uu / days, 2),
            round(ou / days, 2)
        ]
        rows.append(row)

    result_df = pd.DataFrame(rows, columns=[
        'Player A', 'Player B', 'Days Matched',
        'OVER/OVER', 'UNDER/UNDER', 'OVER/UNDER',
        'OVER/OVER%', 'UNDER/UNDER%', 'OVER/UNDER%'
    ])

    out_path = os.path.join(os.path.dirname(input_path), "Player_Comparison.xlsx")
    result_df.to_excel(out_path, index=False)

     # Auto-adjust column width
    wb = openpyxl.load_workbook(out_path)
    ws = wb.active

    for col in ws.columns:
        header_value = col[0].value
        col_letter = col[0].column_letter

        if not header_value or str(header_value).strip() == "":
            ws.column_dimensions[col_letter].width = 10
            continue

        max_length = 0
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(out_path)

    result_label.config(text="Conversion completed! → Saved as Player_Comparison.xlsx", fg="green")

# --- GUI Logic ---
def browse_file():
    path = filedialog.askopenfilename(
        title="Select Excel File (No Header)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if path:
        input_path_var.set(path)

def run_conversion():
    path = input_path_var.get()
    if not path:
        result_label.config(text="Please select a file first.", fg="red")
        return
    result_label.config(text="Processing...", fg="blue")
    root.update()
    try:
        process_file(path)
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}", fg="red")

# --- GUI Layout ---
root = tk.Tk()
root.title("Player Pair Over/Under Matcher")
root.geometry("500x250")

input_path_var = tk.StringVar()

tk.Button(root, text="Select Excel File", command=browse_file).pack(pady=(20, 5))
tk.Label(root, textvariable=input_path_var, wraplength=480).pack()

tk.Button(root, text="Run Conversion", command=run_conversion, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
