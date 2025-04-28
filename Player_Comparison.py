import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from itertools import combinations
import openpyxl

def process_file(input_path, oo_threshold, uu_threshold, group_option):
    # Read Excel file without header
    df = pd.read_excel(input_path, header=None)

    try:
        df = df.iloc[:, [0, 7, 11, 13, 15]]  # Columns A, H, L, N, P
    except IndexError:
        raise ValueError("The Excel file does not contain columns A, H, L, N, P.")

    df.columns = ['Player', 'Result', 'L', 'N', 'Day']
    df.dropna(subset=['Player', 'Result', 'Day'], inplace=True)

    # Normalize values
    df['Result'] = df['Result'].str.upper().str.strip()
    result_map = {'WIN': 'OVER', 'LOSE': 'UNDER'}
    df['Result'] = df['Result'].replace(result_map)

    # Determine group keys
    if group_option == 'L':
        group_keys = ['Day', 'L']
        group_label = 'L'
    elif group_option == 'N':
        group_keys = ['Day', 'N']
        group_label = 'N'
    else:
        group_keys = ['Day']
        group_label = None

    pair_stats = {}

    for keys, group in df.groupby(group_keys):
        players_today = group[['Player', 'Result']].values

        if isinstance(keys, tuple) and len(keys) > 1:
            group_val = keys[1]
        else:
            group_val = None

        for (p1, r1), (p2, r2) in combinations(players_today, 2):
            if p1 == p2:
                continue

            if group_val is not None:
                key = tuple(sorted([p1, p2]) + [str(group_val)])
            else:
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
    for key, stats in pair_stats.items():
        if group_label:
            p1, p2, group_val = key
            row = [p1, p2, group_val]
        else:
            p1, p2 = key
            row = [p1, p2]

        days = stats['days']
        oo = stats['oo']
        uu = stats['uu']
        ou = stats['ou']
        row.extend([
            days, oo, uu, ou,
            round(oo / days, 2),
            round(uu / days, 2),
            round(ou / days, 2)
        ])
        rows.append(row)

    if group_label:
        columns = [
            'Player A', 'Player B', f'{group_label} Value',
            'Days Matched', 'OVER/OVER', 'UNDER/UNDER', 'OVER/UNDER',
            'OVER/OVER%', 'UNDER/UNDER%', 'OVER/UNDER%'
        ]
    else:
        columns = [
            'Player A', 'Player B',
            'Days Matched', 'OVER/OVER', 'UNDER/UNDER', 'OVER/UNDER',
            'OVER/OVER%', 'UNDER/UNDER%', 'OVER/UNDER%'
        ]

    result_df = pd.DataFrame(rows, columns=columns)

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    folder = os.path.dirname(input_path)

    all_path = os.path.join(folder, f"ALL_{base_name}.xlsx")
    if oo_threshold and oo_threshold != 0:
        oo_path = os.path.join(folder, f"OO_{oo_threshold}_{base_name}.xlsx")
    else:
        oo_path = None
    if uu_threshold and uu_threshold != 0:
        uu_path = os.path.join(folder, f"UU_{uu_threshold}_{base_name}.xlsx")
    else:
        uu_path = None

    result_df.to_excel(all_path, index=False)
    if oo_path:
        result_df[result_df['OVER/OVER%'] >= (oo_threshold/100)].to_excel(oo_path, index=False)
    if uu_path:
        result_df[result_df['UNDER/UNDER%'] >= (uu_threshold/100)].to_excel(uu_path, index=False)

    for path in [all_path, oo_path, uu_path]:
        if path:
            auto_adjust_excel(path)

    result_label.config(
        text=f"Saved:\nALL → {os.path.basename(all_path)}\n"
             f"{'OO → ' + os.path.basename(oo_path) if oo_path else ''}\n"
             f"{'UU → ' + os.path.basename(uu_path) if uu_path else ''}",
        fg="green"
    )

def auto_adjust_excel(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    for col in ws.columns:
        header_value = col[0].value
        col_letter = col[0].column_letter

        max_length = 0
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(path)

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
    try:
        oo_threshold = float(oo_entry.get())
        uu_threshold = float(uu_entry.get())
    except ValueError:
        result_label.config(text="Please enter valid percentages (numbers).", fg="red")
        return

    if not path:
        result_label.config(text="Please select a file first.", fg="red")
        return

    group_option = group_option_var.get()
    result_label.config(text="Processing...", fg="blue")
    root.update()
    try:
        process_file(path, oo_threshold, uu_threshold, group_option)
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}", fg="red")

# --- GUI Layout ---
root = tk.Tk()
root.title("Player Pair Over/Under Matcher")
root.geometry("500x350")

input_path_var = tk.StringVar()
group_option_var = tk.StringVar(value="None")

tk.Button(root, text="Select Excel File", command=browse_file).pack(pady=(15, 5))
tk.Label(root, textvariable=input_path_var, wraplength=480).pack()

entry_frame = tk.Frame(root)
entry_frame.pack(pady=10)
tk.Label(entry_frame, text="OVER/OVER% ≥").grid(row=0, column=0, padx=5)
oo_entry = tk.Entry(entry_frame, width=5)
oo_entry.insert(0, "")
oo_entry.grid(row=0, column=1, padx=5)
tk.Label(entry_frame, text="UNDER/UNDER% ≥").grid(row=0, column=2, padx=5)
uu_entry = tk.Entry(entry_frame, width=5)
uu_entry.insert(0, "")
uu_entry.grid(row=0, column=3, padx=5)

# Group Option Radio Buttons
group_frame = tk.LabelFrame(root, text="Group By")
group_frame.pack(pady=10)
for text in ["None", "L", "N"]:
    tk.Radiobutton(group_frame, text=text, variable=group_option_var, value=text).pack(side="left", padx=10)

tk.Button(root, text="Run Conversion", command=run_conversion, bg="green", fg="white").pack(pady=10)
result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
