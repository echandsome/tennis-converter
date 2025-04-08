import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime
import platform
import openpyxl

def convert_date_format(date_input):
    if isinstance(date_input, pd.Timestamp):
        dt = date_input.to_pydatetime()
    elif isinstance(date_input, datetime):
        dt = date_input
    elif isinstance(date_input, str):
        dt = datetime.strptime(date_input, "%Y-%m-%d %H:%M:%S")
    else:
        raise ValueError("Input must be a string, datetime, or pandas.Timestamp")

    if platform.system() == "Windows":
        return dt.strftime("%#m/%#d/%Y")
    else:
        return dt.strftime("%-m/%-d/%Y")

def process_files(matches_path, players_path):
    matches_df = pd.read_excel(matches_path, header=0)
    players_df = pd.read_excel(players_path, header=0)

    output_rows = []

    for _, match_row in matches_df.iterrows():
        player_name = match_row.iloc[2]  # Column C
        date = match_row.iloc[0]         # Column A
        result = match_row.iloc[18]      # Column S

        try:
            formatted_date = '"' + pd.to_datetime(date).strftime('%m-%d') + '"'
        except:
            formatted_date = '""'

        player_row = players_df[players_df['Player'] == player_name]
        if not player_row.empty:
            birth_place = player_row.iloc[0, 1]
            date_series = pd.to_datetime([player_row.iloc[0, 2]])
            date_of_birth = convert_date_format(date_series[0])
        else:
            birth_place = ''
            date_of_birth = ''

        row = [''] * 75
        row[52] = player_name

        for i in range(53, 59):
            row[i] = 1

        row[59] = formatted_date

        for i in range(60, 62):
            row[i] = 1

        row[62] = birth_place
        row[63] = date_of_birth

        for i in range(64, 74):
            row[i] = 1

        if result == 'W':
            row[74] = 'WIN'
        elif result == 'L':
            row[74] = 'LOSE'

        output_rows.append(row)

    header_temps = ["Player Name", "Number", "Odds", "Projection", "Avg", "Home/Away Avg", "Home/Away", "Date",
                    "Stat Category", "Team", "Place of Birth", "Date of Birth", "Illumination", "Age", "Type",
                    "Moon Cycle", "Result", "H/A DIF", "H/A Results DIF", "H/A Spread Result", "AVG DIF", "H / A DIF", "RESULT O/U"]

    headers = [""] * 52 + header_temps

    output_df = pd.DataFrame(output_rows, columns=headers)
    out_dir = os.path.dirname(matches_path)
    output_file = os.path.join(out_dir, "Tennis_Output_Example.xlsx")
    output_df.to_excel(output_file, index=False)

    # Auto-adjust column widths
    wb = openpyxl.load_workbook(output_file)
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

    wb.save(output_file)

    result_label.config(text="Conversion complete!\nFile saved as 'Tennis_Output_Example.xlsx'", fg="green")

def browse_matches():
    path = filedialog.askopenfilename(title="Select tennis_matches1.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        matches_path_var.set(path)

def browse_players():
    path = filedialog.askopenfilename(title="Select Input_Tennis_Players.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        players_path_var.set(path)

def run_conversion():
    matches_path = matches_path_var.get()
    players_path = players_path_var.get()
    if not matches_path or not players_path:
        result_label.config(text="Please select both files first.", fg="red")
        return
    result_label.config(text="Processing...", fg="blue")
    root.update()
    try:
        process_files(matches_path, players_path)
    except Exception as e:
        print(e)
        result_label.config(text=f"Error: {str(e)}", fg="red")

root = tk.Tk()
root.title("Tennis Match File Converter")
root.geometry("500x270")

matches_path_var = tk.StringVar()
players_path_var = tk.StringVar()

tk.Button(root, text="Select Match File", command=browse_matches).pack(pady=(10, 0))
tk.Label(root, textvariable=matches_path_var, wraplength=480).pack()

tk.Button(root, text="Select Player File", command=browse_players).pack(pady=(10, 0))
tk.Label(root, textvariable=players_path_var, wraplength=480).pack()

tk.Button(root, text="Convert to Tennis_Output_Example.xlsx", command=run_conversion, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
