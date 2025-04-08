import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime
import platform
import openpyxl

# Date format function (for Windows/Unix compatibility)
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

# Merge and save function
def add_gender_and_format(players_path, players_moon_path):
    players_df = pd.read_excel(players_path)
    moon_df = pd.read_excel(players_moon_path)

    # Explicitly specify column names
    players_df.columns.values[0] = "Player"
    players_df.columns.values[2] = "Date"
    moon_df.columns.values[2] = "Player"
    moon_df.columns.values[5] = "Gender"

    # Apply date format (column C of players_df)
    try:
        players_df['Date'] = pd.to_datetime(players_df['Date'], errors='coerce')
        players_df['Date'] = players_df['Date'].apply(lambda x: convert_date_format(x) if pd.notnull(x) else "")
    except Exception as e:
        raise Exception("Error occurred while converting date format: " + str(e))

    # Merge Gender
    moon_gender_df = moon_df[['Player', 'Gender']]
    merged_df = pd.merge(players_df, moon_gender_df, on='Player', how='left')

    # Save file
    out_dir = os.path.dirname(players_path)
    output_path = os.path.join(out_dir, "players_with_gender.xlsx")
    merged_df.to_excel(output_path, index=False)

    # Auto-adjust column width
    wb = openpyxl.load_workbook(output_path)
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

    wb.save(output_path)

    result_label.config(text=f"Completed! File saved:\n{output_path}", fg="green")

# GUI configuration
def browse_players():
    path = filedialog.askopenfilename(title="Select players.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        players_path_var.set(path)

def browse_players_moon():
    path = filedialog.askopenfilename(title="Select players_moon.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        players_moon_path_var.set(path)

def run_process():
    players_path = players_path_var.get()
    players_moon_path = players_moon_path_var.get()

    if not players_path or not players_moon_path:
        result_label.config(text="Please select both files.", fg="red")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        add_gender_and_format(players_path, players_moon_path)
    except Exception as e:
        result_label.config(text=f"Error occurred: {str(e)}", fg="red")

# GUI initialization
root = tk.Tk()
root.title("players.xlsx + Gender Merger")
root.geometry("500x300")

players_path_var = tk.StringVar()
players_moon_path_var = tk.StringVar()

tk.Button(root, text="Select players.xlsx", command=browse_players).pack(pady=(10, 0))
tk.Label(root, textvariable=players_path_var, wraplength=480).pack()

tk.Button(root, text="Select players_moon.xlsx", command=browse_players_moon).pack(pady=(10, 0))
tk.Label(root, textvariable=players_moon_path_var, wraplength=480).pack()

tk.Button(root, text="Merge and Export as New File", command=run_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
