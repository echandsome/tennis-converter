import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

# Month name mapping
MONTH_MAP = {
    '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr',
    '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug',
    '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
}

def format_date(date_str):
    if pd.isna(date_str):
        return ['', '', '']
    try:
        date = pd.to_datetime(date_str)
        year = str(date.year)
        month = MONTH_MAP[str(date.month).zfill(2)]
        day = str(date.day).zfill(2)
        return [day, month, year]
    except:
        return ['', '', '']

def get_player_data(players_df, player_name):
    player_row = players_df[players_df['Player'] == player_name]
    if not player_row.empty:
        return player_row.iloc[0, 3], player_row.iloc[0, 4], player_row.iloc[0, 5], player_row.iloc[0, 1]
    else:
        return '', '', '', ''

def process_files(matches_path, players_path):
    matches_df = pd.read_excel(matches_path, header=0)
    players_df = pd.read_excel(players_path, header=0)

    win_rows = []
    lose_rows = []

    for i in range(1, len(matches_df), 3):
        try:
            lose_row = matches_df.iloc[i - 1]
            win_row = matches_df.iloc[i]

            lose_date = format_date(lose_row.iloc[0])
            win_date = format_date(win_row.iloc[0])

            lose_location = lose_row.iloc[3] if not pd.isna(lose_row.iloc[3]) else ''
            win_location = win_row.iloc[3] if not pd.isna(win_row.iloc[3]) else ''

            lose_player_name = lose_row.iloc[2]
            win_player_name = win_row.iloc[2]

            l_player_h, l_player_i, l_player_j, l_player_k = get_player_data(players_df, lose_player_name)
            w_player_h, w_player_i, w_player_j, w_player_k = get_player_data(players_df, win_player_name)

            lose_output = lose_date + [lose_location] + [''] * 3 + [l_player_h, l_player_i, l_player_j, l_player_k]
            win_output = win_date + [win_location] + [''] * 3 + [w_player_h, w_player_i, w_player_j, w_player_k]

            win_rows.append(win_output)
            lose_rows.append(lose_output)
        except IndexError:
            continue

    header = ['Day', 'Month', 'Year', 'Location', '', '', '',
              'H', 'I', 'J', 'Birth Place']

    out_dir = os.path.dirname(matches_path)
    pd.DataFrame([header] + win_rows).to_csv(os.path.join(out_dir, "Win_Astro.csv"), index=False, header=False)
    pd.DataFrame([header] + lose_rows).to_csv(os.path.join(out_dir, "Lose_Astro.csv"), index=False, header=False)

    result_label.config(text="Conversion complete!\nFiles saved next to input.", fg="green")

# GUI logic
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
        result_label.config(text=f"Error: {str(e)}", fg="red")

# --- GUI SETUP ---
root = tk.Tk()
root.title("Match File Converter")
root.geometry("500x250")

matches_path_var = tk.StringVar()
players_path_var = tk.StringVar()

tk.Button(root, text="Select Match File", command=browse_matches).pack(pady=(10, 0))
tk.Label(root, textvariable=matches_path_var, wraplength=580).pack()

tk.Button(root, text="Select Player File", command=browse_players).pack(pady=(10, 0))
tk.Label(root, textvariable=players_path_var, wraplength=580).pack()

tk.Button(root, text="Convert", command=run_conversion, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
