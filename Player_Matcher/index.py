import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def process_and_fill(matches_path, players_path):
    matches_df = pd.read_excel(matches_path, header=None)
    players_df = pd.read_excel(players_path, header=None)

    result_df = matches_df.copy()

    # Indexes
    match_name_index = 2  # Column C
    player_name_index = 0  # Column A in players file
    player_info_indices = [3, 4, 5, 1]  # B, D, E, F in players file
    target_match_indices = [14, 15, 16, 17, 18]  # O to S (but now shifted: leave O blank)

    # Build player info dictionary
    player_info_map = {
        row[player_name_index]: [row[i] if pd.notna(row[i]) else '' for i in player_info_indices]
        for _, row in players_df.iterrows()
    }

    # Fill the match file
    for idx, row in result_df.iterrows():
        player_name = row[match_name_index]
        if player_name in player_info_map:
            result_df.at[idx, target_match_indices[0]] = ''  # Leave first column (O) blank
            for col_idx, info_value in zip(target_match_indices[1:], player_info_map[player_name]):
                result_df.at[idx, col_idx] = info_value
        else:
            for col_idx in target_match_indices:
                result_df.at[idx, col_idx] = ''

    # Save
    out_path = os.path.join(os.path.dirname(matches_path), "matches_filled.xlsx")
    result_df.to_excel(out_path, index=False, header=False)

    result_label.config(text="Done: matches_filled.xlsx created", fg="green")

# --- GUI Functions ---
def browse_matches():
    path = filedialog.askopenfilename(title="Select tennis_matches.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        matches_path_var.set(path)

def browse_players():
    path = filedialog.askopenfilename(title="Select tennis_players.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        players_path_var.set(path)

def run_process():
    matches_path = matches_path_var.get()
    players_path = players_path_var.get()
    if not matches_path or not players_path:
        result_label.config(text="Please select both files.", fg="red")
        return
    result_label.config(text="Processing...", fg="blue")
    root.update()
    try:
        process_and_fill(matches_path, players_path)
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}", fg="red")

# --- GUI Layout ---
root = tk.Tk()
root.title("Match Player Info Filler")
root.geometry("520x260")

matches_path_var = tk.StringVar()
players_path_var = tk.StringVar()

tk.Button(root, text="Select Match File", command=browse_matches).pack(pady=(10, 0))
tk.Label(root, textvariable=matches_path_var, wraplength=480).pack()

tk.Button(root, text="Select Player File", command=browse_players).pack(pady=(10, 0))
tk.Label(root, textvariable=players_path_var, wraplength=480).pack()

tk.Button(root, text="Run Fill Process", command=run_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
