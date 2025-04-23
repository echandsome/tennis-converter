import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def determine_winner(row1, row2):
    player1_sets_won = 0
    player2_sets_won = 0

    for col in range(4, 11):  # E~K, index 4~10
        val1 = row1.iloc[col]
        val2 = row2.iloc[col]

        if pd.isna(val1) or pd.isna(val2):
            continue

        try:
            val1 = int(val1)
            val2 = int(val2)
        except:
            continue

        if val1 > val2:
            player1_sets_won += 1
        elif val2 > val1:
            player2_sets_won += 1

    if player1_sets_won > player2_sets_won:
        return "W", "L"
    elif player2_sets_won > player1_sets_won:
        return "L", "W"
    else:
        return "", ""

def process_file(filepath):
    df = pd.read_excel(filepath)

    match_blocks = []
    current_block = []

    for idx in range(len(df)):
        row = df.iloc[idx]

        if row.isnull().all():
            if len(current_block) == 2:
                match_blocks.append((current_block[0], current_block[1]))
            current_block = []

            # If there's no valid data from the next row onwards, end the loop
            next_rows = df.iloc[idx+1:]
            if next_rows.dropna(how='all').empty:
                break
        else:
            current_block.append((idx, row))

    # Prevent block omission after loop
    if len(current_block) == 2:
        match_blocks.append((current_block[0], current_block[1]))

    # Ensure columns: There should be at least 19 columns
    while df.shape[1] < 19:
        df[f"Empty_{df.shape[1]}"] = ""

    # Set N~R (index 13~17) headers to empty string
    for col_idx in range(13, 19):
        df.columns.values[col_idx] = ""

    # Calculate win/loss
    for (idx1, row1), (idx2, row2) in match_blocks:
        result1, result2 = determine_winner(row1, row2)
        df.iat[idx1, 18] = result1
        df.iat[idx2, 18] = result2

    # Save
    base, ext = os.path.splitext(filepath)
    output_path = f"{base}_with_results.xlsx"
    df.to_excel(output_path, index=False)

    result_label.config(text=f"Completed!\nSaved at: {output_path}", fg="green")

# GUI
def browse_file():
    path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
    if path:
        file_path_var.set(path)

def run_conversion():
    filepath = file_path_var.get()
    if not filepath:
        result_label.config(text="Please select a file first.", fg="red")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()
    try:
        process_file(filepath)
    except Exception as e:
        result_label.config(text=f"Error occurred: {str(e)}", fg="red")

# --- GUI SETUP ---
root = tk.Tk()
root.title("Tennis Match Win/Loss Analyzer")
root.geometry("500x250")

file_path_var = tk.StringVar()

tk.Button(root, text="Select Excel File", command=browse_file).pack(pady=(10, 0))
tk.Label(root, textvariable=file_path_var, wraplength=480).pack()

tk.Button(root, text="Start Win/Loss Analysis", command=run_conversion, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
