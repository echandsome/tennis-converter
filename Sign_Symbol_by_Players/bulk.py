import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def process_files(daily_path, historical_path, col_choice):
    daily_df = pd.read_excel(daily_path, header=None)
    hist_df = pd.read_excel(historical_path, header=None)

    # Add columns E, F, G (index 4, 5, 6) if they don't exist
    for col in [4, 5, 6]:
        if col >= daily_df.shape[1]:
            daily_df[col] = ""

    group_start = None
    group_data = None

    for i in range(len(daily_df)):
        cell = daily_df.iat[i, 0]

        if pd.isna(cell):
            # End of group
            if group_start is not None:
                over_total = 0
                under_total = 0

                for j in range(group_start + 1, i):
                    player = daily_df.iat[j, 0]
                    match_value = daily_df.iat[j, 13 if col_choice == 'N' else 11]

                    matched = hist_df[(hist_df[0] == player) & (hist_df[13 if col_choice == 'N' else 11] == match_value)]
                    over_count = (matched[7] == "OVER").sum()
                    under_count = (matched[7] == "UNDER").sum()

                    total = over_count + under_count
                    percent = f"{round(over_count / total * 100)}%" if total > 0 else ""

                    daily_df.iat[j, 4] = over_count
                    daily_df.iat[j, 5] = under_count
                    daily_df.iat[j, 6] = percent

                    over_total += over_count
                    under_total += under_count

                # Write group summary (AVG) in the empty row
                total_all = over_total + under_total
                percent_all = f"{round(over_total / total_all * 100)}%" if total_all > 0 else ""

                daily_df.iat[i, 4] = over_total
                daily_df.iat[i, 5] = under_total
                daily_df.iat[i, 6] = percent_all

                if group_data:
                    daily_df.iat[i, 7] = group_data[4]

                group_data = None
                group_start = None
        else:
            # Group header
            if isinstance(cell, str) and cell.startswith("(") and cell.endswith(")"):
                group_start = i
                if i + 1 < len(daily_df):
                    
                    group_data = daily_df.iloc[i + 1, 14:18].tolist()  # O~R 열 from the next row
                else:
                    group_data = ["", "", "", ""]
                group_data.append(daily_df.iat[i + 1, 7])

    # Handle last group if missing trailing empty row
    if group_start is not None:
        i = len(daily_df)
        over_total = 0
        under_total = 0

        for j in range(group_start + 1, i):
            player = daily_df.iat[j, 0]
            match_value = daily_df.iat[j, 13 if col_choice == 'N' else 11]

            matched = hist_df[(hist_df[0] == player) & (hist_df[13 if col_choice == 'N' else 11] == match_value)]
            over_count = (matched[7] == "OVER").sum()
            under_count = (matched[7] == "UNDER").sum()

            total = over_count + under_count
            percent = f"{round(over_count / total * 100)}%" if total > 0 else ""

            daily_df.iat[j, 4] = over_count
            daily_df.iat[j, 5] = under_count
            daily_df.iat[j, 6] = percent

            over_total += over_count
            under_total += under_count

        total_all = over_total + under_total
        percent_all = f"{round(over_total / total_all * 100)}%" if total_all > 0 else ""

        # Append empty row with AVG
        daily_df.loc[len(daily_df)] = [""] * daily_df.shape[1]
        daily_df.iat[i, 4] = over_total
        daily_df.iat[i, 5] = under_total
        daily_df.iat[len(daily_df) - 1, 6] = percent_all

        if group_data:
            daily_df.iat[i, 7] = group_data[4]
            daily_df.iat[i, 14] = group_data[0]
            daily_df.iat[i, 15] = group_data[1]
            daily_df.iat[i, 16] = group_data[2]
            daily_df.iat[i, 17] = group_data[3]

    # Save result
    output_path = os.path.join(os.path.dirname(daily_path), "Daily_with_stats.xlsx")
    daily_df.to_excel(output_path, index=False, header=False)
    return output_path

def process_bulk_files(daily_folder, historical_folder, col_choice):
    parent_folder = os.path.dirname(daily_folder)
    output_folder = os.path.join(parent_folder, "Output")
    os.makedirs(output_folder, exist_ok=True)

    # Get all the filenames with prefix '01_', '02_', '03_' ...
    daily_files = sorted([f for f in os.listdir(daily_folder) if f.endswith('.xlsx') and '_' in f])
    historical_files = sorted([f for f in os.listdir(historical_folder) if f.endswith('.xlsx') and '_' in f])

    if len(daily_files) != len(historical_files):
        raise ValueError("The number of Daily and Historical files must match.")

    # Process each pair of files
    for i in range(len(daily_files)):
        daily_path = os.path.join(daily_folder, daily_files[i])
        historical_path = os.path.join(historical_folder, historical_files[i])

        output_path = process_files(daily_path, historical_path, col_choice)

        # Rename the output file with 'Result_' prefix
        base_name = os.path.basename(daily_files[i])
        result_name = f"Result_{base_name}"
        
        # Create output folder in the same directory as input directories
        result_folder = os.path.dirname(daily_folder)
        result_path = os.path.join(output_folder, result_name)

        os.rename(output_path, result_path)

    return output_folder

# GUI functions
def browse_daily_folder():
    folder = filedialog.askdirectory(title="Select Folder with Daily Files")
    if folder:
        daily_folder_var.set(folder)

def browse_historical_folder():
    folder = filedialog.askdirectory(title="Select Folder with Historical Files")
    if folder:
        historical_folder_var.set(folder)

def run_bulk_process():
    daily_folder = daily_folder_var.get()
    historical_folder = historical_folder_var.get()
    col_choice = col_var.get()

    if not daily_folder or not historical_folder:
        messagebox.showerror("Error", "Please select both Daily and Historical folders.")
        return

    if col_choice not in ['N', 'L']:
        messagebox.showerror("Error", "Please select N or L column.")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        output_path = process_bulk_files(daily_folder, historical_folder, col_choice)
        result_label.config(text=f"Complete! Saved to: " + output_path, fg="green")
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}", fg="red")

# GUI setup
root = tk.Tk()
root.title("Bulk Daily vs Historical Analyzer")
root.geometry("600x350")

daily_folder_var = tk.StringVar()
historical_folder_var = tk.StringVar()
col_var = tk.StringVar(value='N')

tk.Button(root, text="Select Folder with Daily Files", command=browse_daily_folder).pack(pady=(10, 0))
tk.Label(root, textvariable=daily_folder_var, wraplength=580).pack()

tk.Button(root, text="Select Folder with Historical Files", command=browse_historical_folder).pack(pady=(10, 0))
tk.Label(root, textvariable=historical_folder_var, wraplength=580).pack()

# Radio buttons for column selection
frame = tk.Frame(root)
frame.pack(pady=10)
tk.Label(frame, text="Choose column for matching:").pack(side=tk.LEFT)
tk.Radiobutton(frame, text="N", variable=col_var, value='N').pack(side=tk.LEFT, padx=5)
tk.Radiobutton(frame, text="L", variable=col_var, value='L').pack(side=tk.LEFT, padx=5)

tk.Button(root, text="Run Bulk Process", command=run_bulk_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
