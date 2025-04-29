import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import os

def split_by_dates(file_path, num_dates, output_text):
    try:
        df = pd.read_excel(file_path)

        # Always use the 16th column (index 15)
        if len(df.columns) < 16:
            output_text.insert(tk.END, "Error: Not enough columns in the file!\n")
            return

        # Get the 16th column, remove parentheses, and try to convert to MM-DD-YYYY format first
        date_series = df.iloc[:, 15].astype(str).str.replace(r'[()]', '', regex=True)

        # First, try to parse MM-DD-YYYY format
        df['DATE_COL'] = pd.to_datetime(date_series, format='%m-%d-%Y', errors='coerce')

        # If there are NaT values (failed parsing), try parsing MM-DD format
        df['DATE_COL'] = df['DATE_COL'].fillna(pd.to_datetime(date_series, format='%m-%d', errors='coerce'))

        # Remove rows without valid dates
        df = df.dropna(subset=['DATE_COL'])

        # Sort dates as MMDD format strings
        unique_dates = sorted(df['DATE_COL'].dt.strftime('%m%d').unique(), reverse=True)

        output_dir = os.path.join(os.path.dirname(file_path), 'split_output')
        os.makedirs(output_dir, exist_ok=True)

        file_counter = 1
        for i in range(0, len(unique_dates) - num_dates + 1):
            selected_dates = unique_dates[i:i+num_dates]

            subset = df[df['DATE_COL'].dt.strftime('%m%d').isin(selected_dates)]
            if subset.empty:
                continue

            start_date = selected_dates[-1]
            end_date = selected_dates[0]

            file_name = f"File_{file_counter:03d}_{start_date}_{end_date}.xlsx"
            output_path = os.path.join(output_dir, file_name)

            # --- Save file ---
            data_to_save = subset.drop(columns=['DATE_COL'])

            # Create custom header
            header1 = ['Partner A', '', 'Partner B', '', '', '', '', '', '', '', '', 'Signs', 'Symbol', 'Signs-Symbol', '', '', '', '']
            header2 = ['Partner A', '', 'Partner B', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                pd.DataFrame([header1, header2]).to_excel(writer, header=False, index=False)
                data_to_save.to_excel(writer, startrow=2, index=False, header=False)

            output_text.insert(tk.END, f"Saved: {file_name}\n")
            file_counter += 1

        if file_counter == 1:
            output_text.insert(tk.END, "No files were created.\n")
        else:
            output_text.insert(tk.END, f"\nDone! Files saved in {output_dir}\n")

    except Exception as e:
        output_text.insert(tk.END, f"Error: {str(e)}\n")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)

def start_split():
    file_path = entry_file.get()
    num_dates = entry_num.get()
    output_text.delete(1.0, tk.END)

    if not file_path or not os.path.exists(file_path):
        messagebox.showerror("Error", "Please select a valid Excel file.")
        return
    try:
        num_dates = int(num_dates)
        if not (1 <= num_dates <= 100):
            raise ValueError
    except ValueError:
        messagebox.showerror("Error", "Enter a valid number between 1 and 100.")
        return

    split_by_dates(file_path, num_dates, output_text)

# GUI setup
root = tk.Tk()
root.title("Date Group Splitter")
root.geometry("550x450")

# File selection frame
frame_file = tk.Frame(root)
frame_file.pack(pady=5)

btn_file = tk.Button(frame_file, text="Select Excel File", command=select_file)
btn_file.pack(side=tk.LEFT, padx=5)

entry_file = tk.Entry(frame_file, width=50)
entry_file.pack(side=tk.LEFT, padx=5)

# Number input frame
frame_num = tk.Frame(root)
frame_num.pack(pady=5)

lbl_num = tk.Label(frame_num, text="Number of dates per group:")
lbl_num.pack(side=tk.LEFT, padx=5)

entry_num = tk.Entry(frame_num, width=5)
entry_num.pack(side=tk.LEFT, padx=5)

# Start button
btn_start = tk.Button(root, text="Start Splitting", command=start_split)
btn_start.pack(pady=10)

# Result output window
output_text = scrolledtext.ScrolledText(root, width=65, height=20)
output_text.pack(pady=5)

root.mainloop()
