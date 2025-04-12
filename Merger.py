import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import tempfile

# Function to merge headerless CSV or XLSX files in a folder
def merge_files(folder_path):
    # List all CSV and XLSX files in the selected folder
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.csv', '.xlsx'))]
    if not files:
        result_label.config(text="No CSV or XLSX files found.", fg="red")
        return

    # Check extension of the first file to determine the file type
    first_file = files[0]
    ext = os.path.splitext(first_file)[1].lower()

    if ext == ".csv":
        # If files are CSV, merge them directly
        csv_files = [f for f in files if f.lower().endswith(".csv")]
        dfs = [pd.read_csv(os.path.join(folder_path, f), header=None) for f in csv_files]
        merged_df = pd.concat(dfs, ignore_index=True)

        parent_folder = os.path.dirname(folder_path)
        output_path = os.path.join(parent_folder, "Merged_Output.csv")
        merged_df.to_csv(output_path, index=False, header=False)
        result_label.config(text=f"✔ CSV merge complete!\nSaved as: {output_path}", fg="green")

    elif ext == ".xlsx":
        # If files are XLSX, convert each to temporary CSV first
        xlsx_files = [f for f in files if f.lower().endswith(".xlsx")]
        temp_csv_paths = []

        for f in xlsx_files:
            df = pd.read_excel(os.path.join(folder_path, f), header=None)
            temp_fd, temp_path = tempfile.mkstemp(suffix=".csv")
            os.close(temp_fd)  # Close the temp file
            df.to_csv(temp_path, index=False, header=False)
            temp_csv_paths.append(temp_path)

        # Merge all temp CSV files
        dfs = [pd.read_csv(f, header=None) for f in temp_csv_paths]
        merged_df = pd.concat(dfs, ignore_index=True)

        parent_folder = os.path.dirname(folder_path)
        output_path = os.path.join(parent_folder, "Merged_Output.xlsx")
        
        merged_df.to_excel(output_path, index=False, header=False)
        result_label.config(text=f"✔ XLSX merge complete!\nSaved as: {output_path}", fg="green")

        # Clean up temp files
        for f in temp_csv_paths:
            os.remove(f)

    else:
        result_label.config(text="❌ Unsupported file format in folder.", fg="red")

# --- GUI logic ---

# Open folder dialog to let user select a folder
def browse_folder():
    path = filedialog.askdirectory(title="Select folder with headerless CSV/XLSX files")
    if path:
        folder_path_var.set(path)

# Run the merge function and handle UI updates
def run_merge():
    folder_path = folder_path_var.get()
    if not folder_path:
        result_label.config(text="Please select a folder first.", fg="red")
        return
    result_label.config(text="Processing...", fg="blue")
    root.update()
    try:
        merge_files(folder_path)
    except Exception as e:
        print(e)
        result_label.config(text=f"Error: {str(e)}", fg="red")

# --- GUI SETUP ---

root = tk.Tk()
root.title("Headerless File Merger")
root.geometry("500x200")

folder_path_var = tk.StringVar()

# Button to browse folders
tk.Button(root, text="Select Folder", command=browse_folder).pack(pady=(15, 0))

# Display selected folder path
tk.Label(root, textvariable=folder_path_var, wraplength=580).pack()

# Button to start merge
tk.Button(root, text="Merge Files", command=run_merge, bg="blue", fg="white").pack(pady=20)

# Label to show results or errors
result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()