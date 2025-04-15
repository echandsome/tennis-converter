import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def browse_excel_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    excel_entry.delete(0, tk.END)
    excel_entry.insert(0, filename)

def browse_condition_file():
    filename = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    condition_entry.delete(0, tk.END)
    condition_entry.insert(0, filename)

def process_bulk():
    result_label.config(text="Processing...", fg="blue")

    excel_path = excel_entry.get()
    condition_path = condition_entry.get()
    
    if not excel_path or not condition_path:
        result_label.config(text="Please select both Excel and condition CSV files.")
        return

    try:
        output_dir = os.path.join(os.path.dirname(excel_path), "outputs")
        os.makedirs(output_dir, exist_ok=True)

        df = pd.read_excel(excel_path, engine="openpyxl", header=None)
        column_n = df.columns[13]  # Column N
        column_q = df.columns[16]  # Column Q
        column_h = df.columns[7]   # Column H

        conditions = pd.read_csv(condition_path)
        all_over_results = []
        all_under_results = []

        for _, row in conditions.iterrows():
            group_size = int(row["Group Over/Under"])
            percent = float(row["% OVER/UNDER"])

            # Filter groups based on minimum group size
            grouped = df.groupby([column_n, column_q]).size().reset_index(name="count")
            valid_groups = grouped[grouped['count'] >= group_size][[column_n, column_q]]
            df_filtered = df.merge(valid_groups, on=[column_n, column_q], how='inner')

            # Process OVER
            over_rows = []
            for (n, q), group in df_filtered.groupby([column_n, column_q]):
                group = group.sort_index(ascending=False)
                latest = group.head(group_size)
                over_pct = (latest[column_h] == "OVER").sum() / group_size * 100
                if over_pct >= percent:
                    over_rows.append(latest)
            if over_rows:
                df_over = pd.concat(over_rows)
                if percent.is_integer():
                    percent_str = str(int(percent))
                else:
                    percent_str = f"{percent:.2f}".replace(".", "p")
                over_filename = os.path.join(output_dir, f"OVER_{group_size}_{percent_str}.csv")
                df_over.to_csv(over_filename, index=False, header=False)
                all_over_results.append(df_over)

            # Process UNDER
            under_rows = []
            for (n, q), group in df_filtered.groupby([column_n, column_q]):
                group = group.sort_index(ascending=False)
                latest = group.head(group_size)
                under_pct = (latest[column_h] == "UNDER").sum() / group_size * 100
                if under_pct >= percent:
                    under_rows.append(latest)
            if under_rows:
                df_under = pd.concat(under_rows)
                if percent.is_integer():
                    percent_str = str(int(percent))
                else:
                    percent_str = f"{percent:.2f}".replace(".", "p")
                under_filename = os.path.join(output_dir, f"UNDER_{group_size}_{percent_str}.csv")
                df_under.to_csv(under_filename, index=False, header=False)
                all_under_results.append(df_under)

        # Save merged results
        if all_over_results:
            over_all = pd.concat(all_over_results).drop_duplicates()
            over_all.to_csv(os.path.join(output_dir, "OVER_ALL.csv"), index=False, header=False)

        if all_under_results:
            under_all = pd.concat(all_under_results).drop_duplicates()
            under_all.to_csv(os.path.join(output_dir, "UNDER_ALL.csv"), index=False, header=False)

        if all_over_results or all_under_results:
            overunder_all = pd.concat(all_over_results + all_under_results).drop_duplicates()
            overunder_all.to_csv(os.path.join(output_dir, "OVERUNDER_ALL.csv"), index=False, header=False)

        result_label.config(text="Processing complete! Check 'outputs' folder.", fg="green")
    
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}")

# Build GUI
root = tk.Tk()
root.title("Bulk OVER/UNDER Processor")
root.geometry("500x250")

tk.Label(root, text="Select Excel File (.xlsx):").pack()
excel_entry = tk.Entry(root, width=50)
excel_entry.pack()
tk.Button(root, text="Browse", command=browse_excel_file).pack()

tk.Label(root, text="Select Condition CSV File:").pack()
condition_entry = tk.Entry(root, width=50)
condition_entry.pack()
tk.Button(root, text="Browse", command=browse_condition_file).pack()

tk.Button(root, text="Run Bulk Processing", command=process_bulk, bg="#4CAF50", fg="white").pack(pady=10)

result_label = tk.Label(root, text="", fg="blue")
result_label.pack()

root.mainloop()
