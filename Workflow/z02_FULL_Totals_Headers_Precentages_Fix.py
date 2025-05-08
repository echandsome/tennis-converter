import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import openpyxl

moon_phases = [
    "first quarter", "full moon", "last quarter", "new moon",
    "waning crescent", "waning gibbous", "waxing crescent", "waxing gibbous"
]

def process_files():
    data_file = file_entry.get()
    symbol_file = symbol_entry.get()

    if not data_file or not symbol_file:
        result_label.config(text="Please select both files.")
        return

    try:
        # Load Excel/CSV file
        ext = os.path.splitext(data_file)[1].lower()
        if ext == ".xlsx":
            df = pd.read_excel(data_file, engine="openpyxl")
        elif ext == ".csv":
            df = pd.read_csv(data_file)
        else:
            result_label.config(text="Unsupported data file format.")
            return

        fc_name = df.columns[4]

        # Prepare columns
        df = df.rename(columns={df.columns[0]: "Symbol"})
        df = df.rename(columns={df.columns[1]: "Symbol Part A"})
        df = df.rename(columns={df.columns[2]: "Symbol Part B"})
        df = df.rename(columns={df.columns[3]: "Phase"})
        df = df.rename(columns={df.columns[4]: "J"})
        df = df.rename(columns={df.columns[5]: "M"})

        num_columns = len(df.columns)

        i = 0
        if num_columns == 10:
            i += 1
        elif num_columns == 11:
            i += 2

        # C and D numeric (assume C is col[4], D is col[5])
        col_c = pd.to_numeric(df.iloc[:, 4 + i], errors='coerce').fillna(0)  # Over
        col_d = pd.to_numeric(df.iloc[:, 5 + i], errors='coerce').fillna(0)  # Under
        df["C"] = col_c
        df["D"] = col_d

        # Read symbol list
        with open(symbol_file, "r", encoding="utf-8") as f:
            symbol_list = [line.strip() for line in f if line.strip()]

        output_rows = []

        for symbol in symbol_list:
            t_over_count = 0
            t_under_count = 0
            t_total = 0

            if num_columns == 10:
                # First group by J values
                for j_value in df[df["Symbol Part A"] == symbol]["J"].unique():
                    for phase in moon_phases:
                        matched = df[(df["Symbol Part A"] == symbol) & (df["Phase"] == phase) & (df["J"] == j_value)]

                        # print(matched["J"].values[0])
                        over_count = matched["C"].sum()
                        under_count = matched["D"].sum()
                        total = over_count + under_count

                        win_over = round(over_count / total, 2) if total > 0 else 0
                        win_under = round(under_count / total, 2) if total > 0 else 0

                        t_over_count += over_count
                        t_under_count += under_count
                        t_total += total

                        output_rows.append({
                            "Symbol": "",
                            "Symbol Part A": symbol,
                            f"{fc_name}": j_value,
                            "Phase": phase,
                            "Over count": over_count,
                            "Under count": under_count,
                            "Total": total,
                            "WIN% OVER": win_over,
                            "WIN% UNDER": win_under
                        })

                    # Add summary row for each J value
                    output_rows.append({
                        "Symbol": f"{symbol}-BLK",
                        "Symbol Part A": "",
                        f"{fc_name}": j_value,
                        "Phase": "",
                        "Over count": t_over_count,
                        "Under count": t_under_count,
                        "Total": t_total,
                        "WIN% OVER": round(t_over_count / t_total, 2) if t_total > 0 else 0,
                        "WIN% UNDER": round(t_under_count / t_total, 2) if t_total > 0 else 0
                    })
                    
                    # Reset counters for next J value
                    t_over_count = 0
                    t_under_count = 0
                    t_total = 0

            elif num_columns == 11:
                 # First group by J and M values
                for j_value in df[df["Symbol Part A"] == symbol]["J"].unique():
                    for m_value in df[(df["Symbol Part A"] == symbol) & (df["J"] == j_value)]["M"].unique():
                        for phase in moon_phases:
                            matched = df[(df["Symbol Part A"] == symbol) & 
                                        (df["Phase"] == phase) & 
                                        (df["J"] == j_value) & 
                                        (df["M"] == m_value)]

                            over_count = matched["C"].sum()
                            under_count = matched["D"].sum()
                            total = over_count + under_count

                            win_over = round(over_count / total, 2) if total > 0 else 0
                            win_under = round(under_count / total, 2) if total > 0 else 0

                            t_over_count += over_count
                            t_under_count += under_count
                            t_total += total

                            output_rows.append({
                                "Symbol": "",
                                "Symbol Part A": symbol,
                                "J": j_value,
                                "M": m_value,
                                "Phase": phase,
                                "Over count": over_count,
                                "Under count": under_count,
                                "Total": total,
                                "WIN% OVER": win_over,
                                "WIN% UNDER": win_under
                            })

                        # Add summary row for each J and M value combination
                        output_rows.append({
                            "Symbol": f"{symbol}-{j_value}-{m_value}",
                            "Symbol Part A": "",
                            "J": j_value,
                            "M": m_value,
                            "Phase": "",
                            "Over count": t_over_count,
                            "Under count": t_under_count,
                            "Total": t_total,
                            "WIN% OVER": round(t_over_count / t_total, 2) if t_total > 0 else 0,
                            "WIN% UNDER": round(t_under_count / t_total, 2) if t_total > 0 else 0
                        })
                        
                        # Reset counters for next combination
                        t_over_count = 0
                        t_under_count = 0
                        t_total = 0
            else :
                for phase in moon_phases:
                    matched = df[(df["Symbol Part A"] == symbol) & (df["Phase"] == phase)]

                    # print(matched["J"].values[0])
                    over_count = matched["C"].sum()
                    under_count = matched["D"].sum()
                    total = over_count + under_count

                    win_over = round(over_count / total, 2) if total > 0 else 0
                    win_under = round(under_count / total, 2) if total > 0 else 0

                    t_over_count += over_count
                    t_under_count += under_count
                    t_total += total
                
                    output_rows.append({
                        "Symbol": "",
                        "Symbol Part A": symbol,
                        "Symbol Part B": "",
                        "Phase": phase,
                        "Over count": over_count,
                        "Under count": under_count,
                        "Total": total,
                        "WIN% OVER": win_over,
                        "WIN% UNDER": win_under
                    })

                output_rows.append({
                    "Symbol": f"{symbol}-BLK",
                    "Symbol Part A": "",
                    "Symbol Part B": "",
                    "Phase": "",
                    "Over count": t_over_count,
                    "Under count": t_under_count,
                    "Total": t_total,
                    "WIN% OVER": round(t_over_count / t_total, 2) if t_total > 0 else 0,
                    "WIN% UNDER": round(t_under_count / t_total, 2) if t_total > 0 else 0
                })
            
        # Save to Excel
        result_df = pd.DataFrame(output_rows)
        output_path = os.path.splitext(data_file)[0] + "_phase_summary.xlsx"
        result_df.to_excel(output_path, index=False)

        auto_adjust_excel(output_path)

        result_label.config(text=f"File saved to:\n{output_path}")

    except Exception as e:
        result_label.config(text=f"Error: {str(e)}")

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

def browse_data_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel or CSV", "*.xlsx *.csv")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, filename)

def browse_symbol_file():
    filename = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    symbol_entry.delete(0, tk.END)
    symbol_entry.insert(0, filename)

# GUI
root = tk.Tk()
root.title("Moon Phase Symbol Analyzer")
root.geometry("530x360")

tk.Label(root, text="Select Excel or CSV File:").pack(pady=5)
file_entry = tk.Entry(root, width=60)
file_entry.pack()
tk.Button(root, text="Browse Data File", command=browse_data_file).pack(pady=5)

tk.Label(root, text="Select Symbol List (TXT):").pack(pady=5)
symbol_entry = tk.Entry(root, width=60)
symbol_entry.pack()
tk.Button(root, text="Browse Symbol File", command=browse_symbol_file).pack(pady=5)

tk.Button(root, text="Process", command=process_files, width=20).pack(pady=15)

result_label = tk.Label(root, text="", wraplength=480, fg="blue")
result_label.pack(pady=10)

root.mainloop()
