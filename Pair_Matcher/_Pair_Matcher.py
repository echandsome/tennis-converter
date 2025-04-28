import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import openpyxl

def clean_column_names(df):
    # Replace any column name starting with 'Unnamed' with empty values
    df.columns = [col if not col.startswith('Unnamed') else '' for col in df.columns]
    return df

def auto_adjust_columns_width(out_path):
    # Auto-adjust column width
    wb = openpyxl.load_workbook(out_path)
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

    wb.save(out_path)

def process_combined_output(comparison_path, daily_path, group_option):
    # Load the comparison and daily data Excel files
    comp_df = pd.read_excel(comparison_path)
    daily_df = pd.read_excel(daily_path)

    output_rows = []

    for idx, row in comp_df.iterrows():
        player_a = str(row.iloc[0]).strip()
        player_b = str(row.iloc[1]).strip()
        additional_match = str(row.iloc[2]).strip()

        comp_values = row.tolist()

        # Determine group keys
        if group_option == 'L':
            # Search for player A's info
            daily_a = daily_df[(daily_df.iloc[:, 0].astype(str).str.strip() == player_a) &
                (daily_df.iloc[:, 11].astype(str).str.strip() == additional_match)   ]

            # Search for player B's info
            daily_b = daily_df[(daily_df.iloc[:, 0].astype(str).str.strip() == player_b) &
                (daily_df.iloc[:, 11].astype(str).str.strip() == additional_match)]
        elif group_option == 'N':
             # Search for player A's info
            daily_a = daily_df[(daily_df.iloc[:, 0].astype(str).str.strip() == player_a) &
                (daily_df.iloc[:, 13].astype(str).str.strip() == additional_match)   ]

            # Search for player B's info
            daily_b = daily_df[(daily_df.iloc[:, 0].astype(str).str.strip() == player_b) &
                (daily_df.iloc[:, 13].astype(str).str.strip() == additional_match)]
        else:
            # Search for player A's info
            daily_a = daily_df[daily_df.iloc[:, 0].astype(str).str.strip() == player_a]

            # Search for player B's info
            daily_b = daily_df[daily_df.iloc[:, 0].astype(str).str.strip() == player_b]

        if daily_a.empty or daily_b.empty:
            continue

        daily_a_row = daily_a.iloc[0].tolist()
        output_rows.append(comp_values + [''] + daily_a_row)
    
        daily_b_row = daily_b.iloc[0].tolist()
        output_rows.append(comp_values + [''] + daily_b_row)

        # Insert an empty row after each pair
        output_rows.append([''] * (len(comp_values) + 1 + len(daily_df.columns)))

    daily_df = clean_column_names(daily_df)

    # Combine headers
    comp_headers = comp_df.columns.tolist()
    daily_headers = daily_df.columns.tolist()
    combined_headers = comp_headers + [''] + daily_headers

    # Create final DataFrame
    df_output = pd.DataFrame([combined_headers] + output_rows)

    # Save as .xlsx file
    out_dir = os.path.dirname(comparison_path)
    out_path = os.path.join(out_dir, "Combined_Match_Result.xlsx")
    df_output.to_excel(out_path, index=False, header=False)

    auto_adjust_columns_width(out_path)

    process_grouped_output(out_path, group_option)

def process_grouped_output(input_path, group_option):
    df = pd.read_excel(input_path)

    # Determine if columns are offset (N or L implies one extra column at front)
    offset = 1 if group_option in ['N', 'L'] else 0

    # Adjusted output header
    base_header = [
        "", "Player Count", "Total", "", "", "", "", "", "", "",
        "Player", "FILL", "", "", "", "", "",                           
        "FILL", "", "",  "",                                          
        "FILL", "FILL", "FILL",                           
        "", "", "" , ""                                   
    ]
    output_header = [""] * offset + base_header  # prepend empty string if offset

    # Determine group keys based on option
    if group_option in ['L', 'N']:
        group_keys = [df.iloc[:, 11 + offset], df.iloc[:, 2]]
    else:
        group_keys = [df.iloc[:, 10]]

    output_rows = []

    grouped = df.groupby(group_keys)
    
    for _, group_df in grouped:
        
        modified_group = group_df.copy()

        modified_group.iloc[:, 12 + offset:17 + offset] = ''
        modified_group.iloc[:, 18 + offset:21 + offset] = ''

        summary_row = pd.Series([''] * df.shape[1], index=df.columns)
        summary_row.iloc[1 + offset] = len(group_df)              
        summary_row.iloc[2 + offset] = group_df.iloc[:, 2 + offset].sum()

        for idx in [10, 11, 17, 21, 22, 23]:
            summary_row.iloc[idx + offset] = group_df.iloc[0, idx + offset]

        output_rows.append(modified_group)
        output_rows.append(pd.DataFrame([summary_row]))

    final_df = pd.concat(output_rows, ignore_index=True)

    final_df.columns = output_header

    # Save as .xlsx file
    out_dir = os.path.dirname(input_path)
    out_path = os.path.join(out_dir, "Grouped_Match_Result.xlsx")
    final_df.to_excel(out_path, index=False, header=False)

    auto_adjust_columns_width(out_path)

    result_label.config(text="XLSX file saved successfully in the same folder.", fg="green")

# GUI logic
def browse_comparison():
    path = filedialog.askopenfilename(title="Select Comparison.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        comparison_path_var.set(path)

def browse_daily():
    path = filedialog.askopenfilename(title="Select Daily_File.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if path:
        daily_path_var.set(path)

def run_combination():
    comparison_path = comparison_path_var.get()
    daily_path = daily_path_var.get()
    group_option = group_option_var.get()
    if not comparison_path or not daily_path:
        result_label.config(text="Please select both files.", fg="red")
        return
    result_label.config(text="Processing...", fg="blue")
    root.update()
    try:
        process_combined_output(comparison_path, daily_path, group_option)
    except Exception as e:
        print(e)
        result_label.config(text=f"Error: {str(e)}", fg="red")

# GUI Setup
root = tk.Tk()
root.title("Player Comparison Merger (.xlsx output)")
root.geometry("500x300")

comparison_path_var = tk.StringVar()
daily_path_var = tk.StringVar()
group_option_var = tk.StringVar(value="None")

tk.Button(root, text="Select Comparison File", command=browse_comparison).pack(pady=(10, 0))
tk.Label(root, textvariable=comparison_path_var, wraplength=480).pack()

tk.Button(root, text="Select Daily File", command=browse_daily).pack(pady=(10, 0))
tk.Label(root, textvariable=daily_path_var, wraplength=480).pack()

# Group Option Radio Buttons
group_frame = tk.LabelFrame(root, text="Group By")
group_frame.pack(pady=10)
for text in ["None", "L", "N"]:
    tk.Radiobutton(group_frame, text=text, variable=group_option_var, value=text).pack(side="left", padx=10)

tk.Button(root, text="Run Matcher", command=run_combination, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
