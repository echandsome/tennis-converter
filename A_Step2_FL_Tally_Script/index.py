import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to add an aggregate row to each player's DataFrame
def add_aggregate_row(df):
    print(df.columns)
    agg_row = {
        'Over count': df['Over count'].sum(),
        'Under count': df['Under count'].sum(),
        'Total': df['Total'].sum(),
        'WIN% OVER': df['WIN% OVER'].sum(),
        'WIN% UNDER': df['WIN% UNDER'].sum(),
        'COL VW': df['COL VW'].map(lambda x: x == 'O').sum(),
        'COL VW.1': df['COL VW.1'].map(lambda x: x == 'U').sum(),
        "Remarks": "Result"
    }
    
    # Copy last row to preserve other attributes
    last_row = df.iloc[-1].copy()
    for col in agg_row:
        last_row[col] = agg_row[col]
    
    # Concatenate the new row to the DataFrame
    df = pd.concat([df, last_row.to_frame().T], ignore_index=True)
    
    return df
# Function to separate each player's data into a dictionary of DataFrames
def separate_player_data(data):
    player_dfs = {}
    current_player_data = []
    player_count = 1
    
    for _, row in data.iterrows():
        if pd.isna(row['Partner A']) and pd.isna(row['Partner B']):
            if current_player_data:
                player_df = pd.DataFrame(current_player_data)
                # Add an empty "Remarks" column when creating the DataFrame
                player_df['Remarks'] = ""

                player_dfs[f'Player_{player_count}'] = player_df
                player_count += 1
                current_player_data = []
        else:
            current_player_data.append(row)
    
    return player_dfs

# Function to handle file selection and processing
def process_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return
    
    try:
        status_label.config(text="File loaded. Processing file now...", fg="blue")
        root.update()  # Refresh the GUI
        # Load the Excel file
        data = pd.read_excel(file_path, sheet_name='Sheet1')
        original_columns = data.columns
        second_header = data.iloc[0]
        data.drop(index=0, inplace=True)
        
        # Separate player data
        player_dfs = separate_player_data(data)
        
        # Combine all player DataFrames
        combined_df = pd.DataFrame()
        for i, (player, df) in enumerate(player_dfs.items()):
            df = add_aggregate_row(df)
            if i == 0:
                combined_df = pd.concat([combined_df, df], ignore_index=True)
            else:
                combined_df = pd.concat([combined_df, df], ignore_index=True)

        # Add original headers and second row at the top
        new_columns = [None if str(x).startswith('Unnamed') else x for x in original_columns]
        if len(new_columns) < combined_df.shape[1]:
            new_columns.append("Remarks")
        combined_df.columns = new_columns
        combined_df.loc[-1] = second_header
        combined_df.index = combined_df.index + 1
        combined_df = combined_df.sort_index()

        # Save to a new Excel file
        # output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        # if output_file:
        output_file = "after.xlsx"
        combined_df.to_excel(output_file, index=False)
        status_label.config(text=f"File saved to: {output_file}", fg="green")
        messagebox.showinfo("Success", f"File saved to: {output_file}")
    
    except Exception as e:
        raise e
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI Setup
root = tk.Tk()
root.title("Player Data Aggregator")
root.geometry("400x200")

label = tk.Label(root, text="Select an Excel file to process:", font=("Arial", 12))
label.pack(pady=10)

btn = tk.Button(root, text="Select File", command=process_file, font=("Arial", 12), bg="#4CAF50", fg="white", padx=10, pady=5)
btn.pack(pady=5)

label2 = tk.Label(root, text="Select output name", font=("Arial", 12))
label2.pack(pady=10)



# Status label
status_label = tk.Label(root, text="", font=("Arial", 10))
status_label.pack(pady=5)

root.mainloop()