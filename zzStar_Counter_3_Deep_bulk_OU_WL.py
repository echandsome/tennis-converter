import pandas as pd
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os

def analyze_excel_data(input_file, output_file, group_by_choice):
    try:
        df = pd.read_excel(input_file)
        columns = df.columns

        group_columns = []

        if group_by_choice == 'AL' or group_by_choice == 'AN':
            group_columns = [columns[0], columns[11]]
            if group_by_choice == 'AN':
                group_columns = [columns[0], columns[13]]
        elif group_by_choice == 'AJN' or group_by_choice == 'AJNQ':
            group_columns = [columns[0], columns[9], columns[13]]
            if group_by_choice == 'AJNQ':
                group_columns = [columns[0], columns[9], columns[13], columns[16]]
        else:
            # Column index assumption: L(13), Q(16), J(9), M(12)
            group_columns = [columns[13], columns[16]]  # Default L, Q

            if group_by_choice == "LQJ":
                group_columns.append(columns[9])  # J
            elif group_by_choice == "LQM":
                group_columns.append(columns[12])  # M
            elif group_by_choice == "LQJM":
                group_columns.extend([columns[9], columns[12]])  # J, M

        grouped = df.groupby(group_columns)

        results = []
        for keys, group in grouped:
            key_values = keys if isinstance(keys, tuple) else (keys,)

            if group_by_choice == 'AL' or group_by_choice == 'AN':
                if group_by_choice == "AL":
                    result = dict(zip(['Player', 'Signs'][:len(key_values)], key_values))
                else :
                    result = dict(zip(['Player', 'Signs-Symbol'][:len(key_values)], key_values))
            elif group_by_choice == 'AJN' or group_by_choice == 'AJNQ':
                if group_by_choice == "AJN":
                    result = dict(zip(['Player', 'J', 'Signs-Symbol'][:len(key_values)], key_values))
                else :
                    result = dict(zip(['Player', 'J', 'Signs-Symbol', 'Phase'][:len(key_values)], key_values))
            else:
                if group_by_choice == "LQM":
                    result = dict(zip(['Symbol', 'Phase', 'M', 'M'][:len(key_values)], key_values))
                else :
                    result = dict(zip(['Symbol', 'Phase', 'J', 'M'][:len(key_values)], key_values))

            over_count = sum(group[columns[7]].str.lower() == 'over')
            under_count = sum(group[columns[7]].str.lower() == 'under')
            win_count = sum(group[columns[7]].str.lower() == 'win')
            lose_count = sum(group[columns[7]].str.lower() == 'lose')

            result['Over count'] = over_count + win_count
            result['Under count'] = under_count + lose_count
            results.append(result)

        with open(output_file, 'w', newline='') as csvfile:
            fieldnames = list(results[0].keys()) if results else ['Symbol', 'Phase', 'Over count', 'Under count']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for row in results:
                writer.writerow(row)

        return True, f"{os.path.basename(input_file)} - Completed"
    except Exception as e:
        return False, f"{os.path.basename(input_file)} - Error occurred: {str(e)}"

class BulkExcelAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Excel Data Analysis")
        self.root.geometry("750x500")

        self.group_by_choice = tk.StringVar(value="AL")

        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        folder_frame = ttk.Frame(self.main_frame)
        folder_frame.pack(fill=tk.X, pady=10)

        ttk.Label(folder_frame, text="Input Folder:").pack(side=tk.LEFT, padx=5)

        self.input_folder = tk.StringVar()
        self.input_entry = ttk.Entry(folder_frame, textvariable=self.input_folder, width=50)
        self.input_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        browse_btn = ttk.Button(folder_frame, text="Browse", command=self.browse_folder)
        browse_btn.pack(side=tk.RIGHT, padx=5)

        # Radio buttons
        group_frame = ttk.LabelFrame(self.main_frame, text="Group By")
        group_frame.pack(fill=tk.X, pady=10)

        ttk.Radiobutton(group_frame, text="A & L", variable=self.group_by_choice, value="AL").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(group_frame, text="A & N", variable=self.group_by_choice, value="AN").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(group_frame, text="A, J & N", variable=self.group_by_choice, value="AJN").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(group_frame, text="A, J, N & Q", variable=self.group_by_choice, value="AJNQ").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(group_frame, text="L & Q", variable=self.group_by_choice, value="LQ").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(group_frame, text="L, Q & J", variable=self.group_by_choice, value="LQJ").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(group_frame, text="L, Q & M", variable=self.group_by_choice, value="LQM").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(group_frame, text="L, Q, J & M", variable=self.group_by_choice, value="LQJM").pack(side=tk.LEFT, padx=10)

        self.process_btn = ttk.Button(self.main_frame, text="Process All Excel Files", command=self.process_all_files)
        self.process_btn.pack(pady=20)

        self.status_frame = ttk.LabelFrame(self.main_frame, text="Status")
        self.status_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.status_text = tk.Text(self.status_frame, height=10, wrap=tk.WORD)
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.status_text.config(state=tk.DISABLED)

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self.input_folder.set(folder)

    def update_status(self, message, is_error=False):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n")
        if is_error:
            self.status_text.tag_configure("error", foreground="red")
            self.status_text.tag_add("error", "end-2l linestart", "end-1c")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.root.update()

    def process_all_files(self):
        input_dir = self.input_folder.get().strip()
        if not input_dir:
            self.update_status("❌ Please select an input folder.", True)
            return

        parent_dir = os.path.dirname(input_dir)
        folder_name = os.path.basename(input_dir)
        output_dir = os.path.join(parent_dir, folder_name + "_analysis")
        os.makedirs(output_dir, exist_ok=True)

        self.process_btn.config(state=tk.DISABLED)
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.update_status(f"📂 Input folder: {input_dir}")
        self.update_status(f"💾 Output folder: {output_dir}\n")

        for filename in os.listdir(input_dir):
            if filename.endswith((".xlsx", ".xls")) and not filename.startswith("~$"):
                input_file = os.path.join(input_dir, filename)
                base_name = os.path.splitext(filename)[0]
                output_file = os.path.join(output_dir, f"{base_name}_analysis.csv")
                success, msg = analyze_excel_data(input_file, output_file, self.group_by_choice.get())
                self.update_status(("✅ " if success else "❌ ") + msg, not success)

        self.update_status("\n🎉 All files processed!")
        self.process_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = BulkExcelAnalysisApp(root)
    root.mainloop()
