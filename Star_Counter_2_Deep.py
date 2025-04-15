import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os

def analyze_excel_data(input_file, output_file):
    try:
        df = pd.read_excel(input_file)

        grouped = df.groupby([df.columns[13], df.columns[16]])
        
        results = []
        for (symbol, phase), group in grouped:
            over_count = sum(group[df.columns[7]] == 'OVER')
            under_count = sum(group[df.columns[7]] == 'UNDER')
            total = over_count + under_count
            win_pct_over = round(over_count / total, 2) if total != 0 else 0.0
            win_pct_under = round(under_count / total, 2) if total != 0 else 0.0

            results.append({
                'Symbol': symbol,
                'Phase': phase,
                'Over count': over_count,
                'Under count': under_count,
                'Total': total,
                'WIN% OVER': win_pct_over,
                'WIN% UNDER': win_pct_under
            })

        result_df = pd.DataFrame(results)
        result_df.to_excel(output_file, index=False)

        return True, f"Analysis complete. Results saved to {output_file}"
    
    except Exception as e:
        return False, f"Error: {str(e)}"

class ExcelAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Analysis")
        self.root.geometry("600x300")
        self.root.resizable(True, True)

        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Input
        self.input_frame = ttk.Frame(self.main_frame)
        self.input_frame.pack(fill=tk.X, pady=10)

        ttk.Label(self.input_frame, text="Excel File:").pack(side=tk.LEFT, padx=5)

        self.input_path = tk.StringVar()
        self.input_entry = ttk.Entry(self.input_frame, textvariable=self.input_path, width=50)
        self.input_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.browse_btn = ttk.Button(self.input_frame, text="Browse", command=self.browse_file)
        self.browse_btn.pack(side=tk.RIGHT, padx=5)

        # Output
        self.output_frame = ttk.Frame(self.main_frame)
        self.output_frame.pack(fill=tk.X, pady=10)

        ttk.Label(self.output_frame, text="Output Excel:").pack(side=tk.LEFT, padx=5)

        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(self.output_frame, textvariable=self.output_path, width=50)
        self.output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.browse_output_btn = ttk.Button(self.output_frame, text="Browse", command=self.browse_output_file)
        self.browse_output_btn.pack(side=tk.RIGHT, padx=5)

        self.process_btn = ttk.Button(self.main_frame, text="Process Data", command=self.process_data)
        self.process_btn.pack(pady=20)

        self.status_frame = ttk.LabelFrame(self.main_frame, text="Status")
        self.status_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.status_text = tk.Text(self.status_frame, height=5, wrap=tk.WORD)
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.status_text.config(state=tk.DISABLED)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filename:
            self.input_path.set(filename)
            base_dir = os.path.dirname(filename)
            base_name = os.path.splitext(os.path.basename(filename))[0]
            self.output_path.set(os.path.join(base_dir, f"{base_name}_analysis.xlsx"))

    def browse_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if filename:
            self.output_path.set(filename)

    def update_status(self, message, is_error=False):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(tk.END, message)
        if is_error:
            self.status_text.tag_configure("error", foreground="red")
            self.status_text.tag_add("error", "1.0", "end")
        self.status_text.config(state=tk.DISABLED)

    def process_data(self):
        input_file = self.input_path.get().strip()
        output_file = self.output_path.get().strip()

        if not input_file:
            self.update_status("Please select an input Excel file.", True)
            return

        if not output_file:
            self.update_status("Please specify an output Excel file.", True)
            return

        self.process_btn.config(state=tk.DISABLED)
        self.update_status("Processing data... Please wait.")
        self.root.update()

        success, message = analyze_excel_data(input_file, output_file)

        self.process_btn.config(state=tk.NORMAL)

        self.update_status(message, not success)
        if success:
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", message)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelAnalysisApp(root)
    root.mainloop()
