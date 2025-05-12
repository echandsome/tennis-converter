import pandas as pd
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
import os
import time
from pybaseball import playerid_lookup

def make_link(name):
    parts = str(name).strip().split()
    if len(parts) < 2:
        return name
    
    lastname = parts[-1]
    firstname = parts[0]
    
    lastname_part = lastname[:5] if len(lastname) >= 5 else lastname
    firstname_part = firstname[:2] if len(firstname) >= 2 else firstname
    lastname_first_char = lastname[0]
    
    return f"players/{lastname_first_char}/{lastname_part}{firstname_part}"

def process_file():
    file_path = file_entry.get()
    if not file_path:
        result_label.config(text="Please select an Excel file.")
        return

    try:
        df = pd.read_excel(file_path, header=None, engine="openpyxl")

        df.iloc[:, 0] = df.iloc[:, 0].astype(str)

        # Selenium - Display browser
        driver = webdriver.Chrome()
        driver.get("https://www.baseball-reference.com/leaders/HR_career.shtml")
        time.sleep(3)

        # Find HR table
        table = driver.find_element(By.ID, "leader_standard_HR")
        links = table.find_elements(By.TAG_NAME, "a")

        # Collect all links' href
        hrefs = [(link.text.strip(), link.get_attribute("href")) for link in links if link.get_attribute("href")]
        for i, row in df.iterrows():
            name = row.iloc[1] 
            reversed_name = make_link(name)
            matched_href = ""
            for text, href in hrefs:
                if reversed_name in href:
                    matched_href = href
                    break
            
            if matched_href:
                df.at[i, 0] = matched_href  # Save to column A
            else:
                parts = str(name).strip().split()
                if len(parts) < 2:
                    return name
                
                lastname = parts[-1]
                firstname = parts[0]
                data = playerid_lookup(lastname, firstname) 
                try:
                    player_id = data.at[0, 'key_bbref']
                    df.at[i, 0] = f"https://www.baseball-reference.com/players/{player_id[0]}/{player_id}.shtml"
                except Exception as e:
                    df.at[i, 0] = ''

        driver.quit()

        output_path = os.path.join(os.path.dirname(file_path), "player_links_output.xlsx")
        df.to_excel(output_path, index=False)
        result_label.config(text=f"Done! Saved to:\n{output_path}")

    except Exception as e:
        print(e)
        result_label.config(text=f"Error: {str(e)}")

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, filename)

# GUI
root = tk.Tk()
root.title("Player Link Finder")
root.geometry("460x240")

tk.Label(root, text="Select Excel File:").pack()
file_entry = tk.Entry(root, width=55)
file_entry.pack()
tk.Button(root, text="Browse", command=browse_file).pack(pady=5)

tk.Button(root, text="Start", command=process_file).pack(pady=15)

result_label = tk.Label(root, text="", wraplength=440)
result_label.pack()

root.mainloop()
