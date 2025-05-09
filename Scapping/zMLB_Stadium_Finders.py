import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager



from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def scrape_games(date):
    """Scrape MLB games and return a dictionary of away teams."""
    url = f"https://www.mlb.com/scores/{date}"
    print(f"[DEBUG] Scraping MLB games from: {url}")

    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(url)

    away_teams = {}

    # Find all <a> elements with the specific class and href starting with "/"
    games = driver.find_elements(By.CSS_SELECTOR, 'div.TeamMatchupLayerstyle__InlineWrapper-sc-7tca6g-1')   

    for i in range(len(games)): 
        try:
            group = games[i].find_elements(By.CSS_SELECTOR, 'div.TeamWrappersstyle__DesktopTeamWrapper-sc-uqs6qh-0') 
            if len(group) >= 2:
                away_team = ''.join(c for c in group[1].text.lower() if c.isalpha())
                home_team = ''.join(c for c in group[0].text.lower() if c.isalpha())

                if away_team != home_team:
                    away_teams[away_team] = home_team
                    print(f"[INFO] {away_team.upper()} @ {home_team.upper()}")
                else:
                    print(f"[WARN] Skipping game {i}: Home and Away teams are the same ({away_team.upper()})")
        except Exception as e:
            print(f"[WARN] Skipping game {i}: {e}")

    driver.quit()

    print(away_teams)
    print(f"[DEBUG] Found {len(away_teams)} valid away teams.")
    return away_teams

def load_player_data(excel_file1, excel_file2):
    """Load player data from Excel files and return two dictionaries."""
    print(f"[DEBUG] Loading player data from: {excel_file1}")
    df1 = pd.read_excel(excel_file1, engine="openpyxl", header=None)
    df1 = df1.iloc[1:, :]  # Skip header row

    # Extract player names (Column BA) and team abbreviations (Column BJ)
    players = df1.iloc[:, 52].dropna().str.lower().tolist()  # Column BA (Player Name)
    teams = df1.iloc[:, 61].dropna().str.upper().tolist()  # Column BJ (Team Abbreviations)
    player_team_map = dict(zip(players, teams))

    print(f"[DEBUG] Loading birth data from: {excel_file2}")
    df2 = pd.read_excel(excel_file2, engine="openpyxl")

    # Ensure 'Betting Pros Com Slug' column exists
    if "Betting Pros Com Slug" not in df2.columns or "BirthDate" not in df2.columns or "Address" not in df2.columns:
        print("[ERROR] Missing required columns in the second Excel file!")
        return {}, {}

    # Create a mapping of 'Betting Pros Com Slug' -> BirthDate, Address
    birth_data_map = {}
    for _, row in df2.iterrows():
        slug = str(row['Betting Pros Com Slug']).strip().lower()  # Convert to string before applying .strip() and .lower()
        
        # Ensure BirthDate and Address are valid before adding
        if pd.notna(row["BirthDate"]) and pd.notna(row["Address"]):
            birth_data_map[slug] = {"dob": row["BirthDate"], "location": row["Address"]}
        else:
            print(f"[WARN] Missing birth info for slug: {slug}")

    print(f"[DEBUG] Loaded {len(birth_data_map)} birth records from the second Excel file.")

    # Match players from the first file with birth data using 'Betting Pros Com Slug'
    player_birth_map = {}
    for player in players:
        if player in birth_data_map:
            player_birth_map[player] = birth_data_map[player]
            print(f"[MATCH] Found birth data for: {player}")
        else:
            print(f"[NO MATCH] No birth data found for: {player}")

    print(f"[DEBUG] Matched {len(player_birth_map)} players with birth records.")

    return player_team_map, player_birth_map, players  # Add `players` list


# --- Function to Load Team Locations ---
def load_team_locations(txt_file):
    """Load team locations from a TXT file into dictionaries for lookup."""
    print(f"[DEBUG] Loading team locations from: {txt_file}")
    team_locations = {}
    team_name_to_abbr = {}

    with open(txt_file, 'r') as file:
        lines = [line.strip().lower() for line in file.readlines() if line.strip()]
        for i in range(0, len(lines), 3):
            if i + 2 < len(lines):
                abbr = lines[i].upper()  # Team abbreviation (ATH, BAL, etc.)
                team_name = lines[i + 1]  # Full team name (athletics, orioles, etc.)
                location = lines[i + 2]  # Stadium location

                team_locations[abbr] = location
                team_name_to_abbr[team_name] = abbr  # Map full name to abbreviation

    print(f"[DEBUG] Team abbreviations loaded: {team_name_to_abbr}")
    print(f"[DEBUG] Team locations loaded: {team_locations}")
    return team_locations, team_name_to_abbr

# --- Function to Load Team Locations ---
def load_team_locations(txt_file):
    """Load team locations from a TXT file into dictionaries for lookup."""
    print(f"[INFO] Loading team locations from: {txt_file}")
    team_locations = {}
    team_name_to_abbr = {}

    with open(txt_file, 'r') as file:
        lines = [line.strip().lower() for line in file.readlines() if line.strip()]
        for i in range(0, len(lines), 3):
            if i + 2 < len(lines):
                abbr = lines[i].upper()  # Team abbreviation (ATH, BAL, etc.)
                team_name = lines[i + 1]  # Full team name (athletics, orioles, etc.)
                location = lines[i + 2]  # Stadium location

                team_locations[abbr] = location
                team_name_to_abbr[team_name] = abbr  # Map full name to abbreviation

    print(f"[INFO] Loaded {len(team_locations)} team locations.")
    return team_locations, team_name_to_abbr

def process_and_save(date, away_teams, team_locations, team_name_to_abbr, player_team_map, player_birth_map, players):
    """Match players with games (in original order) and save to CSV."""
    print("[DEBUG] Matching players with games in original order...")
    data = []
    year, month, day = date.split("-")
    month_names = {
        '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun',
        '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
    }
    month_name = month_names.get(month, "Unknown")

    for player in players:
        player_team = player_team_map.get(player)
        if not player_team:
            continue

        matched_game = None
        for away_team_name, home_team_name in away_teams.items():
            home_abbr = team_name_to_abbr.get(home_team_name)
            away_abbr = team_name_to_abbr.get(away_team_name)

            if not home_abbr or not away_abbr:
                continue

            if player_team == home_abbr or player_team == away_abbr:
                matched_game = (home_abbr, away_abbr)
                break

        if not matched_game:
            continue  # Player not in any game today

        # Use away team location (i.e., where the game is happening)
        location = team_locations.get(matched_game[1], "Unknown Location")
        player_info = player_birth_map.get(player, {})
        dob = player_info.get('dob', "Unknown DOB")
        birth_location = player_info.get('location', "Unknown Birth Location")

        if dob != "Unknown DOB":
            dob_parsed = pd.to_datetime(dob)
            dob_day = str(dob_parsed.day).zfill(2)
            dob_month_name = month_names.get(str(dob_parsed.month).zfill(2), "Unknown")
            dob_year = str(dob_parsed.year)
        else:
            dob_day = dob_month_name = dob_year = "Unknown"

        data.append([
            day, month_name, year, location,
            "", "", "",  # Three empty columns as separators
            dob_day, dob_month_name, dob_year,
            birth_location
        ])

    # Save to CSV
    csv_filename = f"mlb_players_{date}.csv"
    df_output = pd.DataFrame(data, columns=[
        "Day", "Month", "Year", "Location",
        "", "", "",
        "DOB Day", "DOB Month", "DOB Year",
        "Birth Location"
    ])
    df_output.to_csv(csv_filename, index=False)
    print(f"[DEBUG] Data saved to {csv_filename}")




# --- Main Processing Function ---
def run_process(date, excel_file1, excel_file2, txt_file):
    try:
        away_teams = scrape_games(date)
        team_locations, team_name_to_abbr = load_team_locations(txt_file)
        player_team_map, player_birth_map, players = load_player_data(excel_file1, excel_file2)

        process_and_save(date, away_teams, team_locations, team_name_to_abbr, player_team_map, player_birth_map, players)

        messagebox.showinfo("Success", "Processing completed successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# --- Tkinter GUI ---
def browse_file(entry, filetypes):
    filename = filedialog.askopenfilename(filetypes=filetypes)
    if filename:
        entry.delete(0, tk.END)
        entry.insert(0, filename)

def start_thread():
    date_val = date_entry.get().strip()
    excel1_val = excel1_entry.get().strip()
    excel2_val = excel2_entry.get().strip()
    txt_val = txt_entry.get().strip()
    # time_option = time_var.get()  # Get the selected radio button value

    if not date_val or not excel1_val or not excel2_val or not txt_val:
        messagebox.showwarning("Missing Data", "Please fill in all fields!")
        return

    thread = threading.Thread(target=run_process, args=(date_val, excel1_val, excel2_val, txt_val))
    thread.start()

# Create main window
root = tk.Tk()
root.title("MLB Data Processor")

# Date input
tk.Label(root, text="Enter date (YYYY-MM-DD):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
date_entry = tk.Entry(root, width=20)
date_entry.grid(row=0, column=1, padx=5, pady=5)

# Time option radio buttons
# time_var = tk.StringVar(value="current")
# time_frame = tk.Frame(root)
# time_frame.grid(row=1, column=0, columnspan=3, pady=5)
# tk.Label(time_frame, text="Time option:").pack(side=tk.LEFT, padx=5)
# tk.Radiobutton(time_frame, text="Before", variable=time_var, value="before").pack(side=tk.LEFT, padx=5)
# tk.Radiobutton(time_frame, text="Current", variable=time_var, value="current").pack(side=tk.LEFT, padx=5)

# Excel 1 input
tk.Label(root, text="Select 1st Excel file (.xlsx):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
excel1_entry = tk.Entry(root, width=40)
excel1_entry.grid(row=2, column=1, padx=5, pady=5)
tk.Button(root, text="Browse...", command=lambda: browse_file(excel1_entry, [("Excel Files", "*.xlsx")])).grid(row=2, column=2, padx=5, pady=5)

# Excel 2 input
tk.Label(root, text="Select 2nd Excel file (.xlsx):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
excel2_entry = tk.Entry(root, width=40)
excel2_entry.grid(row=3, column=1, padx=5, pady=5)
tk.Button(root, text="Browse...", command=lambda: browse_file(excel2_entry, [("Excel Files", "*.xlsx")])).grid(row=3, column=2, padx=5, pady=5)

# TXT input
tk.Label(root, text="Select team locations file (.txt):").grid(row=4, column=0, padx=5, pady=5, sticky="e")
txt_entry = tk.Entry(root, width=40)
txt_entry.grid(row=4, column=1, padx=5, pady=5)
tk.Button(root, text="Browse...", command=lambda: browse_file(txt_entry, [("Text Files", "*.txt")])).grid(row=4, column=2, padx=5, pady=5)

# Start button
tk.Button(root, text="Start Processing", command=start_thread, bg="lightgreen").grid(row=5, column=0, columnspan=3, pady=10)

root.mainloop()
