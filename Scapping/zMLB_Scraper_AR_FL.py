### Scraping bettingpros MLB Data

import requests
import pandas as pd
import json
from bs4 import BeautifulSoup
import re

# ----------------- Home/Away --------------------------------------------
url = 'https://www.bettingpros.com/mlb/props/'
response = requests.get(url)
soup = BeautifulSoup(response.text, 'html.parser')

scripts = soup.find_all('script')

var_pattern = re.compile(r'var\s+(\w+)\s*=\s*({.+?})(?=;|\n)', re.DOTALL)
result = {}

for script in scripts:
    if script.string:  
        matches = var_pattern.finditer(script.string)
        for match in matches:
            var_name = match.group(1)
            var_value = match.group(2)
            try:
                result[var_name] = json.loads(var_value)
            except json.JSONDecodeError:
                result[var_name] = var_value.strip()

results = result['playerPropAnalyzer']['events']

home_txt_list = []
visitor_txt_list = []
id_txt_list = []
for result in results:
    home_txt = result['home']
    visitor_txt = result['visitor']
    id_txt = result['id']
    
    home_txt_list.append(home_txt)
    visitor_txt_list.append(visitor_txt)
    id_txt_list.append(id_txt)


MyID_List_str = ":".join([str(num) for num in id_txt_list])
# print(home_txt_list)
# print(visitor_txt_list)
# print(MyID_List_str)

# ------------------------------------------------------------------------

headers = {
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'en-US,en;q=0.9,ar;q=0.8',
    'origin': 'https://www.bettingpros.com',
    'priority': 'u=1, i',
    'referer': 'https://www.bettingpros.com/',
    'sec-ch-ua': '"Chromium";v="134", "Not:A-Brand";v="24", "Google Chrome";v="134"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36',
    'x-api-key': 'CHi8Hy5CEE4khd46XNYL23dCFX96oUdw6qOt1Dnh',
}


def get_selected_markets():
    markets = {
        1: ("Home Runs", 299),
        2: ("Hits", 287),
        3: ("Runs", 288),
        4: ("RBI", 289),
        5: ("Strikeouts", 285),
        6: ("Doubles", 291),
        7: ("Triples", 292),
        8: ("Total Bases", 293),  
        9: ("Singles", 295),
        10: ("Steals", 294),
        11: ("Earned Runs", 290)
    }
    
    print("Select stats to scrape (enter numbers separated by commas):")
    for num, (name, id) in markets.items():
        print(f"{num}: {name}")
    
    while True:
        user_input = input("Your choice: ")
        try:
            selected_nums = [int(num.strip()) for num in user_input.split(',')]
            selected_markets = {num: markets[num] for num in selected_nums}
            break
        except (ValueError, KeyError):
            print("Invalid input. Please enter numbers separated by commas (e.g., 1,3,5)")
    
    market_names = [markets[num][0] for num in selected_nums]
    market_ids = [markets[num][1] for num in selected_nums]
    
    return {
        'selected_numbers': selected_nums,
        'market_names': market_names,
        'market_ids': market_ids
    }

result = get_selected_markets()
print("\nSelected Markets:")
for num, name, id in zip(result['selected_numbers'], 
                        result['market_names'], 
                        result['market_ids']):
    print(f"{num}: {name} (ID: {id})")

response = requests.get(
    f'https://api.bettingpros.com/v3/props?limit=1000_000&page=1&sport=MLB&market_id={id}&event_id={MyID_List_str}&location=INT&sort=trending&include_selections=false&include_markets=true&include_counts=true',
    headers=headers,
)

results = response.json()['props']
utc_date = response.json()['utc']

participant_slug_list = []
participant_team_list = []
consensus_line_list = []
odds_list = []
date_list = []
home_away_list = []

print(f'\nProcessing {len(results)} Player ...\n')

for result in results:
    participant_slug= result['participant']['player']['slug'] #participant_slug
    participant_team= result['participant']['player']['team'] #participant_team
    try:
        consensus_line= result['under']['consensus_line'] #consensus_line
    except Exception as e:
        consensus_line = 0

    over_odds= result['over']['consensus_odds'] #-125  O -125 | U -109

    try:
        under_odds= result['under']['consensus_odds'] #-109  O -125 | U -109
    except Exception as e:
        under_odds = 0

    odds = f'O {over_odds} | U {under_odds}'
    date = utc_date
    if participant_team in home_txt_list:
        home_away = "HOME"
    else:
        home_away = "AWAY"

    participant_slug_list.append(participant_slug)
    participant_team_list.append(participant_team)
    consensus_line_list.append(consensus_line)
    odds_list.append(odds)
    date_list.append(f'"{date[5:7]}-{date[8:10]}"')
    home_away_list.append(home_away)
    
    df = pd.DataFrame({
        'Player Name': participant_slug_list,
        'Number': consensus_line_list,
        'Odds': odds_list,
        'Projection': '1',
        'Avg': '1',
        'Home/Away Avg': '1',
        'Home/Away': home_away_list,
        'Date': date_list,
        'Stat Category': name,
        'Team': participant_team_list
    })

    output_name = f'{name}_{date[5:7]}{date[8:10]}{date[0:4]}_player_props_FL.csv'
    df.to_csv(output_name, index=False)
    
print('Data is Successfully Saved to Excel File')
