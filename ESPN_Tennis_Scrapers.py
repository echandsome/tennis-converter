from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
import time
from bs4 import BeautifulSoup
import pandas as pd

# Function to generate URL based on date
def generate_url(date):
    return f"https://www.espn.com/tennis/scoreboard/_/date/{date.strftime('%Y%m%d')}"

# Initialize WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Input date range (from and to)
start_date_input = input("Enter the start date (YYYY-MM-DD): ")
end_date_input = input("Enter the end date (YYYY-MM-DD): ")

# Convert input dates to datetime objects
start_date = datetime.strptime(start_date_input, "%Y-%m-%d")
end_date = datetime.strptime(end_date_input, "%Y-%m-%d")

# Prepare the list to store the data
data = []

# Loop through the date range
current_date = start_date
while current_date <= end_date:
    url = generate_url(current_date)
    print(f"Scraping data for {url}")
    
    # Navigate to the URL
    driver.get(url)
    time.sleep(3)  # Wait for the page to load
    
    # Find all tournament cards on the page
    cards = driver.find_elements(By.CLASS_NAME, "Card")
    print(f"Found {len(cards)} tournament cards on the page.")

    tournament_locations = {}  # Dictionary to store tournament locations
    
    # Loop through each tournament card
    for idx, card in enumerate(cards):
        card_html = card.get_attribute("innerHTML")  
        soup = BeautifulSoup(card_html, 'html.parser')

        
        # Extract Tournament Name & Link
        tournament_element = soup.select_one(".Tournament_Header .Tournament_Link")
        if tournament_element:
            tournament_name = tournament_element.text.strip()
            tournament_link = tournament_element['href'] if 'href' in tournament_element.attrs else None
        else:
            tournament_name = "Unknown Tournament"
            tournament_link = None

        tournament_location = "Unknown Location"

        # If tournament link is found, open it in a new tab and scrape the location
        if tournament_link:
            tournament_url = f"https://www.espn.com{tournament_link}"
            driver.execute_script(f"window.open('{tournament_url}', '_blank');")  # Open in new tab
            driver.switch_to.window(driver.window_handles[-1])  # Switch to the new tab
            time.sleep(3)  # Wait for the page to load

            # Try to find the tournament location
            try:
                location_element = driver.find_element(By.CLASS_NAME, "TournamentHeader_Location")
                tournament_location = location_element.text.strip()
                print(f"Scraped Location for {tournament_name}: {tournament_location}")
            except:
                print(f"Could not find location for {tournament_name}")

            # Store location in dictionary
            tournament_locations[tournament_name] = tournament_location

            driver.close()  # Close the tab
            driver.switch_to.window(driver.window_handles[0])  # Switch back to main tab


        groupings = soup.select(".Grouping")

        for grouping in groupings:
            grouping_name_el = grouping.select_one(".Grouping_Name")
            if not grouping_name_el:
                continue

            grouping_name = grouping_name_el.text.strip()
            gender = "Man" if "Men" in grouping_name else "Woman"

            match_lists = grouping.select('ul.VZTD.rEPuv.dAmzA')
            if not match_lists:
                continue

            tournament_data = []

            print(f"Processing matches for Group: {grouping_name}")
            if "Singles" not in grouping_name:
                continue  # Skip if it's not a singles event

            gender = 'Man' if 'Men' in grouping_name else 'Woman'

            for match_idx, match_list in enumerate(match_lists):
                print(f"Processing match {match_idx+1}...")

                # Extract Player Names
                player_2 = match_list.find('a', {'class': 'TXMzn UbGlr ibBnq qdXbA WCDhQ DbOXS tqUtK GpWVU iJYzE cgHdO xTell GpQCA tuAKv spGOb'})
                player_1 = match_list.find('a', {'class': 'uQOvX UbGlr ibBnq qdXbA WCDhQ DbOXS tqUtK GpWVU iJYzE cgHdO xTell GpQCA tuAKv spGOb'})

                player_1_name = player_1.text if player_1 else 'Unknown Player 1'
                player_2_name = player_2.text if player_2 else 'Unknown Player 2'

                # Extract Player 1 Set Scores
                player_1_sets_div = match_list.select(
                    'li.VZTD.mLASH.dAmzA.oFFrS.lZur.bmjsw.koWyY.QXDKT.lXtzD:not(.ZRifP) '
                    '> div.VZTD.UlLJn.zkpVE.pgHdv.jasY.Ykpiq.DHRxp.ZrRMd.sOEJO '
                    '> div.FuEs.jasY.TXMzn.DHRxp.ZrRMd.VZTD.CLwPV.xTell.zkpVE.BHLB.kfeMl.CFZbp.BSXrm'
                )

                player_1_sets = []
                for div in player_1_sets_div:
                    spans = div.find_all('span', {'class': ['oXUjY ZdLIa', 'oXUjY ZdLIa nPLaK']})
                    for span in spans:
                        for sup in span.find_all('sup'):
                            sup.extract()
                        set_score = span.get_text(strip=True)
                        player_1_sets.append(set_score)

                # Extract Player 2 Set Scores
                player_2_sets_div = match_list.select(
                    'li.VZTD.mLASH.dAmzA.oFFrS.lZur.bmjsw.koWyY.QXDKT.lXtzD.ZRifP.LWFYs.JUvxK.qNQKc.vGxIM.fkNaB.QxNFt.WjJEf.XXaiR.tDWSI.QoUNa.dQXC.ZVbdw.iRjAJ.Hncps '
                    '> div.VZTD.UlLJn.zkpVE.pgHdv.jasY.Ykpiq.DHRxp.ZrRMd.sOEJO '
                    '> div.FuEs.jasY.TXMzn.DHRxp.ZrRMd.VZTD.CLwPV.xTell.zkpVE.BHLB.kfeMl.CFZbp.BSXrm'
                )

                player_2_sets = []
                for div in player_2_sets_div:
                    spans = div.find_all('span', {'class': ['oXUjY ZdLIa', 'oXUjY ZdLIa nPLaK']})
                    for span in spans:
                        for sup in span.find_all('sup'):
                            sup.extract()
                        set_score = span.get_text(strip=True)
                        player_2_sets.append(set_score)

                player_1_set_wins = 0
                player_2_set_wins = 0

                # Compare each set score
                for set_1, set_2 in zip(player_1_sets, player_2_sets):
                    try:
                        if int(set_1) > int(set_2):  # Player 1 wins this set
                            player_1_set_wins += 1
                        elif int(set_1) < int(set_2):  # Player 2 wins this set
                            player_2_set_wins += 1
                    except ValueError:
                        continue  # Skip if there is an invalid set score

                # Determine the winner
                if player_1_set_wins > player_2_set_wins:
                    winner = player_1_name
                elif player_1_set_wins < player_2_set_wins:
                    winner = player_2_name
                else:
                    winner = "Draw"  # In case of a tie (though this is rare in singles)

                # Format Set Scores
                player_1_score = '-'.join(player_1_sets) if player_1_sets else 'No sets'
                player_2_score = '-'.join(player_2_sets) if player_2_sets else 'No sets'

                # Ensure there are exactly 7 set scores, fill with None if necessary
                player_1_sets += [None] * (7 - len(player_1_sets))
                player_2_sets += [None] * (7 - len(player_2_sets))

                # Store the data for Player 1
                tournament_data.append([
                    current_date.strftime('%Y-%m-%d'),
                    tournament_name,
                    player_1_name,
                    tournament_locations.get(tournament_name, "Unknown Location"),
                    *player_1_sets, 
                    player_1_name,
                    gender
                ])

                # Store the data for Player 2
                tournament_data.append([
                    current_date.strftime('%Y-%m-%d'),
                    tournament_name,
                    player_2_name,
                    tournament_locations.get(tournament_name, "Unknown Location"),
                    *player_2_sets,
                    player_2_name,
                    gender
                ])

                # Add a blank row after each match (for both players)
                tournament_data.append([])
            
            # Append the tournament data to the main list
            data.extend(tournament_data)

    # Move to the next date
    current_date += timedelta(days=1)

# Close the WebDriver
driver.quit()

# Create DataFrame and save to Excel
columns = ['Date', 'Tournament Name', 'Player', 'Location', 'Set1', 'Set2', 'Set3', 'Set4', 'Set5', 'Set6', 'Set7', 'Player', 'Gender']
df = pd.DataFrame(data, columns=columns)

# Save to Excel
output_filename = 'tennis_matches.xlsx'
df.to_excel(output_filename, index=False)

print(f"Data has been saved to {output_filename}")