from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import pandas as pd
import time

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
    print(f"\nScraping data for {url}")

    driver.get(url)
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".DailyScoreboard_Tournaments section.Card"))
        )
        time.sleep(5)  # Wait for the page to load
    except:
        print("Timeout waiting for tournament cards. Skipping date.")
        current_date += timedelta(days=1)
        continue

    cards = driver.find_elements(By.CSS_SELECTOR, ".DailyScoreboard_Tournaments section.Card")
    print(f"Found {len(cards)} tournament cards on the page.")

    tournament_locations = {}

    for idx, card in enumerate(cards):
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, f".DailyScoreboard_Tournaments .Card:nth-of-type({idx+1}) ul.VZTD.rEPuv.dAmzA"))
            )
            temps = driver.find_elements(By.CSS_SELECTOR, ".DailyScoreboard_Tournaments section.Card")
            card = temps[idx]
            time.sleep(1) 
        except:
            print(f"Card {idx + 1} match list not fully loaded. Skipping.")
            continue

        card_html = card.get_attribute("innerHTML")
        soup = BeautifulSoup(card_html, 'html.parser')

        # Extract Tournament Name & Link
        tournament_element = soup.select_one(".Tournament_Header .Tournament_Link")
        tournament_name = tournament_element.text.strip() if tournament_element else "Unknown Tournament"
        tournament_link = tournament_element['href'] if tournament_element and 'href' in tournament_element.attrs else None
        tournament_location = "Unknown Location"

        # Visit tournament detail page
        if tournament_link:
            tournament_url = f"https://www.espn.com{tournament_link}"
            driver.execute_script(f"window.open('{tournament_url}', '_blank');")
            driver.switch_to.window(driver.window_handles[-1])
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "TournamentHeader_Location"))
                )
                time.sleep(3)  # Wait for the page to load
                location_element = driver.find_element(By.CLASS_NAME, "TournamentHeader_Location")
                tournament_location = location_element.text.strip()
                print(f"Scraped Location for {tournament_name}: {tournament_location}")
            except:
                print(f"Could not find location for {tournament_name}")
            tournament_locations[tournament_name] = tournament_location
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        groupings = soup.select(".Grouping")
        for grouping in groupings:
            grouping_name_el = grouping.select_one(".Grouping_Name")
            if not grouping_name_el:
                continue
            grouping_name = grouping_name_el.text.strip()

            if "Singles" not in grouping_name:
                continue  # Only Singles
            gender = "Man" if "Men" in grouping_name else "Woman"

            match_lists = grouping.select('ul.VZTD.rEPuv.dAmzA')
            if not match_lists:
                continue

            tournament_data = []

            for match_idx, match_list in enumerate(match_lists):
                print(f"Processing match {match_idx + 1} in {grouping_name}...")

                li_tags = match_list.find_all('li', limit=2)
                # Players
                player_1 = li_tags[0].find('a')
                player_2 = li_tags[1].find('a')
                player_1_name = player_1.text if player_1 else 'Unknown Player 1'
                player_2_name = player_2.text if player_2 else 'Unknown Player 2'

                # Set scores
                p1_divs = li_tags[0].select(
                    'div.FuEs.jasY.TXMzn.DHRxp.ZrRMd.VZTD.CLwPV.xTell.zkpVE.BHLB.kfeMl.BSXrm')
                p2_divs = li_tags[1].select(
                    'div.FuEs.jasY.TXMzn.DHRxp.ZrRMd.VZTD.CLwPV.xTell.zkpVE.BHLB.kfeMl.BSXrm')

                def extract_scores(divs):
                    sets = []
                    for div in divs:
                        for span in div.find_all('span', class_=['oXUjY ZdLIa', 'oXUjY ZdLIa nPLaK', 'oXUjY ZdLIa vuLkQ']):
                            for sup in span.find_all('sup'):
                                sup.extract()
                            sets.append(span.get_text(strip=True))
                    return sets

                player_1_sets = extract_scores(p1_divs)
                player_2_sets = extract_scores(p2_divs)

                # Determine winner
                p1_wins = sum(int(s1) > int(s2) for s1, s2 in zip(player_1_sets, player_2_sets) if s1.isdigit() and s2.isdigit())
                p2_wins = sum(int(s2) > int(s1) for s1, s2 in zip(player_1_sets, player_2_sets) if s1.isdigit() and s2.isdigit())
                winner = player_1_name if p1_wins > p2_wins else (player_2_name if p2_wins > p1_wins else "Draw")

                # Pad sets
                player_1_sets += [None] * (7 - len(player_1_sets))
                player_2_sets += [None] * (7 - len(player_2_sets))

                # Append rows
                tournament_data.append([
                    current_date.strftime('%Y-%m-%d'), tournament_name, player_1_name,
                    tournament_locations.get(tournament_name, "Unknown Location"),
                    *player_1_sets, player_1_name, gender
                ])
                tournament_data.append([
                    current_date.strftime('%Y-%m-%d'), tournament_name, player_2_name,
                    tournament_locations.get(tournament_name, "Unknown Location"),
                    *player_2_sets, player_2_name, gender
                ])
                tournament_data.append([])

            data.extend(tournament_data)

    current_date += timedelta(days=1)

# Close WebDriver
driver.quit()

# Save to Excel
columns = ['Date', 'Tournament Name', 'Player', 'Location', 'Set1', 'Set2', 'Set3', 'Set4', 'Set5', 'Set6', 'Set7', 'Player', 'Gender']
df = pd.DataFrame(data, columns=columns)
df.to_excel("tennis_matches.xlsx", index=False)
print("\nData has been saved to 'tennis_matches.xlsx'")
