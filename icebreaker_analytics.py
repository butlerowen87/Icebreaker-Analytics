from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import re
import requests


def setup_driver(headless=True):
    """Setup Chrome driver with options"""
    chrome_options = Options()
    if headless:
        chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')

    driver = webdriver.Chrome(options=chrome_options)
    return driver


def load_team_data(excel_file='Teams with URLs.xlsx'):
    """Load team data from Excel file"""
    df = pd.read_excel(excel_file)
    team_dict = {}
    for _, row in df.iterrows():
        initials = row['Initials'].strip().upper()
        team_dict[initials] = {
            'team_name': row['Team Name'],
            'edge_url': row['Edge URL'],
            'sv_pct_url': row['Team SV% URL'],
            'team_stats_url': row['Team Stats URL']
        }
    return team_dict


def scrape_nhl_edge(driver, url, team_name):
    """Scrape NHL Edge stats"""
    print(f"\nScraping NHL Edge for {team_name}...")
    data = {}

    try:
        driver.get(url)
        time.sleep(10)

        # Get games played
        try:
            gp_elems = driver.find_elements(By.CSS_SELECTOR, "div.sc-cnjBov.gtDprA")
            for elem in gp_elems[:5]:
                try:
                    parent = elem.find_element(By.XPATH, "./ancestor::div[2]")
                    parent_text = parent.text
                    if 'GP' in parent_text or 'Games' in parent_text:
                        data['games_played'] = elem.text.strip()
                        break
                except:
                    continue
        except:
            pass

        # Get offensive and defensive zone time
        try:
            oz_elements = driver.find_elements(By.CSS_SELECTOR, "div.sc-eEpesX.gDXYNW")
            # Pattern is: [defensive, neutral, offensive, ...]
            zone_times = []
            for elem in oz_elements:
                text = elem.text.strip()
                # Only collect elements that look like percentages or numbers
                if text and (('%' in text) or (text.replace('.', '').isdigit())):
                    zone_times.append(text)

            # Pattern: index 0=defensive, 1=neutral, 2=offensive
            if len(zone_times) >= 3:
                # First is defensive
                dz_text = zone_times[0]
                if '%' in dz_text:
                    data['defensive_zone_time_pct'] = dz_text.replace('%', '')
                else:
                    data['defensive_zone_time_pct'] = dz_text

                # Third is offensive (skip neutral at index 1)
                oz_text = zone_times[2]
                if '%' in oz_text:
                    data['offensive_zone_time_pct'] = oz_text.replace('%', '')
                else:
                    data['offensive_zone_time_pct'] = oz_text
        except:
            pass

        # Get high danger shots
        try:
            low_slot = None
            crease = None

            try:
                low_slot_elem = driver.find_element(By.ID, "Low Slot")
                nums = re.findall(r'\d+', low_slot_elem.text.strip())
                if nums:
                    low_slot = int(nums[0])
            except:
                pass

            try:
                crease_elem = driver.find_element(By.ID, "Crease")
                nums = re.findall(r'\d+', crease_elem.text.strip())
                if nums:
                    crease = int(nums[0])
            except:
                pass

            if low_slot is None or crease is None:
                all_text = driver.find_elements(By.XPATH,
                                                "//*[contains(text(), 'Low Slot') or contains(text(), 'Crease')]")
                for elem in all_text:
                    try:
                        parent = elem.find_element(By.XPATH, "..")
                        nums = re.findall(r'\d+', parent.text)
                        if 'Low Slot' in elem.text and nums and low_slot is None:
                            low_slot = int(nums[0])
                        elif 'Crease' in elem.text and nums and crease is None:
                            crease = int(nums[0])
                    except:
                        continue

            if low_slot is not None and crease is not None:
                data['high_danger_shots'] = str(low_slot + crease)
        except:
            pass

        print(f"  Games Played: {data.get('games_played', 'Not found')}")
        print(f"  High Danger Shots: {data.get('high_danger_shots', 'Not found')}")
        print(f"  Offensive Zone Time %: {data.get('offensive_zone_time_pct', 'Not found')}")
        print(f"  Defensive Zone Time %: {data.get('defensive_zone_time_pct', 'Not found')}")

    except Exception as e:
        print(f"  Error: {e}")

    return data


def scrape_save_percentage(driver, url, team_name):
    """Scrape Team Save Percentage"""
    print(f"\nScraping Save Percentage for {team_name}...")
    data = {}

    try:
        driver.get(url)
        time.sleep(5)
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

        try:
            sv_elem = driver.find_element(By.CSS_SELECTOR, "td.sc-fylBCY.cCvCop.rt-td.sorted.sorted-0.sorted-desc")
            sv_text = sv_elem.text.strip()
            if sv_text.startswith('.') or sv_text.startswith('0.'):
                data['team_sv_pct'] = sv_text
        except:
            rows = driver.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) > 0:
                    for cell in cells:
                        text = cell.text.strip()
                        if text.startswith('.') and len(text) <= 5:
                            try:
                                val = float(text)
                                if 0.800 <= val <= 0.950:
                                    data['team_sv_pct'] = text
                                    break
                            except:
                                pass
                    if 'team_sv_pct' in data:
                        break

        print(f"  Team SV%: {data.get('team_sv_pct', 'Not found')}")

    except Exception as e:
        print(f"  Error: {e}")

    return data


def scrape_team_stats(driver, url, team_name):
    """Scrape GF/GP, GA/GP, PP%, PK%, Shots/GP"""
    print(f"\nScraping Team Stats for {team_name}...")
    data = {}

    try:
        driver.get(url)
        time.sleep(5)
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

        headers = driver.find_elements(By.TAG_NAME, "th")
        header_names = [h.text.strip() for h in headers]

        indices = {}
        for i, header in enumerate(header_names):
            if 'GF/GP' in header:
                indices['gf_gp'] = i
            elif 'GA/GP' in header:
                indices['ga_gp'] = i
            elif 'PP%' in header:
                indices['pp_pct'] = i
            elif 'PK%' in header:
                indices['pk_pct'] = i
            elif 'S/GP' in header or 'Shots/GP' in header:
                indices['shots_gp'] = i

        rows = driver.find_elements(By.TAG_NAME, "tr")
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) > 0:
                if 'gf_gp' in indices and len(cells) > indices['gf_gp']:
                    data['gf_gp'] = cells[indices['gf_gp']].text.strip()
                if 'ga_gp' in indices and len(cells) > indices['ga_gp']:
                    data['ga_gp'] = cells[indices['ga_gp']].text.strip()
                if 'pp_pct' in indices and len(cells) > indices['pp_pct']:
                    data['pp_pct'] = cells[indices['pp_pct']].text.strip()
                if 'pk_pct' in indices and len(cells) > indices['pk_pct']:
                    data['pk_pct'] = cells[indices['pk_pct']].text.strip()
                if 'shots_gp' in indices and len(cells) > indices['shots_gp']:
                    data['shots_gp'] = cells[indices['shots_gp']].text.strip()
                break

        print(f"  GF/GP: {data.get('gf_gp', 'Not found')}")
        print(f"  GA/GP: {data.get('ga_gp', 'Not found')}")
        print(f"  PP%: {data.get('pp_pct', 'Not found')}")
        print(f"  PK%: {data.get('pk_pct', 'Not found')}")
        print(f"  Shots/GP: {data.get('shots_gp', 'Not found')}")

    except Exception as e:
        print(f"  Error: {e}")

    return data


def scrape_wins_from_nhl(driver, team_name, team_initials):
    """Scrape Wins from NHL.com stats"""
    print(f"\nScraping Wins from NHL.com for {team_name}...")
    data = {}

    try:
        url = 'https://www.nhl.com/stats/teams?reportType=season&seasonFrom=20252026&seasonTo=20252026&gameType=2&sort=points,wins&page=0&pageSize=50'
        driver.get(url)
        time.sleep(6)

        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

        headers = driver.find_elements(By.TAG_NAME, "th")
        header_names = [h.text.strip() for h in headers]

        wins_index = None
        for i, header in enumerate(header_names):
            if header == 'W' or header == 'Wins':
                wins_index = i
                break

        rows = driver.find_elements(By.TAG_NAME, "tr")
        for row in rows:
            if team_initials in row.text or team_name.split()[0] in row.text:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) > 0 and wins_index is not None and len(cells) > wins_index:
                    wins_text = cells[wins_index].text.strip()
                    if wins_text.isdigit():
                        data['wins'] = wins_text
                break

        print(f"  Wins: {data.get('wins', 'Not found')}")

    except Exception as e:
        print(f"  Error: {e}")

    return data


def scrape_l10_from_espn(team_name, team_initials):
    """Scrape L10 wins from ESPN API"""
    print(f"\nScraping L10 from ESPN for {team_name}...")
    data = {}

    try:
        url = "https://site.web.api.espn.com/apis/v2/sports/hockey/nhl/standings"
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        standings = resp.json()

        for group in standings["children"]:
            if "standings" not in group:
                continue
            if "entries" not in group["standings"]:
                continue

            for entry in group["standings"]["entries"]:
                team = entry.get("team", {})
                team_abbr = team.get("abbreviation", "")

                if team_abbr == team_initials:
                    stats = entry.get("stats", [])

                    for stat in stats:
                        stat_name = stat.get("name", "")
                        # The stat is called "Last Ten Games" not "lastTen"
                        if stat_name == "Last Ten Games":
                            record = stat.get("displayValue", "")
                            if record and "-" in record:
                                wins = int(record.split("-")[0])
                                data["l10_wins"] = str(wins)
                                print(f"  L10 record: {record}, Wins: {wins}")
                                return data

                    print(f"  'Last Ten Games' stat not found")
                    return data

        print(f"  Team '{team_initials}' not found in API")

    except Exception as e:
        print(f"  Error: {e}")

    return data


def scrape_team(team_initials, team_data_dict):
    """Main function to scrape all stats for a specific team"""
    team_initials = team_initials.upper()
    if team_initials not in team_data_dict:
        print(f"Error: Team '{team_initials}' not found!")
        print(f"Available teams: {', '.join(sorted(team_data_dict.keys()))}")
        return None

    team_info = team_data_dict[team_initials]
    team_name = team_info['team_name']

    print("=" * 60)
    print(f"NHL Data Scraper - {team_name} ({team_initials})")
    print("=" * 60)

    driver = setup_driver()
    all_data = {'team': team_name, 'initials': team_initials}

    try:
        all_data.update(scrape_nhl_edge(driver, team_info['edge_url'], team_name))
        all_data.update(scrape_save_percentage(driver, team_info['sv_pct_url'], team_name))
        all_data.update(scrape_team_stats(driver, team_info['team_stats_url'], team_name))
        all_data.update(scrape_wins_from_nhl(driver, team_name, team_initials))
        all_data.update(scrape_l10_from_espn(team_name, team_initials))
    finally:
        driver.quit()

    print("\n" + "=" * 60)
    print("FINAL RESULTS")
    print("=" * 60)
    print(f"Team: {all_data.get('team')}")
    print(f"Initials: {all_data.get('initials')}")
    print(f"Games Played: {all_data.get('games_played', 'N/A')}")
    print(f"Wins: {all_data.get('wins', 'N/A')}")
    print(f"High Danger Shots: {all_data.get('high_danger_shots', 'N/A')}")
    print(f"Offensive Zone Time %: {all_data.get('offensive_zone_time_pct', 'N/A')}")
    print(f"Defensive Zone Time %: {all_data.get('defensive_zone_time_pct', 'N/A')}")
    print(f"Team SV%: {all_data.get('team_sv_pct', 'N/A')}")
    print(f"GF/GP: {all_data.get('gf_gp', 'N/A')}")
    print(f"GA/GP: {all_data.get('ga_gp', 'N/A')}")
    print(f"PP%: {all_data.get('pp_pct', 'N/A')}")
    print(f"PK%: {all_data.get('pk_pct', 'N/A')}")
    print(f"Shots/GP: {all_data.get('shots_gp', 'N/A')}")
    print(f"L10 Wins: {all_data.get('l10_wins', 'N/A')}")

    return all_data


def calculate_team_score(team_data):
    """Calculate score using the provided formula"""
    try:
        # Extract and convert variables
        GP = float(team_data.get('games_played', 0))
        W = float(team_data.get('wins', 0))
        LG = float(team_data.get('l10_wins', 0))

        # Convert percentages (from "29.2" to 0.292)
        PP = float(team_data.get('pp_pct', 0)) / 100
        PK = float(team_data.get('pk_pct', 0)) / 100

        S = float(team_data.get('shots_gp', 0))
        GF = float(team_data.get('gf_gp', 0))
        GA = float(team_data.get('ga_gp', 0))
        A = float(team_data.get('high_danger_shots', 0))

        # Convert zone times (from "41.3" to 0.413)
        OZ = float(team_data.get('offensive_zone_time_pct', 0)) / 100
        DZ = float(team_data.get('defensive_zone_time_pct', 0)) / 100

        # Convert save percentage (from ".872" to 0.872)
        SV_str = team_data.get('team_sv_pct', '0')
        if SV_str.startswith('.'):
            SV = float(SV_str)
        else:
            SV = float(SV_str)

        # Calculate score using the formula
        score = (((W / GP) / 0.6) + (2.6 * ((LG / 10) / 0.6)) + (1.2 * (PP / 0.23)) + (2.75 * (PK / 0.8)) +
                 (2.5 * (S / 30)) + (1.85 * (GF / GA)) + (2.5 * ((A / GP) / 6.5)) + (1.75 * (OZ / DZ)) +
                 (1.7 * (SV / 0.9)))

        return round(score, 2)

    except Exception as e:
        print(f"Error calculating score: {e}")
        return 0.0


def compare_teams(team1_initials, team2_initials, team_data_dict):
    """Compare two teams and return winner"""
    print("=" * 60)
    print("NHL TEAM COMPARISON")
    print("=" * 60)

    # Scrape both teams
    print(f"\n>>> SCRAPING TEAM 1: {team1_initials}")
    team1_data = scrape_team(team1_initials, team_data_dict)

    if team1_data is None:
        return None

    print(f"\n>>> SCRAPING TEAM 2: {team2_initials}")
    team2_data = scrape_team(team2_initials, team_data_dict)

    if team2_data is None:
        return None

    # Calculate scores
    team1_score = calculate_team_score(team1_data)
    team2_score = calculate_team_score(team2_data)

    # Determine winner
    print("\n" + "=" * 60)
    print("COMPARISON RESULTS")
    print("=" * 60)
    print(f"\n{team1_data['team']} ({team1_initials}): {team1_score}")
    print(f"{team2_data['team']} ({team2_initials}): {team2_score}")

    if team1_score > team2_score:
        winner = team1_data['team']
        winner_initials = team1_initials
        winner_score = team1_score
        loser_initials = team2_initials
        loser_score = team2_score
        T1 = team1_score
        T2 = team2_score
    else:
        winner = team2_data['team']
        winner_initials = team2_initials
        winner_score = team2_score
        loser_initials = team1_initials
        loser_score = team1_score
        T1 = team2_score
        T2 = team1_score

    # Calculate new formula: ((T1 / T2) * 100) - 50
    if T1 >= T2:
        final_score = ((T1 / T2) * 100) - 50
        final_score = round(final_score, 2)
    else:
        final_score = ((T2 / T1) * 100) - 50
        final_score = round(final_score, 2)

    print("\n" + "=" * 60)
    print(f"🏆 WINNER: {winner} ({winner_initials})")
    print(f"   Score: {winner_score}")
    print(f"\n   Opponent: {loser_initials} - Score: {loser_score}")
    print(f"   Difference: +{round(winner_score - loser_score, 2)}")
    print(f"\n   Win Probability: {final_score}%")
    print("=" * 60)

    return {
        'winner': winner,
        'winner_initials': winner_initials,
        'winner_score': winner_score,
        'loser_initials': loser_initials,
        'loser_score': loser_score,
        'final_score': final_score
    }


def main():
    """Main function with user input"""
    print("Loading team data from Excel file...")
    team_data = load_team_data('Teams with URLs.xlsx')
    print(f"Loaded {len(team_data)} teams successfully!\n")

    print("Enter two teams to compare:")
    team1_initials = input("Team 1 initials (e.g., VGK, TOR, MTL): ").strip().upper()
    team2_initials = input("Team 2 initials (e.g., EDM, BOS, MTL): ").strip().upper()

    # Compare teams
    results = compare_teams(team1_initials, team2_initials, team_data)

    return results


if __name__ == "__main__":
    main()