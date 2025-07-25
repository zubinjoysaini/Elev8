#!/usr/bin/env python
# coding: utf-8

# In[4]:


import pandas as pd
import random
import os
import datetime
import glob
import itertools 
from collections import defaultdict

# ===== Team class =====
class Team:
    def __init__(self, name):
        self.name = name
        self.points = 0
        self.matches_played = 0
        self.wins = 0
        self.losses = 0
        self.extra_time_losses = 0
        self.extra_time_wins = 0
        self.points_difference = 0
        self.opponents = set()

    def add_result(self, opponent, diff):
        self.opponents.add(opponent)
        self.matches_played += 1
        self.points_difference += diff
        if diff > 7:
            self.points += 6
            self.wins += 1
        elif 1 <= diff <= 7:
            self.points += 5
            self.wins += 1
        elif diff == 0:
            self.points += 4
            self.extra_time_wins += 1

    def add_loss(self, opponent, diff):
        self.opponents.add(opponent)
        self.matches_played += 1
        self.points_difference -= diff
        if diff > 7:
            self.losses += 1
            self.points += 0
        elif 1 <= diff <= 7:
            self.losses += 1
            self.points += 1
        elif diff == 0:
            self.extra_time_losses += 1
            self.points += 2

# ===== Helper functions =====
def calculate_difference_and_points(matches_df):
    for i, row in matches_df.iterrows():
        if pd.notna(row["Points_Winner"]) and pd.notna(row["Points_Loser"]):
            try:
                pw = int(row["Points_Winner"])
                pl = int(row["Points_Loser"])
                matches_df.at[i, "Difference"] = pw - pl
            except ValueError:
                continue
    return matches_df

def calculate_standings(matches_df):
    teams = defaultdict(Team)
    for _, row in matches_df.iterrows():
        if pd.isna(row["Winner"]) or pd.isna(row["Loser"]):
            continue
        try:
            diff = int(row["Difference"])
        except (ValueError, TypeError):
            continue
        winner = row["Winner"]
        loser = row["Loser"]

        if winner not in teams:
            teams[winner] = Team(winner)
        if loser not in teams:
            teams[loser] = Team(loser)

        teams[winner].add_result(loser, diff)
        teams[loser].add_loss(winner, diff)

    return teams

def swiss_pairing(teams_dict):
    teams = list(teams_dict.values())
    sorted_teams = sorted(teams, key=lambda t: (-t.points, -t.wins, -t.points_difference, t.name))
    team_names = [t.name for t in sorted_teams]

    # Create a dictionary of previous matches
    previous_matches = {t.name: t.opponents for t in teams}

    # Try all possible pairings of the team list
    def is_valid_pairing(pairing):
        for t1, t2 in pairing:
            if t2 in previous_matches[t1]:
                return False
        return True

    def generate_pairings(teams_left):
        if not teams_left:
            return []
        t1 = teams_left[0]
        for i in range(1, len(teams_left)):
            t2 = teams_left[i]
            if t2 not in previous_matches[t1]:
                rest = teams_left[1:i] + teams_left[i+1:]
                sub_pairing = generate_pairings(rest)
                if sub_pairing is not None:
                    return [(t1, t2)] + sub_pairing
        return None  # No valid pairing found

    pairings = generate_pairings(team_names)

    if pairings is None:
        raise Exception("❌ Could not find a valid set of pairings without rematches. Try reducing rounds or check constraints.")

    return pairings

def generate_leaderboard(teams):
    leaderboard_rows = []

    # Calculate opponent points for each team
    opponent_points = {}
    for t in teams.values():
        total_opponent_points = sum(
            [teams[opp].points for opp in t.opponents if opp in teams]
        )
        opponent_points[t.name] = total_opponent_points

    # Build leaderboard rows
    for t in sorted(teams.values(), key=lambda t: (-t.points, -t.wins, -t.points_difference, t.name)):
        leaderboard_rows.append({
            "Team": t.name,
            "Points": t.points,
            "Played": t.matches_played,
            "Wins": t.wins,
            "Losses": t.losses,
            "ET Wins": t.extra_time_wins,
            "ET Losses": t.extra_time_losses,
            "Points Diff": t.points_difference,
            "Opponent Pts Sum": opponent_points.get(t.name, 0)
        })

    return pd.DataFrame(leaderboard_rows)

def safe_excel_write(matches_df, settings_df, leaderboard_df, base_path):
    try:
        round_num = int(matches_df["Round"].max())
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        new_filename = base_path.replace(".xlsx", f"_round{round_num}_{timestamp}.xlsx")

        writer = pd.ExcelWriter(new_filename, engine='openpyxl')
        matches_df.to_excel(writer, sheet_name="Matches", index=False)
        settings_df.to_excel(writer, sheet_name="Settings", index=False)
        leaderboard_df.to_excel(writer, sheet_name="Leaderboard", index=False)
        writer.close()

        print(f"✅ Round {round_num} saved to: {new_filename}")
        return True
    except Exception as e:
        print(f"❌ Error writing file: {e}")
        return False

def update_next_round(filepath):
    try:
        xls = pd.ExcelFile(filepath)
        settings_df = pd.read_excel(xls, sheet_name="Settings")
        matches_df = pd.read_excel(xls, sheet_name="Matches")
    except Exception as e:
        print("❌ Error reading the Excel file:", e)
        return

    num_teams = int(settings_df.loc[settings_df["Parameter"] == "Number of Teams", "Value"].values[0])
    num_rounds = int(settings_df.loc[settings_df["Parameter"] == "Number of Rounds", "Value"].values[0])

    matches_df = calculate_difference_and_points(matches_df)

    if matches_df.empty:
        current_round = 0
    else:
        current_round = matches_df["Round"].max()
        incomplete = matches_df[matches_df["Round"] == current_round][["Winner", "Loser", "Points_Winner", "Points_Loser"]].isnull().any(axis=1).sum()
        if incomplete > 0:
            print(f"⚠️ Please complete all results for Round {current_round} before generating the next round.")
            return

    if current_round >= num_rounds:
       print("🏁 Tournament is complete.")
       leaderboard_df = generate_leaderboard(calculate_standings(matches_df))
       safe_excel_write(matches_df, settings_df, leaderboard_df, filepath)
       return

    round_number = current_round + 1

    if round_number == 1:
        team_names = []

        # Try reading from column 'Team Names' in the same row as 'Number of Teams'
        if "Team Names" in settings_df.columns:
            team_row = settings_df[settings_df["Parameter"] == "Number of Teams"]
            if not team_row.empty and pd.notna(team_row["Team Names"].values[0]):
                raw_names = team_row["Team Names"].values[0]
                team_names = [name.strip() for name in raw_names.split(",") if name.strip()]

        if not team_names:
            print("⚠️ No valid team names found in 'Team Names' column. Using default names.")
            team_names = [f"Team {i+1}" for i in range(num_teams)]

        if len(team_names) != num_teams:
            print("⚠️ Mismatch: 'Number of Teams' is", num_teams, "but", len(team_names), "names found.")
            return

        random.shuffle(team_names)
        pairings = [(team_names[i], team_names[i+1]) if i+1 < len(team_names) else (team_names[i], "BYE")
                    for i in range(0, len(team_names), 2)]
    else:
        teams = calculate_standings(matches_df)
        pairings = swiss_pairing(teams)

    new_rows = []
    for t1, t2 in pairings:
        new_rows.append({
            "Round": round_number,
            "Team 1": t1,
            "Team 2": t2,
            "Winner": "", "Loser": "",
            "Points_Winner": "", "Points_Loser": "",
            "Difference": ""
        })

    updated_matches_df = pd.concat([matches_df, pd.DataFrame(new_rows)], ignore_index=True)

    if round_number > 1:
        leaderboard_df = generate_leaderboard(calculate_standings(updated_matches_df))
    else:
        leaderboard_df = pd.DataFrame(columns=["Team", "Points", "Played", "Wins", "Losses", "ET Wins", "ET Losses", "Points Diff"])

    safe_excel_write(updated_matches_df, settings_df, leaderboard_df, filepath)

def get_latest_round_file(base_path):
    base_dir = os.path.dirname(base_path)
    base_name = os.path.basename(base_path).replace('.xlsx', '')
    files = glob.glob(os.path.join(base_dir, f'{base_name}*.xlsx'))
    if not files:
        return base_path
    latest = max(files, key=os.path.getmtime)
    return latest

# === RUN ===
base_template_path = "C:\\Users\\Zubin\\Downloads\\kabaddi_swiss_dynamic_template.xlsx"
latest_file = get_latest_round_file(base_template_path)
update_next_round(latest_file)


# In[1]:


import pandas as pd

def generate_generalized_template(filename="C:\\Users\\Zubin\\Downloads\\kabaddi_swiss_dynamic_template.xlsx", default_teams=30, default_rounds=4):
    # Settings sheet
    settings_df = pd.DataFrame({
        "Parameter": ["Number of Teams", "Number of Rounds"],
        "Value": [default_teams, default_rounds],
        "Team Names": ["Team 1, Team 2, Team 3, Team 4, Team 5, Team 6, Team 7, Team 8", ""]  # Sample default
    })

    # Empty Matches sheet
    matches_df = pd.DataFrame(columns=[
        "Round", "Team 1", "Team 2", 
        "Winner", "Loser", 
        "Points_Winner", "Points_Loser", "Difference"
    ])

    # Write Excel
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        matches_df.to_excel(writer, sheet_name="Matches", index=False)
        settings_df.to_excel(writer, sheet_name="Settings", index=False)

    print(f"✅ Excel template created at: {filename}")

# Usage
generate_generalized_template()


# In[ ]:





# In[ ]:




