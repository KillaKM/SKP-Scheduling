import copy
from datetime import datetime, timedelta, time
import io
import random
import re
import openpyxl
from openpyxl.cell.cell import MergedCell
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="SKP Schedule & Task Automation", layout="wide"
)

# --- WACHTWOORDBEVEILIGING ---
def check_password():
    def password_entered():
        if st.session_state["password"] == "Tantalus2627!":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 Inloggen vereist")
        st.info("Voer het wachtwoord in om toegang te krijgen tot het SKP Taaksysteem.")
        st.text_input(
            "Wachtwoord:",
            type="password",
            on_change=password_entered,
            key="password",
        )
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 Inloggen vereist")
        st.text_input(
            "Wachtwoord:",
            type="password",
            on_change=password_entered,
            key="password",
        )
        st.error("Onjuist wachtwoord. Probeer het opnieuw.")
        return False
    return True


if not check_password():
    st.stop()
# ------------------------------

st.title("🏀 SKP Taakindeling & Scheidsrechters Systeem")

uploaded_file = st.sidebar.file_uploader(
    "Upload je Excel-bestand", type=["xlsx"], key="main_file_uploader"
)

LOCK_ALL_COL = "Lock All"
LOCK_COLS = ["Lock Ref 1", "Lock Ref 2", "Lock Scorer", "Lock Timer", "Lock 24s"]
ALL_LOCK_COLS = [LOCK_ALL_COL] + LOCK_COLS
TASK_COLS = ["Referee 1", "Referee 2", "Scorer", "Timer", "24 sec operator"]
LOCK_MAP = dict(zip(TASK_COLS, LOCK_COLS))

DUTCH_DAYS = {
    0: "ma",
    1: "di",
    2: "wo",
    3: "do",
    4: "vr",
    5: "za",
    6: "zo"
}


def ensure_lock_columns(df):
    df_res = df.copy()
    for col in ALL_LOCK_COLS:
        if col not in df_res.columns:
            df_res[col] = False
        else:
            df_res[col] = df_res[col].fillna(False).astype(bool)

    all_locked = df_res[LOCK_COLS].all(axis=1)
    df_res[LOCK_ALL_COL] = df_res[LOCK_ALL_COL] | all_locked
    return df_res


def normalize_time_str(val):
    if pd.isna(val) or val is None or str(val).strip() in ["", "nan", "None"]:
        return "00:00"
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%H:%M")
    if isinstance(val, time):
        return val.strftime("%H:%M")
    s = str(val).strip()
    m = re.search(r"(\d{1,2}):(\d{2})", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return s


def format_time_columns_in_df(df):
    df_clean = df.copy()
    for col in df_clean.columns:
        c_lower = str(col).strip().lower()
        if c_lower in ["time", "tijd", "starttijd", "start time", "eindtijd", "end time"]:
            df_clean[col] = df_clean[col].apply(lambda v: normalize_time_str(v) if str(v).strip() not in ["", "nan", "None"] else "")
    return df_clean


def make_arrow_compatible(df):
    df_clean = df.copy()
    for col in df_clean.columns:
        if col in ALL_LOCK_COLS:
            df_clean[col] = df_clean[col].fillna(False).astype(bool)
        elif df_clean[col].dtype == object or str(col) in [
            "Referee 1",
            "Referee 2",
            "Scorer",
            "Timer",
            "24 sec operator",
            "Home Team",
            "Away Team",
            "Division",
            "Court",
            "Veld",
            "Dag",
            "Day",
            "Date",
            "Time",
            "Starttijd"
        ]:
            df_clean[col] = df_clean[col].fillna("").astype(str)
            df_clean[col] = df_clean[col].replace(
                {"nan": "", "None": "", "NoneType": ""}
            )
    return df_clean


def find_sheet(sheets_dict, candidates):
    for cand in candidates:
        for name in sheets_dict.keys():
            if cand.lower() == str(name).strip().lower():
                return name
    for cand in candidates:
        for name in sheets_dict.keys():
            if cand.lower() in str(name).strip().lower():
                return name
    return None


def find_col(df, candidates, fallback_index=None):
    for col in df.columns:
        c_clean = str(col).strip().lower().replace("_", " ").replace("-", " ")
        for cand in candidates:
            if cand.lower() == c_clean:
                return col
    for col in df.columns:
        c_clean = str(col).strip().lower()
        for cand in candidates:
            if cand.lower() in c_clean:
                return col
    if fallback_index is not None and len(df.columns) > fallback_index:
        return df.columns[fallback_index]
    return None


def clean_team_code(team_name):
    if not team_name or pd.isna(team_name):
        return ""
    t = (
        str(team_name)
        .lower()
        .replace("tantalus", "")
        .replace("-", " ")
        .strip()
    )
    return "".join(t.split())


def build_division_map(divisions_df):
    div_map = {}
    if divisions_df is None or divisions_df.empty:
        return div_map
    df = divisions_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    team_col = find_col(
        df, ["team", "tantalus team", "teams", "teamnaam"], fallback_index=0
    )
    div_col = find_col(
        df,
        ["division", "divisie", "klasse", "poule"],
        fallback_index=1 if len(df.columns) > 1 else 0,
    )

    for _, row in df.dropna(subset=[team_col]).iterrows():
        t_val = str(row[team_col]).strip()
        d_val = str(row[div_col]).strip()
        num = 5
        for d in ["1", "2", "3", "4", "5", "6"]:
            if d in d_val.lower():
                num = int(d)
                break
        div_map[clean_team_code(t_val)] = num
    return div_map


def determine_division_for_team(team_str, div_map):
    if not team_str or str(team_str).strip() in ["", "nan", "None"]:
        return 5
    code = clean_team_code(team_str)
    if code in div_map:
        return div_map[code]
    for k, v in div_map.items():
        if k and (k in code or code in k):
            return v
    m = re.search(r"(?:mse|vse|xse|u\d+)[\s\-]*([1-9])", str(team_str).lower())
    if m:
        return int(m.group(1))
    return 5


def parse_date_obj(val):
    if pd.isna(val) or val is None or str(val).strip() in ["", "nan", "None"]:
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.date()

    try:
        dt = pd.to_datetime(val, dayfirst=True)
        return dt.date()
    except Exception:
        pass

    s = str(val).strip()
    m = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})", s)
    if m:
        d, m_val, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        try:
            return datetime(y, m_val, d).date()
        except ValueError:
            pass

    m2 = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m2:
        y, m_val, d = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        try:
            return datetime(y, m_val, d).date()
        except ValueError:
            pass

    return None


def normalize_date_str(val):
    d_obj = parse_date_obj(val)
    if d_obj:
        return d_obj.strftime("%Y-%m-%d")
    s = str(val).strip().lower()
    m = re.search(r"(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})", s)
    if m:
        return m.group(1)
    return s


def calculate_comm_points(comm_str, comm_points_map):
    if not comm_str or pd.isna(comm_str) or str(comm_str).strip() in ["", "nan", "None"]:
        return 0.0
    comms = [c.strip() for c in str(comm_str).split(",") if c.strip()]
    return sum(comm_points_map.get(c, 0.0) for c in comms)


def is_player_eligible_for_date(season_type, date_val):
    if not season_type:
        return True

    st_clean = str(season_type).strip().lower()

    if any(k in st_clean for k in ["full", "heel", "vol", "beide", "all"]):
        return True

    d_obj = parse_date_obj(date_val)
    if not d_obj:
        return False

    is_first_half = d_obj.month in [8, 9, 10, 11, 12, 1]

    if any(k in st_clean for k in ["1st", "1e", "first", "eerste", "1st half", "first half"]):
        return is_first_half

    if any(k in st_clean for k in ["2nd", "2e", "second", "tweede", "2nd half", "second half"]):
        return not is_first_half

    return True


def standardize_players_df(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_last = find_col(df, ["last name", "lastname", "achternaam", "last"], fallback_index=1 if len(df.columns) > 1 else None)
    col_first = find_col(df, ["first name", "firstname", "voornaam", "first"], fallback_index=0 if len(df.columns) > 0 else None)
    col_dip = find_col(df, ["diploma", "licentie", "certificaat", "niveau"], fallback_index=2 if len(df.columns) > 2 else None)
    col_comm = find_col(df, ["committee", "commissie"])
    col_extra = find_col(df, ["extra", "opmerking", "status", "notes"])
    col_team = find_col(df, ["team", "teamnaam", "spelend team"])
    col_season = find_col(df, ["full/ half season", "full/half season", "season", "seizoen", "half season", "seizoenshelft", "half"])
    col_extra_pts = find_col(df, ["extra points", "extra punten", "commissie punten", "comm points"])
    col_total_pts = find_col(df, ["total points", "totaal punten", "punten"])

    renames = {}
    if col_first and col_first != "First name":
        renames[col_first] = "First name"
    if col_last and col_last != "Last name":
        renames[col_last] = "Last name"
    if col_dip and col_dip != "Diploma":
        renames[col_dip] = "Diploma"
    if col_comm and col_comm != "Committee":
        renames[col_comm] = "Committee"
    if col_extra and col_extra != "Extra":
        renames[col_extra] = "Extra"
    if col_team and col_team != "Team":
        renames[col_team] = "Team"
    if col_season and col_season != "Full/ half season":
        renames[col_season] = "Full/ half season"
    if col_extra_pts and col_extra_pts != "Extra points":
        renames[col_extra_pts] = "Extra points"
    if col_total_pts and col_total_pts != "Total points":
        renames[col_total_pts] = "Total points"

    if renames:
        df = df.rename(columns=renames)

    if "First name" not in df.columns and len(df.columns) > 0:
        df["First name"] = df.iloc[:, 0]
    if "Last name" not in df.columns and len(df.columns) > 1:
        df["Last name"] = df.iloc[:, 1]
    if "Diploma" not in df.columns:
        df["Diploma"] = ""
    if "Committee" not in df.columns:
        df["Committee"] = ""
    if "Extra" not in df.columns:
        df["Extra"] = ""
    if "Team" not in df.columns:
        df["Team"] = ""
    if "Full/ half season" not in df.columns:
        df["Full/ half season"] = "Full season"

    if "Extra points" not in df.columns:
        df["Extra points"] = 0.0
    else:
        df["Extra points"] = pd.to_numeric(df["Extra points"], errors="coerce").fillna(0.0)

    if "Referee" not in df.columns:
        df["Referee"] = 0
    if "Table duty" not in df.columns:
        df["Table duty"] = 0
    if "Total points" not in df.columns:
        df["Total points"] = df["Extra points"]
    else:
        df["Total points"] = pd.to_numeric(df["Total points"], errors="coerce").fillna(df["Extra points"])

    return df


def standardize_skp_df(df, div_map=None):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_home = find_col(df, ["home team", "thuis team", "home", "thuis"])
    col_away = find_col(df, ["away team", "uit team", "away", "uit"])
    col_date = find_col(df, ["date", "datum"])
    col_time = find_col(df, ["time", "tijd"])
    col_div = find_col(df, ["division", "divisie", "poule", "klasse", "league", "div"])
    col_court = find_col(df, ["court", "veld", "zaal", "field"])
    col_dag = find_col(df, ["dag", "day"])

    renames = {}
    if col_home and col_home != "Home Team":
        renames[col_home] = "Home Team"
    if col_away and col_away != "Away Team":
        renames[col_away] = "Away Team"
    if col_date and col_date != "Date":
        renames[col_date] = "Date"
    if col_time and col_time != "Time":
        renames[col_time] = "Time"
    if col_div and col_div != "Division":
        renames[col_div] = "Division"
    if col_court and col_court not in ["Court", "Veld"]:
        renames[col_court] = "Court"
    if col_dag and col_dag not in ["Dag", "Day"]:
        renames[col_dag] = "Dag"

    if renames:
        df = df.rename(columns=renames)

    court_col = "Court" if "Court" in df.columns else ("Veld" if "Veld" in df.columns else None)
    if not court_col:
        df["Court"] = ""

    if div_map is not None:
        div_vals = []
        for _, r in df.iterrows():
            h_team = str(r.get("Home Team", ""))
            d_num = determine_division_for_team(h_team, div_map)
            div_vals.append(f"Division {d_num}")
        df["Division"] = div_vals
    elif "Division" not in df.columns:
        df["Division"] = "Division 5"

    df = format_time_columns_in_df(df)
    df = ensure_lock_columns(df)
    return df


def sort_skp_schedule(df):
    """Sorteert het rooster strikt chronologisch op datum en tijd."""
    df_sorted = df.copy()

    def get_sort_datetime(row):
        d_obj = parse_date_obj(row.get("Date"))
        t_str = normalize_time_str(row.get("Time"))
        if not d_obj:
            return datetime(2099, 1, 1, 0, 0)
        try:
            h, m = [int(x) for x in t_str.split(":")]
            return datetime(d_obj.year, d_obj.month, d_obj.day, h, m)
        except:
            return datetime(d_obj.year, d_obj.month, d_obj.day, 0, 0)

    sort_keys = df_sorted.apply(get_sort_datetime, axis=1)
    df_sorted["_sort_key"] = sort_keys
    df_sorted = df_sorted.sort_values(by="_sort_key").drop(columns=["_sort_key"]).reset_index(drop=True)
    return df_sorted


def build_team_busy_slots(all_games_df):
    busy_slots = set()
    if all_games_df is not None and not all_games_df.empty:
        col_home = find_col(all_games_df, ["home team", "thuis team", "home", "thuis"]) or "Home Team"
        col_away = find_col(all_games_df, ["away team", "uit team", "away", "uit"]) or "Away Team"
        col_date = find_col(all_games_df, ["date", "datum"]) or "Date"
        col_time = find_col(all_games_df, ["time", "tijd", "starttijd", "start time"]) or "Time"

        for _, row in all_games_df.iterrows():
            d_val = normalize_date_str(row.get(col_date))
            t_val = normalize_time_str(row.get(col_time))
            h_team = str(row.get(col_home, "")).lower()
            a_team = str(row.get(col_away, "")).lower()

            if "tantalus" in h_team:
                c_h = clean_team_code(h_team)
                if c_h:
                    busy_slots.add((c_h, d_val, t_val))
            if "tantalus" in a_team:
                c_a = clean_team_code(a_team)
                if c_a:
                    busy_slots.add((c_a, d_val, t_val))
    return busy_slots


def is_player_playing(player_team, date_val, time_val, curr_home, curr_away, busy_slots):
    if not player_team or str(player_team).strip() in ["", "nan", "None"]:
        return False
    p_code = clean_team_code(player_team)
    if not p_code:
        return False

    h_str = str(curr_home).lower()
    a_str = str(curr_away).lower()
    c_home = clean_team_code(h_str)
    c_away = clean_team_code(a_str)

    if "tantalus" in h_str and (p_code == c_home or p_code in c_home or c_home in p_code):
        return True
    if "tantalus" in a_str and (p_code == c_away or p_code in c_away or c_away in p_code):
        return True

    d_norm = normalize_date_str(date_val)
    t_norm = normalize_time_str(time_val)

    if (p_code, d_norm, t_norm) in busy_slots:
        return True
    for (b_team, b_d, b_t) in busy_slots:
        if b_d == d_norm and b_t == t_norm:
            if p_code in b_team or b_team in p_code:
                return True
    return False


def update_player_stats(sheets_dict):
    p_k = find_sheet(sheets_dict, ["Players skp", "Players", "Spelers"])
    s_k = find_sheet(sheets_dict, ["SKP", "Rooster"])
    c_k = find_sheet(sheets_dict, ["Committees", "Commissies"])

    if not p_k or not s_k or p_k not in sheets_dict or s_k not in sheets_dict:
        return sheets_dict

    players_df = standardize_players_df(sheets_dict[p_k].copy())
    skp_df = sheets_dict[s_k]
    committees_df = sheets_dict.get(c_k, pd.DataFrame())

    comm_points_map = {}
    if not committees_df.empty:
        col_c = committees_df.columns[0]
        col_p = committees_df.columns[-1] if len(committees_df.columns) > 1 else None
        for _, row in committees_df.dropna(subset=[col_c]).iterrows():
            c_name = str(row[col_c]).strip()
            try:
                pts = float(row[col_p])
            except:
                pts = 0.0
            comm_points_map[c_name] = pts

    ref_counts = {}
    table_counts = {}

    for _, row in skp_df.iterrows():
        for ref_col in ["Referee 1", "Referee 2"]:
            val = str(row.get(ref_col, "")).strip()
            if val and val not in ["nan", "None", "x", ""]:
                ref_counts[val] = ref_counts.get(val, 0) + 1

        for col in ["Scorer", "Timer", "24 sec operator"]:
            val = str(row.get(col, "")).strip()
            if val and val not in ["nan", "None", "x", ""]:
                table_counts[val] = table_counts.get(val, 0) + 1

    for idx, row in players_df.iterrows():
        f_name = str(row.get("First name", "")).strip()
        l_name = str(row.get("Last name", "")).strip()
        if not l_name or l_name in ["nan", "", "none", "None"]:
            continue

        full_name = f"{f_name} {l_name}".strip()
        r_count = ref_counts.get(full_name, 0)
        t_count = table_counts.get(full_name, 0)

        comm_name = str(row.get("Committee", "")).strip()
        if comm_name and comm_points_map:
            extra_pts = calculate_comm_points(comm_name, comm_points_map)
        else:
            try:
                extra_pts = float(row.get("Extra points", 0.0))
            except:
                extra_pts = 0.0

        total_pts = extra_pts + (r_count * 2) + (t_count * 1)
        players_df.at[idx, "Extra points"] = extra_pts
        players_df.at[idx, "Referee"] = r_count
        players_df.at[idx, "Table duty"] = t_count
        players_df.at[idx, "Total points"] = total_pts

    sheets_dict[p_k] = players_df
    return sheets_dict


def get_team_player_groups(players_df):
    team_groups = {}
    current_team = "Overig / Geen Team"
    for idx, row in players_df.iterrows():
        f_val = str(row.get("First name", "")).strip()
        l_val = str(row.get("Last name", "")).strip()
        t_explicit = str(row.get("Team", "")).strip()

        if l_val in ["nan", "", "none", "None"]:
            if f_val and f_val not in ["nan", "", "none", "None"]:
                current_team = f_val
                if current_team not in team_groups:
                    team_groups[current_team] = []
        else:
            actual_team = t_explicit if (t_explicit and t_explicit not in ["nan", "", "none", "None"]) else current_team
            if actual_team not in team_groups:
                team_groups[actual_team] = []
            team_groups[actual_team].append(idx)
    return team_groups


def get_all_available_teams(sheets_dict):
    p_key = find_sheet(sheets_dict, ["Players skp", "Players", "Spelers"])
    if not p_key or p_key not in sheets_dict:
        return []

    players_df = standardize_players_df(sheets_dict[p_key])
    groups = get_team_player_groups(players_df)

    valid_teams = [
        t for t in groups.keys()
        if str(t).strip().lower() not in ["overig / geen team", "geen team / overig", "overig", "geen team", "nan", "none", ""]
    ]
    return sorted(valid_teams, key=lambda x: x.lower())


def get_all_tantalus_teams(sheets_dict):
    div_key = find_sheet(sheets_dict, ["Divisions", "Divisies"])
    tantalus_teams = []
    if div_key and div_key in sheets_dict:
        div_df = sheets_dict[div_key]
        team_col = find_col(div_df, ["team", "tantalus team", "teams", "teamnaam"], fallback_index=0)
        if team_col:
            for t in div_df[team_col].dropna().unique():
                t_str = str(t).strip()
                if t_str and t_str not in ["nan", "None"]:
                    name_full = t_str if "tantalus" in t_str.lower() else f"Tantalus {t_str}"
                    if name_full not in tantalus_teams:
                        tantalus_teams.append(name_full)

    if not tantalus_teams:
        all_teams = get_all_available_teams(sheets_dict)
        for t in all_teams:
            name_full = t if "tantalus" in t.lower() else f"Tantalus {t}"
            if name_full not in tantalus_teams:
                tantalus_teams.append(name_full)

    return sorted(tantalus_teams, key=lambda x: x.lower())


def get_member_actual_team(players_df, target_idx):
    current_team = "Overig / Geen Team"
    for i, r in players_df.iterrows():
        f = str(r.get("First name", "")).strip()
        l = str(r.get("Last name", "")).strip()
        t = str(r.get("Team", "")).strip()

        if l in ["nan", "", "none", "None"]:
            if f and f.lower() not in ["nan", "none"]:
                current_team = f
        else:
            actual = t if (t and t.lower() not in ["nan", "none", ""]) else current_team
            if i == target_idx:
                return actual if actual else "Overig / Geen Team"
    return "Overig / Geen Team"


def insert_player_into_team(players_df, target_team_name, new_row_dict):
    df = players_df.copy()
    c_target = clean_team_code(target_team_name)

    insert_idx = None
    in_target_team = False

    for idx, row in df.iterrows():
        f_val = str(row.get("First name", "")).strip()
        l_val = str(row.get("Last name", "")).strip()

        if l_val in ["nan", "", "none", "None"] and f_val:
            if in_target_team:
                insert_idx = idx
                break
            if clean_team_code(f_val) == c_target:
                in_target_team = True
        else:
            if in_target_team:
                insert_idx = idx + 1

    new_row_df = pd.DataFrame([new_row_dict])

    if insert_idx is not None:
        df_top = df.iloc[:insert_idx]
        df_bottom = df.iloc[insert_idx:]
        res_df = pd.concat([df_top, new_row_df, df_bottom], ignore_index=True)
    else:
        res_df = pd.concat([df, new_row_df], ignore_index=True)

    return res_df


def handle_lock_all_synchronization(old_df, new_df):
    synced_df = new_df.copy()
    if LOCK_ALL_COL not in synced_df.columns:
        return synced_df

    for idx in synced_df.index:
        old_all = bool(old_df.at[idx, LOCK_ALL_COL]) if (idx in old_df.index and LOCK_ALL_COL in old_df.columns) else False
        new_all = bool(synced_df.at[idx, LOCK_ALL_COL])

        if new_all != old_all:
            for l_col in LOCK_COLS:
                synced_df.at[idx, l_col] = new_all
        else:
            all_on = all(bool(synced_df.at[idx, l_c]) for l_c in LOCK_COLS if l_c in synced_df.columns)
            synced_df.at[idx, LOCK_ALL_COL] = all_on

    return synced_df


def update_team_division_across_sheets(sheets_dict, team_name, new_div_number):
    """
    Past de divisie van een team aan op:
    1. Divisions sheet
    2. SKP sheet
    3. All games sheet
    Voorkomt KeyError en TypeError/LossySetitemError door veilige kolomdetectie en .loc met object types.
    """
    c_target = clean_team_code(team_name)
    new_div_str = f"Division {new_div_number}"

    # 1. Update Divisions sheet
    div_key = find_sheet(sheets_dict, ["Divisions", "Divisies"])
    if div_key and div_key in sheets_dict:
        div_df = sheets_dict[div_key].copy()
        team_col = find_col(div_df, ["team", "tantalus team", "teams", "teamnaam"], fallback_index=0)
        div_col = find_col(div_df, ["division", "divisie", "klasse", "poule"], fallback_index=1 if len(div_df.columns) > 1 else 0)

        if div_col:
            sample_val = str(div_df[div_col].dropna().iloc[0]).strip().lower() if not div_df[div_col].dropna().empty else ""
            use_raw_number = sample_val.isdigit()

            div_df[div_col] = div_df[div_col].astype(object)
            target_div_val = int(new_div_number) if use_raw_number else new_div_str

            matched = False
            for idx_d in div_df.index:
                team_val = div_df.at[idx_d, team_col] if team_col in div_df.columns else ""
                if clean_team_code(team_val) == c_target:
                    div_df.loc[idx_d, div_col] = target_div_val
                    matched = True

            if not matched and team_col:
                new_row_div = {c: "" for c in div_df.columns}
                new_row_div[team_col] = team_name
                new_row_div[div_col] = target_div_val
                div_df = pd.concat([div_df, pd.DataFrame([new_row_div])], ignore_index=True)

            sheets_dict[div_key] = div_df

    # 2. Update SKP sheet
    skp_key = find_sheet(sheets_dict, ["SKP", "Rooster"])
    affected_match_indices = []
    if skp_key and skp_key in sheets_dict:
        skp_df = sheets_dict[skp_key].copy()
        col_home_skp = find_col(skp_df, ["home team", "thuis team", "home", "thuis"]) or "Home Team"
        col_div_skp = find_col(skp_df, ["division", "divisie", "klasse", "poule"]) or "Division"

        if col_div_skp not in skp_df.columns:
            skp_df[col_div_skp] = ""
        skp_df[col_div_skp] = skp_df[col_div_skp].astype(object)

        for idx_s in skp_df.index:
            h_team = str(skp_df.at[idx_s, col_home_skp]).strip() if col_home_skp in skp_df.columns else ""
            if clean_team_code(h_team) == c_target:
                skp_df.loc[idx_s, col_div_skp] = new_div_str
                affected_match_indices.append(idx_s)
        sheets_dict[skp_key] = skp_df

    # 3. Update All games sheet (veilig met find_col en get om KeyError te voorkomen)
    all_games_key = find_sheet(sheets_dict, ["All games", "all games", "ALL GAMES"])
    if all_games_key and all_games_key in sheets_dict:
        all_games_df = sheets_dict[all_games_key].copy()
        col_home_all = find_col(all_games_df, ["home team", "thuis team", "home", "thuis"])
        col_away_all = find_col(all_games_df, ["away team", "uit team", "away", "uit"])
        div_col_all = find_col(all_games_df, ["division", "divisie", "poule", "klasse"]) or "Division"

        if div_col_all not in all_games_df.columns:
            all_games_df[div_col_all] = ""
        all_games_df[div_col_all] = all_games_df[div_col_all].astype(object)

        for idx_a in all_games_df.index:
            h_team = str(all_games_df.at[idx_a, col_home_all]).strip() if col_home_all else ""
            a_team = str(all_games_df.at[idx_a, col_away_all]).strip() if col_away_all else ""
            if clean_team_code(h_team) == c_target or clean_team_code(a_team) == c_target:
                all_games_df.loc[idx_a, div_col_all] = new_div_str
        sheets_dict[all_games_key] = all_games_df

    return sheets_dict, affected_match_indices


def reassign_invalid_referees_after_division_change(sheets_dict, affected_indices):
    """
    Hercontroleert en herplaatst scheidsrechters en tafelaars op wedstrijden
    waarvan de divisie is gewijzigd, en logt alle vervangingen.
    """
    skp_key = find_sheet(sheets_dict, ["SKP", "Rooster"])
    players_key = find_sheet(sheets_dict, ["Players skp", "Players", "Spelers"])
    div_key = find_sheet(sheets_dict, ["Divisions", "Divisies"])

    if not skp_key or not affected_indices:
        return sheets_dict, []

    skp_df = sheets_dict[skp_key].copy()
    players_df = standardize_players_df(sheets_dict.get(players_key, pd.DataFrame()))
    tantalus_div_map = build_division_map(sheets_dict.get(div_key, pd.DataFrame()))

    player_dips = {}
    for _, p_r in players_df.iterrows():
        f = str(p_r.get("First name", "")).strip()
        l = str(p_r.get("Last name", "")).strip()
        if not l or l.lower() in ["nan", "none"]:
            continue
        full_n = f"{f} {l}".strip()
        d_val = str(p_r.get("Diploma", "")).strip().upper().replace(" ", "").replace("-", "")
        if "L4" in d_val: norm = "L4"
        elif "L3" in d_val: norm = "L3"
        elif "BS3" in d_val: norm = "BS3"
        elif "BS2" in d_val: norm = "BS2"
        elif "BS1" in d_val: norm = "BS1"
        else: norm = "NONE"
        player_dips[full_n] = norm

    change_logs = []
    indices_to_rerun = set()

    col_home = find_col(skp_df, ["home team", "thuis team", "home", "thuis"]) or "Home Team"
    col_away = find_col(skp_df, ["away team", "uit team", "away", "uit"]) or "Away Team"
    col_date = find_col(skp_df, ["date", "datum"]) or "Date"
    col_time = find_col(skp_df, ["time", "tijd"]) or "Time"

    for idx in affected_indices:
        home_team = str(skp_df.at[idx, col_home]).strip() if col_home in skp_df.columns else ""
        away_team = str(skp_df.at[idx, col_away]).strip() if col_away in skp_df.columns else ""
        m_d = str(skp_df.at[idx, col_date]).strip() if col_date in skp_df.columns else ""
        m_t = str(skp_df.at[idx, col_time]).strip() if col_time in skp_df.columns else ""
        div_num = determine_division_for_team(home_team, tantalus_div_map)
        match_desc = f"{m_d} {m_t} ({home_team} vs {away_team})"

        for ref_col in ["Referee 1", "Referee 2"]:
            l_col = LOCK_MAP[ref_col]
            is_locked = bool(skp_df.at[idx, l_col]) if l_col in skp_df.columns else False
            curr_ref = str(skp_df.at[idx, ref_col]).strip() if ref_col in skp_df.columns else ""

            if is_locked or not curr_ref or curr_ref.lower() in ["nan", "none", "x"]:
                continue

            ref_dip = player_dips.get(curr_ref, "NONE")
            is_invalid = False
            reason_str = ""

            if div_num <= 1 and ref_dip not in ["L4", "L3", "BS3", "BS2", "BS1"]:
                is_invalid = True
                reason_str = f"heeft niveau '{ref_dip}' maar Divisie {div_num} vereist minimaal BS3/BS2"
            elif div_num == 2 and ref_dip == "NONE":
                is_invalid = True
                reason_str = f"heeft geen scheidsrechtersdiploma voor Divisie 2"
            elif div_num == 3 and ref_dip == "NONE":
                is_invalid = True
                reason_str = f"heeft geen scheidsrechtersdiploma voor Divisie 3"

            if is_invalid:
                skp_df.at[idx, ref_col] = ""
                indices_to_rerun.add(idx)
                change_logs.append({
                    "match": match_desc,
                    "task": ref_col,
                    "old_person": curr_ref,
                    "new_person": "Wordt heringedeeld",
                    "reason": reason_str
                })

        # Bij degradatie naar Div 4+ vervalt de 24 sec operator
        if div_num > 3 and "24 sec operator" in skp_df.columns:
            curr_24s = str(skp_df.at[idx, "24 sec operator"]).strip()
            l_24s = bool(skp_df.at[idx, "Lock 24s"]) if "Lock 24s" in skp_df.columns else False
            if curr_24s and curr_24s.lower() not in ["nan", "none", ""] and not l_24s:
                skp_df.at[idx, "24 sec operator"] = ""
                change_logs.append({
                    "match": match_desc,
                    "task": "24 sec operator",
                    "old_person": curr_24s,
                    "new_person": "Geen (taak vervalt)",
                    "reason": f"Divisie {div_num} heeft geen 24-seconden operator nodig"
                })

        # Bij promotie naar Div 1-3 is 24 sec operator vereist
        if div_num <= 3 and "24 sec operator" in skp_df.columns:
            curr_24s = str(skp_df.at[idx, "24 sec operator"]).strip()
            if not curr_24s or curr_24s.lower() in ["nan", "none"]:
                indices_to_rerun.add(idx)

    sheets_dict[skp_key] = skp_df

    if indices_to_rerun:
        current_max_tasks = st.session_state.get("slider_max_daily_tasks", 1)
        sheets_dict, _ = run_assignment_core(
            sheets_dict,
            list(indices_to_rerun),
            max_daily_tasks=current_max_tasks,
            preserve_manual=True
        )

        skp_df_updated = sheets_dict[skp_key]
        for log in change_logs:
            if log["new_person"] == "Wordt heringedeeld":
                for idx_m in indices_to_rerun:
                    h_m = str(skp_df_updated.at[idx_m, col_home]).strip() if col_home in skp_df_updated.columns else ""
                    a_m = str(skp_df_updated.at[idx_m, col_away]).strip() if col_away in skp_df_updated.columns else ""
                    d_m = str(skp_df_updated.at[idx_m, col_date]).strip() if col_date in skp_df_updated.columns else ""
                    t_m = str(skp_df_updated.at[idx_m, col_time]).strip() if col_time in skp_df_updated.columns else ""
                    m_str = f"{d_m} {t_m} ({h_m} vs {a_m})"
                    if m_str == log["match"]:
                        assigned_now = str(skp_df_updated.at[idx_m, log["task"]]).strip() if log["task"] in skp_df_updated.columns else ""
                        log["new_person"] = assigned_now if assigned_now else "Openstaand (geen geldige arbiter beschikbaar)"

    return sheets_dict, change_logs


def run_assignment_core(sheets_dict, target_match_indices, max_daily_tasks=1, preserve_manual=False):
    skp_key = find_sheet(sheets_dict, ["SKP", "Rooster"]) or "SKP"
    players_key = find_sheet(sheets_dict, ["Players skp", "Players", "Spelers"]) or "Players"
    all_games_key = find_sheet(sheets_dict, ["All games", "all games", "ALL GAMES"]) or "all games"
    div_key = find_sheet(sheets_dict, ["Divisions", "Divisies"]) or "Divisions"

    divisions_df = sheets_dict.get(div_key, pd.DataFrame())
    tantalus_div_map = build_division_map(divisions_df)

    skp_df = standardize_skp_df(sheets_dict[skp_key], div_map=tantalus_div_map)
    all_games_df = sheets_dict.get(all_games_key, pd.DataFrame())
    if not all_games_df.empty:
        all_games_df = standardize_skp_df(all_games_df)

    players_df = standardize_players_df(sheets_dict.get(players_key, pd.DataFrame()))

    player_team_map = {}
    current_team = ""
    for _, p_row in players_df.iterrows():
        f_val = str(p_row.get("First name", "")).strip()
        l_val = str(p_row.get("Last name", "")).strip()
        t_explicit = str(p_row.get("Team", "")).strip()

        if l_val in ["nan", "", "none", "None"]:
            if f_val and f_val not in ["nan", "", "none", "None"]:
                current_team = f_val
        else:
            full_n = f"{f_val} {l_val}".strip()
            player_team_map[full_n] = t_explicit if (t_explicit and t_explicit not in ["nan", "", "none", "None"]) else current_team

    valid_players_dict = {}
    excluded_board_coach = set()
    excluded_recreational = set()
    for idx, p_row in players_df.iterrows():
        l_val = str(p_row.get("Last name", "")).strip()
        f_val = str(p_row.get("First name", "")).strip()
        if not l_val or l_val in ["nan", "", "none", "None"]:
            continue

        comm_val = str(p_row.get("Committee", "")).strip().lower()
        extra_val = str(p_row.get("Extra", "")).strip().lower()
        roles_combined = f"{comm_val} {extra_val}"

        full_name = f"{f_val} {l_val}".strip()
        is_board = "board" in roles_combined or "bestuur" in roles_combined
        is_coach = "coach" in roles_combined and "assistant coach" not in roles_combined
        if is_board or is_coach:
            excluded_board_coach.add(full_name)
            continue

        is_recreational = "recreation" in extra_val or "recreant" in extra_val
        if is_recreational:
            excluded_recreational.add(full_name)
            continue

        d_val = (
            str(p_row.get("Diploma", ""))
            .strip()
            .upper()
            .replace(" ", "")
            .replace("-", "")
        )
        if "L4" in d_val:
            norm_dip = "L4"
        elif "L3" in d_val:
            norm_dip = "L3"
        elif "BS3" in d_val:
            norm_dip = "BS3"
        elif "BS2" in d_val:
            norm_dip = "BS2"
        elif "BS1" in d_val:
            norm_dip = "BS1"
        else:
            norm_dip = "NONE"

        team = player_team_map.get(full_name, "")
        init_extra_pts = float(p_row.get("Extra points", 0.0))
        season_type = str(p_row.get("Full/ half season", "Full season")).strip()

        valid_players_dict[full_name] = {
            "First name": f_val,
            "Last name": l_val,
            "Full Name": full_name,
            "Team": team,
            "Diploma": norm_dip,
            "Has_Diploma": (norm_dip != "NONE"),
            "Is_BS3_Plus": norm_dip in ["BS3", "L3", "L4"],
            "Is_Aurelie": "aurelie" in full_name.lower(),
            "Base_Points": init_extra_pts,
            "Season_Type": season_type
        }

    auto_assigned = st.session_state.get("auto_assigned_cells", set())

    for idx in target_match_indices:
        for col_c in TASK_COLS:
            l_col = LOCK_MAP[col_c]
            is_locked = bool(skp_df.at[idx, l_col]) if l_col in skp_df.columns else False
            curr_v = str(skp_df.at[idx, col_c]).strip()
            is_manual_entry = (preserve_manual and curr_v not in ["", "nan", "None", "x"] and (idx, col_c) not in auto_assigned)

            if is_manual_entry:
                skp_df.at[idx, l_col] = True
            elif col_c in skp_df.columns and not is_locked:
                if str(skp_df.at[idx, col_c]).strip().lower() != "x":
                    skp_df.at[idx, col_c] = ""
                    if (idx, col_c) in auto_assigned:
                        auto_assigned.remove((idx, col_c))

    busy_game_slots = build_team_busy_slots(all_games_df)
    ref_tasks_counter = {p: 0 for p in valid_players_dict}
    table_tasks_counter = {p: 0 for p in valid_players_dict}
    duty_specific_counter = {
        p: {"Scorer": 0, "Timer": 0, "24 sec operator": 0}
        for p in valid_players_dict
    }
    player_busy_times = {p: set() for p in valid_players_dict}
    player_day_counts = {p: {} for p in valid_players_dict}

    for idx_row in skp_df.index:
        d_val = normalize_date_str(skp_df.at[idx_row, "Date"])
        t_val = normalize_time_str(skp_df.at[idx_row, "Time"])
        for col in ["Referee 1", "Referee 2"]:
            name = str(skp_df.at[idx_row, col]).strip()
            if name in valid_players_dict:
                ref_tasks_counter[name] += 1
                player_busy_times[name].add((d_val, t_val))
                player_day_counts[name][d_val] = player_day_counts[name].get(d_val, 0) + 1
        for col in ["Scorer", "Timer", "24 sec operator"]:
            name = str(skp_df.at[idx_row, col]).strip()
            if name in valid_players_dict:
                table_tasks_counter[name] += 1
                duty_specific_counter[name][col] = duty_specific_counter[name].get(col, 0) + 1
                player_busy_times[name].add((d_val, t_val))
                player_day_counts[name][d_val] = player_day_counts[name].get(d_val, 0) + 1

    def match_sort_key(idx_val):
        d_str = normalize_date_str(skp_df.at[idx_val, "Date"])
        t_str = normalize_time_str(skp_df.at[idx_val, "Time"])
        h_t = str(skp_df.at[idx_val, "Home Team"]).strip()
        div_n = determine_division_for_team(h_t, tantalus_div_map)
        return (d_str, div_n, t_str)

    sorted_match_indices = sorted(target_match_indices, key=match_sort_key)
    assignment_warnings = []

    for idx in sorted_match_indices:
        home_team = str(skp_df.at[idx, "Home Team"]).strip()
        away_team = str(skp_df.at[idx, "Away Team"]).strip()

        if "tantalus" not in home_team.lower():
            continue

        m_date = str(skp_df.at[idx, "Date"])
        d_norm = normalize_date_str(m_date)
        m_time = skp_df.at[idx, "Time"]
        t_norm = normalize_time_str(m_time)
        div_num = determine_division_for_team(home_team, tantalus_div_map)
        skp_df.at[idx, "Division"] = f"Division {div_num}"

        assigned_in_match = {
            str(skp_df.at[idx, c]).strip()
            for c in TASK_COLS
            if str(skp_df.at[idx, c]).strip() not in ["", "nan", "None", "x"]
        }

        h_code = clean_team_code(home_team)
        is_tantalus_mse1 = bool(re.search(r"mse[\s\-]*1\b", h_code))

        def is_physically_free(p_name):
            if (d_norm, t_norm) in player_busy_times[p_name]:
                return False
            if p_name in assigned_in_match:
                return False
            if is_player_playing(valid_players_dict[p_name]["Team"], m_date, m_time, home_team, away_team, busy_game_slots):
                return False
            if not is_player_eligible_for_date(valid_players_dict[p_name]["Season_Type"], m_date):
                return False
            return True

        if is_tantalus_mse1:
            for ref_col in ["Referee 1", "Referee 2"]:
                l_col = LOCK_MAP[ref_col]
                is_locked = bool(skp_df.at[idx, l_col]) if l_col in skp_df.columns else False
                if not is_locked:
                    skp_df.at[idx, ref_col] = "x"
        else:
            for ref_col in ["Referee 1", "Referee 2"]:
                l_col = LOCK_MAP[ref_col]
                is_locked = bool(skp_df.at[idx, l_col]) if l_col in skp_df.columns else False
                if is_locked:
                    continue

                curr_val = str(skp_df.at[idx, ref_col]).strip()
                if curr_val.lower() == "x" or (curr_val and curr_val not in ["", "None", "nan"]):
                    continue

                other_ref_col = "Referee 2" if ref_col == "Referee 1" else "Referee 1"
                other_ref = str(skp_df.at[idx, other_ref_col]).strip()
                other_ref_dip = valid_players_dict.get(other_ref, {}).get("Diploma", "")
                has_aurelie_assigned = ("aurelie" in other_ref.lower())

                primary_dips = set()
                fallback_dips = set()
                emergency_dips = set()

                if div_num <= 1:
                    primary_dips = {"L4", "L3", "BS3"}
                    fallback_dips = {"BS2"}
                    emergency_dips = {"BS1"}
                elif div_num == 2:
                    if has_aurelie_assigned:
                        primary_dips = {"BS3", "L3", "L4"}
                        fallback_dips = {"BS2"}
                        emergency_dips = {"BS1"}
                    elif other_ref_dip in ["BS3", "L3", "L4"]:
                        primary_dips = {"L3", "L4", "BS3", "BS2"}
                        fallback_dips = {"BS2"}
                        emergency_dips = {"BS1"}
                    elif other_ref_dip == "BS2":
                        primary_dips = {"BS3", "L3", "L4", "BS2"}
                        fallback_dips = {"BS1"}
                        emergency_dips = {"BS1"}
                    else:
                        primary_dips = {"BS3", "L3", "L4"}
                        fallback_dips = {"BS2"}
                        emergency_dips = {"BS1"}
                elif div_num == 3:
                    if other_ref_dip in ["BS3", "L3", "L4"]:
                        primary_dips = {"BS2"}
                        fallback_dips = {"BS2"}
                        emergency_dips = {"BS1"}
                    elif other_ref_dip == "BS2":
                        primary_dips = {"BS3", "BS2"}
                        fallback_dips = {"BS2"}
                        emergency_dips = {"BS1"}
                    else:
                        primary_dips = {"BS3", "BS2"}
                        fallback_dips = {"BS2"}
                        emergency_dips = {"BS1"}
                elif div_num == 4:
                    primary_dips = {"BS2"}
                    fallback_dips = {"BS1"}
                    emergency_dips = {"BS1"}
                else:
                    primary_dips = {"BS1", "BS2"}
                    fallback_dips = {"BS1", "BS2"}

                all_ref_pts = [
                    p["Base_Points"] + (ref_tasks_counter[n] * 2) + (table_tasks_counter[n] * 1)
                    for n, p in valid_players_dict.items() if p["Has_Diploma"]
                ]
                avg_ref_pts = sum(all_ref_pts) / len(all_ref_pts) if all_ref_pts else 0.0

                ref_cands = []
                for p_name, p_info in valid_players_dict.items():
                    if not p_info["Has_Diploma"]:
                        continue

                    if not is_physically_free(p_name):
                        continue

                    day_cnt = player_day_counts[p_name].get(d_norm, 0)
                    if day_cnt >= max_daily_tasks:
                        continue

                    curr_pts = p_info["Base_Points"] + (ref_tasks_counter[p_name] * 2) + (table_tasks_counter[p_name] * 1)
                    if (curr_pts + 2.0) > 16.0:
                        continue

                    dip = p_info["Diploma"]
                    is_aurelie = p_info["Is_Aurelie"]

                    under_12 = 0 if curr_pts < 12.0 else 1
                    is_overloaded = (curr_pts > (avg_ref_pts + 6.0))

                    if dip in ["L4", "L3"]:
                        dip_rank = 1
                    elif dip == "BS3":
                        dip_rank = 2
                    elif dip == "BS2":
                        dip_rank = 3
                    elif dip == "BS1":
                        dip_rank = 4
                    else:
                        dip_rank = 5

                    high_div_boost = 0
                    if div_num <= 3 and p_info["Is_BS3_Plus"]:
                        high_div_boost = -1

                    aurelie_priority = 10
                    if is_aurelie:
                        if div_num == 2 and not has_aurelie_assigned:
                            aurelie_priority = 2 if is_overloaded else 0
                        elif div_num in [2, 3]:
                            aurelie_priority = 3 if is_overloaded else 1
                        else:
                            continue

                    tier = 99
                    if dip in primary_dips:
                        tier = 2 if is_overloaded else 1
                    elif dip in fallback_dips:
                        tier = 1 if is_overloaded else 2
                    elif dip in emergency_dips:
                        tier = 3
                    elif p_info["Is_BS3_Plus"] and div_num > 3:
                        tier = 4
                    else:
                        tier = 5

                    ref_cands.append({
                        "name": p_name,
                        "day_count": day_cnt,
                        "aurelie_prio": aurelie_priority,
                        "high_div_boost": high_div_boost,
                        "tier": tier,
                        "dip_rank": dip_rank,
                        "total_points": curr_pts,
                        "under_12": under_12,
                        "ref_tasks": ref_tasks_counter[p_name],
                    })

                ref_cands.sort(
                    key=lambda x: (
                        x["day_count"],
                        x["aurelie_prio"],
                        x["high_div_boost"],
                        x["tier"],
                        x["dip_rank"],
                        x["total_points"],
                        x["under_12"],
                        x["ref_tasks"],
                        random.random(),
                    )
                )

                if ref_cands:
                    chosen = ref_cands[0]["name"]
                    skp_df.at[idx, ref_col] = chosen
                    assigned_in_match.add(chosen)
                    player_busy_times[chosen].add((d_norm, t_norm))
                    player_day_counts[chosen][d_norm] = player_day_counts[chosen].get(d_norm, 0) + 1
                    ref_tasks_counter[chosen] += 1
                    auto_assigned.add((idx, ref_col))
                else:
                    reasons = []
                    for p_name, p_info in valid_players_dict.items():
                        if not p_info["Has_Diploma"]:
                            continue
                        p_team = p_info["Team"]
                        p_day_cnt = player_day_counts[p_name].get(d_norm, 0)
                        p_pts = p_info["Base_Points"] + (ref_tasks_counter[p_name] * 2) + (table_tasks_counter[p_name] * 1)

                        if (p_pts + 2.0) > 16.0:
                            reasons.append(f"**{p_name}** ({p_info['Diploma']}): Bereikt maximum van 16 punten ({p_pts} pnt).")
                        elif (d_norm, t_norm) in player_busy_times[p_name]:
                            reasons.append(f"**{p_name}** ({p_info['Diploma']}): Heeft al een taak om {t_norm}.")
                        elif is_player_playing(p_team, m_date, m_time, home_team, away_team, busy_game_slots):
                            reasons.append(f"**{p_name}** ({p_info['Diploma']}): Speelt zelf met team *{p_team}*.")
                        elif p_day_cnt >= max_daily_tasks:
                            reasons.append(f"**{p_name}** ({p_info['Diploma']}): Daglimiet van {max_daily_tasks} ta(a)k(en) bereikt.")
                        elif not is_player_eligible_for_date(p_info["Season_Type"], m_date):
                            reasons.append(f"**{p_name}**: Speelt halve seizoen ({p_info['Season_Type']}).")

                    for exc_name in excluded_board_coach:
                        reasons.append(f"**{exc_name}**: Vrijgesteld van taken (Board / Coach).")

                    for exc_rec in excluded_recreational:
                        reasons.append(f"**{exc_rec}**: Vrijgesteld van taken (Recreational).")

                    if not reasons:
                        reasons.append("Geen actieve gediplomeerde arbiters beschikbaar (of allen overschrijden 16 punten).")

                    assignment_warnings.append({
                        "match": f"{home_team} vs {away_team}",
                        "slot": f"{m_date} om {t_norm} ({ref_col})",
                        "reasons": reasons[:6]
                    })

        if div_num > 3:
            skp_df.at[idx, "24 sec operator"] = ""

        table_tasks_needed = ["Scorer", "Timer"]
        if div_num <= 3:
            table_tasks_needed.append("24 sec operator")

        for col in table_tasks_needed:
            l_col = LOCK_MAP[col]
            is_locked = bool(skp_df.at[idx, l_col]) if l_col in skp_df.columns else False
            if is_locked:
                continue

            curr_val = str(skp_df.at[idx, col]).strip()
            if curr_val and curr_val not in ["", "None", "nan"]:
                continue

            table_cands = []
            for p_name, p_info in valid_players_dict.items():
                if p_info["Has_Diploma"]:
                    continue
                if not is_physically_free(p_name):
                    continue

                day_cnt = player_day_counts[p_name].get(d_norm, 0)
                if day_cnt >= max_daily_tasks:
                    continue

                curr_pts = p_info["Base_Points"] + (ref_tasks_counter[p_name] * 2) + (table_tasks_counter[p_name] * 1)

                if (curr_pts + 1.0) > 16.0:
                    continue

                under_12 = 0 if curr_pts < 12.0 else 1
                duty_done = duty_specific_counter[p_name].get(col, 0)

                table_cands.append({
                    "name": p_name,
                    "day_count": day_cnt,
                    "total_points": curr_pts,
                    "under_12": under_12,
                    "duty_specific": duty_done,
                    "total_table": table_tasks_counter[p_name],
                })

            table_cands.sort(
                key=lambda x: (
                    x["day_count"],
                    x["total_points"],
                    x["under_12"],
                    x["duty_specific"],
                    x["total_table"],
                    random.random(),
                )
            )

            if table_cands:
                chosen = table_cands[0]["name"]
                skp_df.at[idx, col] = chosen
                assigned_in_match.add(chosen)
                player_busy_times[chosen].add((d_norm, t_norm))
                player_day_counts[chosen][d_norm] = player_day_counts[chosen].get(d_norm, 0) + 1
                table_tasks_counter[chosen] += 1
                duty_specific_counter[chosen][col] = duty_specific_counter[chosen].get(col, 0) + 1
                auto_assigned.add((idx, col))

    for i in skp_df.index:
        all_on = all(bool(skp_df.at[i, l_c]) for l_c in LOCK_COLS if l_c in skp_df.columns)
        skp_df.at[i, LOCK_ALL_COL] = all_on

    st.session_state["auto_assigned_cells"] = auto_assigned
    sheets_dict[skp_key] = make_arrow_compatible(skp_df)
    sheets_dict = update_player_stats(sheets_dict)
    return sheets_dict, assignment_warnings


def auto_reassign_future_schedule(sheets_dict, days_ahead=7):
    skp_key = find_sheet(sheets_dict, ["SKP", "Rooster"]) or "SKP"
    if skp_key not in sheets_dict:
        return sheets_dict, []

    skp_df = sheets_dict[skp_key]
    cutoff_date = datetime.now().date() + timedelta(days=days_ahead)

    target_indices = []
    for idx, r in skp_df.iterrows():
        d_obj = parse_date_obj(r.get("Date"))
        if d_obj and d_obj >= cutoff_date:
            target_indices.append(idx)

    if target_indices:
        current_max_tasks = st.session_state.get("slider_max_daily_tasks", 1)
        return run_assignment_core(sheets_dict, target_indices, max_daily_tasks=current_max_tasks, preserve_manual=True)
    return sheets_dict, []


def validate_schedule_rules(sheets_dict):
    """Controleert of het rooster aan alle regels voldoet."""
    skp_key = find_sheet(sheets_dict, ["SKP", "Rooster"])
    players_key = find_sheet(sheets_dict, ["Players skp", "Players", "Spelers"])
    all_games_key = find_sheet(sheets_dict, ["All games", "all games", "ALL GAMES"])
    div_key = find_sheet(sheets_dict, ["Divisions", "Divisies"])

    if not skp_key or skp_key not in sheets_dict:
        return ["Rooster sheet (SKP) niet gevonden."]
    if not players_key or players_key not in sheets_dict:
        return ["Spelers sheet (Players) niet gevonden."]

    skp_df = sheets_dict[skp_key]
    players_df = standardize_players_df(sheets_dict[players_key])
    divisions_df = sheets_dict.get(div_key, pd.DataFrame())
    tantalus_div_map = build_division_map(divisions_df)
    all_games_df = sheets_dict.get(all_games_key, pd.DataFrame())
    busy_game_slots = build_team_busy_slots(all_games_df)

    player_team_map = {}
    current_team = ""
    for _, p_row in players_df.iterrows():
        f_val = str(p_row.get("First name", "")).strip()
        l_val = str(p_row.get("Last name", "")).strip()
        t_explicit = str(p_row.get("Team", "")).strip()

        if l_val in ["nan", "", "none", "None"]:
            if f_val and f_val not in ["nan", "", "none", "None"]:
                current_team = f_val
        else:
            full_n = f"{f_val} {l_val}".strip()
            player_team_map[full_n] = t_explicit if (t_explicit and t_explicit not in ["nan", "", "none", "None"]) else current_team

    player_meta = {}
    for idx, p_row in players_df.iterrows():
        l_val = str(p_row.get("Last name", "")).strip()
        f_val = str(p_row.get("First name", "")).strip()
        if not l_val or l_val in ["nan", "", "none", "None"]:
            continue

        full_n = f"{f_val} {l_val}".strip()
        comm_val = str(p_row.get("Committee", "")).strip().lower()
        extra_val = str(p_row.get("Extra", "")).strip().lower()
        roles_c = f"{comm_val} {extra_val}"

        is_board = "board" in roles_c or "bestuur" in roles_c
        is_coach = "coach" in roles_c and "assistant coach" not in roles_c
        is_rec = "recreation" in extra_val or "recreant" in extra_val

        d_val = str(p_row.get("Diploma", "")).strip().upper().replace(" ", "").replace("-", "")
        if "L4" in d_val: norm_dip = "L4"
        elif "L3" in d_val: norm_dip = "L3"
        elif "BS3" in d_val: norm_dip = "BS3"
        elif "BS2" in d_val: norm_dip = "BS2"
        elif "BS1" in d_val: norm_dip = "BS1"
        else: norm_dip = "NONE"

        player_meta[full_n] = {
            "team": player_team_map.get(full_n, ""),
            "diploma": norm_dip,
            "season_type": str(p_row.get("Full/ half season", "Full season")).strip(),
            "is_board": is_board,
            "is_coach": is_coach,
            "is_rec": is_rec,
            "total_pts": float(p_row.get("Total points", 0.0)),
        }

    violations = []
    slot_assignments = {}

    col_home = find_col(skp_df, ["home team", "thuis team", "home", "thuis"]) or "Home Team"
    col_away = find_col(skp_df, ["away team", "uit team", "away", "uit"]) or "Away Team"
    col_date = find_col(skp_df, ["date", "datum"]) or "Date"
    col_time = find_col(skp_df, ["time", "tijd"]) or "Time"

    for idx, row in skp_df.iterrows():
        home = str(row.get(col_home, "")).strip()
        away = str(row.get(col_away, "")).strip()
        date_str = str(row.get(col_date, "")).strip()
        time_str = str(row.get(col_time, "")).strip()
        d_norm = normalize_date_str(date_str)
        t_norm = normalize_time_str(time_str)
        is_tantalus_home = "tantalus" in home.lower()
        h_code = clean_team_code(home)
        is_tantalus_mse1 = bool(re.search(r"mse[\s\-]*1\b", h_code))
        div_num = determine_division_for_team(home, tantalus_div_map)

        match_label = f"Rij {idx + 1} ({date_str} {t_norm}: {home} vs {away})"

        refs_in_match = []
        for r_c in ["Referee 1", "Referee 2"]:
            val = str(row.get(r_c, "")).strip()
            if is_tantalus_home and not is_tantalus_mse1:
                if not val or val.lower() in ["nan", "none", ""]:
                    violations.append(f"{match_label}: **{r_c}** is niet ingevuld.")
                elif val.lower() == "x":
                    violations.append(f"{match_label}: **{r_c}** staat op 'x' terwijl dit geen MSE 1 wedstrijd is.")
            if val and val.lower() not in ["nan", "none", "x", ""]:
                refs_in_match.append((r_c, val))

        table_needed = ["Scorer", "Timer"]
        if div_num <= 3:
            table_needed.append("24 sec operator")
        for t_col in table_needed:
            val = str(row.get(t_col, "")).strip()
            if is_tantalus_home:
                if not val or val.lower() in ["nan", "none", ""]:
                    violations.append(f"{match_label}: **{t_col}** is niet ingevuld.")

        ref_dips = []
        for r_col, r_name in refs_in_match:
            dip = player_meta.get(r_name, {}).get("diploma", "NONE")
            ref_dips.append((r_col, r_name, dip))

        for r_col, r_name, dip in ref_dips:
            if div_num <= 1 and dip not in ["L4", "L3", "BS3", "BS2", "BS1"]:
                violations.append(f"{match_label}: {r_name} fluit Divisie 1 met diploma '{dip}' (vereist minimaal BS3/BS2).")
            elif div_num == 2 and dip == "NONE":
                violations.append(f"{match_label}: {r_name} fluit Divisie 2 zonder scheidsrechtersdiploma.")
            elif div_num == 3 and dip == "NONE":
                violations.append(f"{match_label}: {r_name} fluit Divisie 3 zonder scheidsrechtersdiploma.")

        for col_name in TASK_COLS:
            person = str(row.get(col_name, "")).strip()
            if not person or person.lower() in ["nan", "none", "x", ""]:
                continue

            slot_key = (person, d_norm, t_norm)
            if slot_key in slot_assignments:
                prev_match, prev_col = slot_assignments[slot_key]
                violations.append(f"{person} is dubbel geboekt op {d_norm} om {t_norm}: zowel bij {prev_match} ({prev_col}) als bij {match_label} ({col_name}).")
            else:
                slot_assignments[slot_key] = (match_label, col_name)

            if person in player_meta:
                p_info = player_meta[person]

                if is_player_playing(p_info["team"], date_str, time_str, home, away, busy_game_slots):
                    violations.append(f"{match_label}: {person} staat ingedeeld als **{col_name}**, maar moet zelf spelen met team *{p_info['team']}*.")

                if not is_player_eligible_for_date(p_info["season_type"], date_str):
                    violations.append(f"{match_label}: {person} ({col_name}) is niet actief op {date_str} volgens seizoenshelft ({p_info['season_type']}).")

                if p_info["is_board"] or p_info["is_coach"]:
                    violations.append(f"{match_label}: {person} ({col_name}) is vrijgesteld van taken (Bestuur/Coach).")
                if p_info["is_rec"]:
                    violations.append(f"{match_label}: {person} ({col_name}) is vrijgesteld van taken (Recreant).")

    for p_name, p_data in player_meta.items():
        if p_data["total_pts"] > 16.0:
            violations.append(f"**{p_name}** overschrijdt het maximum van 16 punten ({p_data['total_pts']:.1f} punten).")

    return violations


# =========================================================================
# FILE INLEZEN / VERWERKEN MET DIRECTE RERUN
# =========================================================================
if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    curr_filename = getattr(uploaded_file, "name", "excel")

    need_reload = (
        "file_bytes" not in st.session_state
        or st.session_state.get("last_uploaded_filename") != curr_filename
        or st.sidebar.button("🔄 Bestand opnieuw inlezen")
    )

    if need_reload:
        st.session_state["file_bytes"] = file_bytes
        st.session_state["last_uploaded_filename"] = curr_filename

        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        raw_sheets = {}
        for sheet in xls.sheet_names:
            df = xls.parse(sheet)
            raw_sheets[sheet] = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]

        div_sheet_name = find_sheet(raw_sheets, ["Divisions", "Divisies"])
        div_map = (
            build_division_map(raw_sheets[div_sheet_name])
            if div_sheet_name
            else {}
        )

        orig_sheets = {}
        for sheet, df in raw_sheets.items():
            if "player" in sheet.lower():
                df = standardize_players_df(df)
            elif "skp" in sheet.lower() and "player" not in sheet.lower():
                df = standardize_skp_df(df, div_map=div_map)
                for col in TASK_COLS:
                    if col in df.columns:
                        df[col] = df[col].fillna("").astype(str).replace({"nan": "", "None": ""})
                df = sort_skp_schedule(df)
            elif "game" in sheet.lower():
                df = format_time_columns_in_df(df)
            orig_sheets[sheet] = df

        st.session_state["original_sheets"] = copy.deepcopy(orig_sheets)
        st.session_state["sheets"] = copy.deepcopy(orig_sheets)
        st.session_state["indeling_gedaan"] = False
        st.session_state["assignment_warnings"] = []
        st.session_state["division_change_logs"] = []
        st.session_state["validation_results"] = None
        st.session_state["auto_assigned_cells"] = set()
        st.rerun()


# =========================================================================
# HOOFDPROGRAMMA & ZIJBALK
# =========================================================================
if "sheets" in st.session_state:
    sheets = st.session_state["sheets"]

    players_key = find_sheet(sheets, ["Players skp", "Players", "Spelers"]) or "Players"
    skp_key = find_sheet(sheets, ["SKP", "Rooster"]) or "SKP"
    comm_key = find_sheet(sheets, ["Committees", "Commissies"]) or "Committees"
    all_games_key = find_sheet(sheets, ["All games", "all games", "ALL GAMES"]) or "all games"
    div_key = find_sheet(sheets, ["Divisions", "Divisies"]) or "Divisions"

    divisions_df = sheets.get(div_key, pd.DataFrame())
    tantalus_div_map = build_division_map(divisions_df)

    if skp_key in sheets:
        sheets[skp_key] = standardize_skp_df(sheets[skp_key], div_map=tantalus_div_map)
    if all_games_key in sheets:
        sheets[all_games_key] = format_time_columns_in_df(sheets[all_games_key])

    tab_names = list(sheets.keys())
    col_sel_sheet, _ = st.columns([1, 2])
    with col_sel_sheet:
        selected_tab = st.selectbox(
            "📋 Kies een Sheet om te bekijken/bewerken:",
            tab_names,
            key="sheet_selector_main",
        )

    st.subheader(f"Sheet: {selected_tab}")
    display_df = make_arrow_compatible(sheets[selected_tab])

    column_config = {}
    if selected_tab == skp_key:
        column_config[LOCK_ALL_COL] = st.column_config.CheckboxColumn(
            "🔒 All",
            help="Zet direct alle 5 taken voor deze wedstrijd vast.",
            default=False,
        )
        for l_col in LOCK_COLS:
            column_config[l_col] = st.column_config.CheckboxColumn(
                f"🔒 {l_col.replace('Lock ', '')}",
                help="Vink aan om deze specifieke taak vast te zetten.",
                default=False,
            )

    edited_df = st.data_editor(
        display_df,
        num_rows="dynamic",
        width="stretch",
        key=f"editor_{selected_tab}",
        column_config=column_config,
    )

    if not edited_df.equals(sheets[selected_tab]):
        if selected_tab == skp_key:
            edited_df = handle_lock_all_synchronization(sheets[selected_tab], edited_df)
        sheets[selected_tab] = edited_df
        sheets = update_player_stats(sheets)
        st.rerun()

    # --- MELDINGENVENSTER VOOR DIVISIE WIJZIGINGEN & VERVANGEN SCHEIDSRECHTERS ---
    if st.session_state.get("division_change_logs"):
        with st.expander("🔄 Wijzigingenoverzicht: Vervangen taken & scheidsrechters na divisiewijziging", expanded=True):
            st.info("De divisie van een team is gewijzigd. Hieronder staan de posities die automatisch zijn aangepast of vervangen om aan de regels te voldoen:")
            for log in st.session_state["division_change_logs"]:
                st.markdown(
                    f"- **{log['match']}** | **{log['task']}**: `{log['old_person']}` ➔ **{log['new_person']}** "
                    f"*(Reden: {log['reason']})*"
                )

    # --- MELDINGENVENSTER VOOR OPENSTAANDE PLEKKEN ---
    if st.session_state.get("assignment_warnings"):
        with st.expander("⚠️ Meldingenoverzicht: Waarom scheidsrechterplekken openstaan", expanded=False):
            st.error("Niet alle posities konden automatisch worden ingedeeld:")
            for w in st.session_state["assignment_warnings"]:
                st.markdown(f"**🏀 {w['match']}** - *{w['slot']}*:")
                for r_line in w["reasons"]:
                    st.write(f"- {r_line}")

    # --- REGELVALIDATIE RESULTATEN VENSTER ---
    if st.session_state.get("validation_results") is not None:
        v_issues = st.session_state["validation_results"]
        if not v_issues:
            st.success("✅ **Het rooster voldoet aan alle regels!** Geen dubbele boekingen, diploma-conflicten of openstaande taken.")
        else:
            with st.expander(f"❌ **Regelvalidatie: {len(v_issues)} knelpunten gevonden**", expanded=False):
                st.error("De volgende toewijzingen of posities voldoen niet aan de regels van het systeem:")
                for issue in v_issues:
                    st.markdown(f"- {issue}")

    # =========================================================================
    # ZIJBALK STRUCTUUR
    # =========================================================================

    # --- 1. MENU: LEDENBEHEER (INCL. TEAMBEHEER / DIVISIE WIJZIGEN) ---
    st.sidebar.divider()
    with st.sidebar.expander("👤 1. Ledenbeheer", expanded=False):
        if players_key in sheets:
            p_df_manage = sheets[players_key]
            available_teams_list = get_all_available_teams(sheets)
            tantalus_teams_list = get_all_tantalus_teams(sheets)
            diploma_options = ["Geen", "BS1", "BS2", "BS3", "L3", "L4"]
            season_options = ["Full season", "1st half season", "2nd half season"]

            # 1.1 TEAMDIVISIE WIJZIGEN (PROMOTIE / DEGRADATIE)
            with st.expander("🏆 Teamdivisie wijzigen", expanded=False):
                st.caption("Wijzig de divisie van een team bij promotie of degradatie. Sheets en taken worden direct bijgewerkt.")
                team_choices_for_div = tantalus_teams_list if tantalus_teams_list else available_teams_list
                if team_choices_for_div:
                    chosen_div_team = st.selectbox(
                        "Kies team:",
                        team_choices_for_div,
                        key="sel_team_for_div_change"
                    )

                    curr_team_div_num = determine_division_for_team(chosen_div_team, tantalus_div_map)
                    div_options_list = [f"Division {i}" for i in range(1, 7)]
                    curr_div_idx = (curr_team_div_num - 1) if (1 <= curr_team_div_num <= 6) else 4

                    new_div_selected = st.selectbox(
                        "Kies nieuwe divisie:",
                        div_options_list,
                        index=curr_div_idx,
                        key="sel_new_div_for_team"
                    )

                    new_div_num_val = int(new_div_selected.replace("Division ", "").strip())

                    if st.button("💾 Pas divisie aan & update rooster", key="btn_apply_div_change"):
                        sheets, aff_indices = update_team_division_across_sheets(
                            sheets, chosen_div_team, new_div_num_val
                        )
                        sheets, change_logs = reassign_invalid_referees_after_division_change(
                            sheets, aff_indices
                        )
                        st.session_state["division_change_logs"] = change_logs
                        st.session_state["validation_results"] = None
                        if change_logs:
                            st.success(f"Divisie van {chosen_div_team} gewijzigd naar {new_div_selected}! {len(change_logs)} ta(a)k(en) automatisch vervangen.")
                        else:
                            st.success(f"Divisie van {chosen_div_team} gewijzigd naar {new_div_selected}! Huidige scheidsrechters voldeden reeds aan de regels.")
                        st.rerun()
                else:
                    st.info("Geen teams gevonden.")

            # 1.2 LID TOEVOEGEN
            with st.expander("➕ Lid toevoegen", expanded=False):
                with st.form("form_add_member_unified"):
                    new_first = st.text_input("Voornaam:")
                    new_last = st.text_input("Achternaam:")
                    new_dip = st.selectbox("Scheidsrechter Diploma:", diploma_options)
                    new_season = st.selectbox("Seizoenshelft:", season_options)
                    new_team = st.selectbox("Team:", available_teams_list, index=0 if available_teams_list else None)
                    btn_add = st.form_submit_button("➕ Lid toevoegen & Rooster updaten")

                    if btn_add:
                        if not new_last.strip():
                            st.error("Achternaam is verplicht!")
                        else:
                            new_row = {
                                "First name": new_first.strip(),
                                "Last name": new_last.strip(),
                                "Diploma": "" if new_dip == "Geen" else new_dip,
                                "Full/ half season": new_season,
                                "Team": new_team,
                                "Committee": "",
                                "Extra": "",
                                "Extra points": 0.0,
                                "Referee": 0,
                                "Table duty": 0,
                                "Total points": 0.0,
                            }
                            p_df_manage = insert_player_into_team(p_df_manage, new_team, new_row)
                            sheets[players_key] = standardize_players_df(p_df_manage)
                            sheets = update_player_stats(sheets)
                            sheets, warns = auto_reassign_future_schedule(sheets, days_ahead=7)
                            st.session_state["assignment_warnings"] = warns
                            st.success(f"Lid {new_first} {new_last} toegevoegd aan team {new_team}! Rooster geüpdatet vanaf 7 dagen.")
                            st.rerun()

            # 1.3 LID WIJZIGEN
            with st.expander("✏️ Lid wijzigen", expanded=False):
                team_filter_edit = st.selectbox(
                    "Kies team:",
                    available_teams_list,
                    key="sel_filter_team_edit_tab"
                )

                valid_members_edit = []
                for i, r in p_df_manage.iterrows():
                    l_val = str(r.get("Last name", "")).strip()
                    if l_val in ["", "nan", "None"]:
                        continue
                    f_val = str(r.get("First name", "")).strip()
                    member_team = get_member_actual_team(p_df_manage, i)

                    if clean_team_code(member_team) == clean_team_code(team_filter_edit):
                        valid_members_edit.append((i, f"{f_val} {l_val} ({member_team})"))

                if valid_members_edit:
                    chosen_idx = st.selectbox(
                        "Kies lid:",
                        [item[0] for item in valid_members_edit],
                        format_func=lambda x: dict(valid_members_edit)[x],
                        key="sel_edit_member_tab"
                    )

                    curr_dip = str(p_df_manage.at[chosen_idx, "Diploma"]).strip()
                    dip_idx = diploma_options.index(curr_dip) if curr_dip in diploma_options else 0
                    curr_season = str(p_df_manage.at[chosen_idx, "Full/ half season"]).strip()
                    seas_idx = season_options.index(curr_season) if curr_season in season_options else 0
                    curr_actual_team = get_member_actual_team(p_df_manage, chosen_idx)

                    team_idx = 0
                    for t_i, t_name in enumerate(available_teams_list):
                        if clean_team_code(t_name) == clean_team_code(curr_actual_team):
                            team_idx = t_i
                            break

                    with st.form("form_edit_member_tab"):
                        edit_dip = st.selectbox("Scheidsrechter Diploma:", diploma_options, index=dip_idx)
                        edit_season = st.selectbox("Seizoenshelft:", season_options, index=seas_idx)
                        edit_team = st.selectbox("Nieuw Team:", available_teams_list, index=team_idx)
                        btn_update = st.form_submit_button("💾 Wijziging opslaan & Rooster updaten")

                        if btn_update:
                            p_df_manage.at[chosen_idx, "Diploma"] = "" if edit_dip == "Geen" else edit_dip
                            p_df_manage.at[chosen_idx, "Full/ half season"] = edit_season
                            p_df_manage.at[chosen_idx, "Team"] = edit_team
                            sheets[players_key] = standardize_players_df(p_df_manage)
                            sheets = update_player_stats(sheets)
                            sheets, warns = auto_reassign_future_schedule(sheets, days_ahead=7)
                            st.session_state["assignment_warnings"] = warns
                            st.success("Gegevens gewijzigd! Rooster geüpdatet vanaf 7 dagen.")
                            st.rerun()
                else:
                    st.info("Geen leden gevonden voor dit team.")

            # 1.4 LID VERWIJDEREN
            with st.expander("🗑️ Lid verwijderen", expanded=False):
                team_filter_del = st.selectbox(
                    "Kies team van de speler:",
                    available_teams_list,
                    key="sel_team_filter_del_tab"
                )

                members_to_del = []
                for i, r in p_df_manage.iterrows():
                    l_val = str(r.get("Last name", "")).strip()
                    if l_val in ["", "nan", "None"]:
                        continue
                    f_val = str(r.get("First name", "")).strip()
                    m_team = get_member_actual_team(p_df_manage, i)

                    if clean_team_code(m_team) == clean_team_code(team_filter_del):
                        members_to_del.append((i, f"{f_val} {l_val} ({m_team})"))

                if members_to_del:
                    del_idx = st.selectbox(
                        "Kies lid:",
                        [m[0] for m in members_to_del],
                        format_func=lambda x: dict(members_to_del)[x],
                        key="sel_member_del_tab"
                    )

                    del_type = st.radio(
                        "Type actie:",
                        ["Alleen uit team halen", "Volledig uit vereniging verwijderen"],
                        key="radio_del_type_tab"
                    )

                    if st.button("❌ Verwijderen uit selectie", key="btn_exec_del_tab"):
                        if del_type == "Alleen uit team halen":
                            p_df_manage.at[del_idx, "Team"] = ""
                            msg = "Speler uit het team gehaald! Rooster direct opnieuw ingedeeld."
                        else:
                            p_df_manage = p_df_manage.drop(index=del_idx).reset_index(drop=True)
                            msg = "Speler definitief verwijderd! Rooster direct opnieuw ingedeeld."

                        sheets[players_key] = standardize_players_df(p_df_manage)
                        sheets = update_player_stats(sheets)
                        sheets, warns = auto_reassign_future_schedule(sheets, days_ahead=0)
                        st.session_state["assignment_warnings"] = warns
                        st.success(msg)
                        st.rerun()
                else:
                    st.info("Geen spelers gevonden voor dit team.")

    # --- 2. MENU: COMMISSIE WIJZIGEN ---
    st.sidebar.divider()
    if players_key in sheets and comm_key in sheets:
        comm_df = sheets[comm_key]
        col_comm = comm_df.columns[0]
        available_committees = [
            str(c).strip()
            for c in comm_df[col_comm].dropna().unique()
            if str(c).strip()
        ]
        players_df = standardize_players_df(sheets[players_key])
        team_player_groups = get_team_player_groups(players_df)

        with st.sidebar.expander("🛠️ 2. Commissie Wijzigen", expanded=False):
            for team_name, player_indices in team_player_groups.items():
                if not player_indices:
                    continue
                with st.expander(f"🏀 {team_name} ({len(player_indices)} spelers)", expanded=False):
                    for idx in player_indices:
                        f_name = str(players_df.at[idx, "First name"])
                        l_name = str(players_df.at[idx, "Last name"])
                        raw_comm = str(players_df.at[idx, "Committee"]).strip()
                        current_comms_list = (
                            []
                            if raw_comm in ["nan", "None", "none", ""]
                            else [
                                c.strip()
                                for c in raw_comm.split(",")
                                if c.strip() in available_committees
                            ]
                        )

                        selected_comms = st.multiselect(
                            f"{f_name} {l_name}",
                            available_committees,
                            default=current_comms_list,
                            key=f"comm_multi_{idx}_{f_name}_{l_name}",
                        )
                        players_df.at[idx, "Committee"] = ", ".join(selected_comms)

            if st.button("💾 Sla Commissies op & Synchroniseer", key="btn_save_comm_menu"):
                sheets[players_key] = players_df
                sheets = update_player_stats(sheets)
                st.session_state["action_feedback"] = (
                    "success",
                    "Commissies succesvol bijgewerkt!",
                )
                st.rerun()

    # --- 3. MENU: WEDSTRIJDBEHEER ---
    st.sidebar.divider()
    with st.sidebar.expander("📅 3. Wedstrijdbeheer", expanded=False):
        if skp_key in sheets:
            skp_df_games = sheets[skp_key]
            court_col = "Court" if "Court" in skp_df_games.columns else ("Veld" if "Veld" in skp_df_games.columns else "Court")
            dag_col = "Dag" if "Dag" in skp_df_games.columns else ("Day" if "Day" in skp_df_games.columns else None)
            tantalus_teams_list = get_all_tantalus_teams(sheets)

            # 3.1 WEDSTRIJD TOEVOEGEN
            with st.expander("➕ Wedstrijd Toevoegen", expanded=False):
                with st.form("form_add_game"):
                    m_date_input = st.date_input("Datum:", value=datetime.now().date())
                    m_time_input = st.time_input("Tijdstip:", value=time(19, 0))

                    if tantalus_teams_list:
                        m_home_input = st.selectbox("Thuis team (Tantalus):", tantalus_teams_list)
                    else:
                        m_home_input = st.text_input("Thuis team:", value="Tantalus MSE 1")

                    m_away_input = st.text_input("Uit team:")
                    m_court_input = st.text_input("Veld / Zaal (Court):", value="Veld 1")

                    btn_add_game = st.form_submit_button("➕ Voeg wedstrijd toe aan rooster")

                    if btn_add_game:
                        if not str(m_home_input).strip() or not m_away_input.strip():
                            st.error("Zowel thuis- als uitteam zijn verplicht!")
                        else:
                            day_abbr = DUTCH_DAYS.get(m_date_input.weekday(), "")
                            date_only = m_date_input.strftime("%d-%m-%Y")

                            sample_date = ""
                            if "Date" in skp_df_games.columns and not skp_df_games["Date"].dropna().empty:
                                sample_date = str(skp_df_games["Date"].dropna().iloc[0]).lower()

                            has_day_in_date_col = any(d in sample_date for d in ["ma", "di", "wo", "do", "vr", "za", "zo"])
                            if has_day_in_date_col:
                                formatted_date_entry = f"{day_abbr} {date_only}"
                            else:
                                formatted_date_entry = date_only

                            t_str = m_time_input.strftime("%H:%M")
                            d_num = determine_division_for_team(m_home_input.strip(), tantalus_div_map)

                            new_game_row = {
                                "Date": formatted_date_entry,
                                "Time": t_str,
                                "Home Team": str(m_home_input).strip(),
                                "Away Team": m_away_input.strip(),
                                "Division": f"Division {d_num}",
                                court_col: m_court_input.strip(),
                                "Referee 1": "",
                                "Referee 2": "",
                                "Scorer": "",
                                "Timer": "",
                                "24 sec operator": "",
                                LOCK_ALL_COL: False,
                                "Lock Ref 1": False,
                                "Lock Ref 2": False,
                                "Lock Scorer": False,
                                "Lock Timer": False,
                                "Lock 24s": False,
                            }
                            if dag_col:
                                new_game_row[dag_col] = day_abbr

                            for col in skp_df_games.columns:
                                if col not in new_game_row:
                                    new_game_row[col] = ""

                            skp_df_games = pd.concat([skp_df_games, pd.DataFrame([new_game_row])], ignore_index=True)
                            skp_df_games = sort_skp_schedule(skp_df_games)
                            sheets[skp_key] = make_arrow_compatible(skp_df_games)
                            sheets = update_player_stats(sheets)
                            st.success(f"Wedstrijd {m_home_input} vs {m_away_input} op {formatted_date_entry} om {t_str} chronologisch ingevoegd!")
                            st.rerun()

            # 3.2 WEDSTRIJD VERWIJDEREN
            with st.expander("🗑️ Wedstrijd Verwijderen", expanded=False):
                game_choices = []
                col_home_g = find_col(skp_df_games, ["home team", "thuis team", "home", "thuis"]) or "Home Team"
                col_away_g = find_col(skp_df_games, ["away team", "uit team", "away", "uit"]) or "Away Team"
                col_date_g = find_col(skp_df_games, ["date", "datum"]) or "Date"
                col_time_g = find_col(skp_df_games, ["time", "tijd"]) or "Time"

                for idx_g, row_g in skp_df_games.iterrows():
                    d_show = str(row_g.get(col_date_g, ""))
                    t_show = str(row_g.get(col_time_g, ""))
                    c_show = str(row_g.get(court_col, ""))
                    h_show = str(row_g.get(col_home_g, ""))
                    a_show = str(row_g.get(col_away_g, ""))
                    court_str = f" [{c_show}]" if c_show and c_show != "nan" else ""
                    game_choices.append((idx_g, f"{d_show} {t_show}{court_str} - {h_show} vs {a_show}"))

                if game_choices:
                    sel_game_del = st.selectbox(
                        "Kies te verwijderen wedstrijd:",
                        [g[0] for g in game_choices],
                        format_func=lambda x: dict(game_choices)[x],
                        key="sel_game_to_delete"
                    )

                    if st.button("🗑️ Verwijder deze wedstrijd", key="btn_confirm_del_game"):
                        skp_df_games = skp_df_games.drop(index=sel_game_del).reset_index(drop=True)
                        sheets[skp_key] = make_arrow_compatible(skp_df_games)
                        sheets = update_player_stats(sheets)
                        st.success("Wedstrijd succesvol verwijderd en taken bijgewerkt!")
                        st.rerun()
                else:
                    st.info("Geen wedstrijden beschikbaar in het rooster.")

    # --- 4. MENU: ROOSTER INDELEN ---
    st.sidebar.divider()
    target_match_indices = []
    start_auto_btn = False

    with st.sidebar.expander("🤖 4. Rooster Indelen", expanded=False):
        if skp_key in sheets:
            skp_df_ctrl = sheets[skp_key]

            assign_mode = st.radio(
                "Indelen per:",
                ["Per dag(en)", "Per regel (specifieke wedstrijd)"],
                key="assign_scope_mode",
            )

            if assign_mode == "Per dag(en)":
                if "Date" in skp_df_ctrl.columns:
                    unique_dates = [d for d in skp_df_ctrl["Date"].dropna().unique() if str(d).strip() not in ["", "nan", "None"]]

                    def toggle_all_days_sync():
                        new_val = st.session_state["select_all_assign_days"]
                        for d in unique_dates:
                            st.session_state[f"assign_date_chk_{str(d).strip()}"] = new_val

                    st.checkbox(
                        "Selecteer Alle Dagen",
                        value=False,
                        key="select_all_assign_days",
                        on_change=toggle_all_days_sync
                    )

                    chosen_dates = []
                    for d in unique_dates:
                        key_chk = f"assign_date_chk_{str(d).strip()}"
                        if key_chk not in st.session_state:
                            st.session_state[key_chk] = st.session_state["select_all_assign_days"]
                        chk = st.checkbox(str(d), key=key_chk)
                        if chk:
                            chosen_dates.append(d)

                    chosen_dates_norm = {normalize_date_str(d) for d in chosen_dates}
                    for idx_r in skp_df_ctrl.index:
                        d_row_norm = normalize_date_str(skp_df_ctrl.at[idx_r, "Date"])
                        if d_row_norm in chosen_dates_norm:
                            target_match_indices.append(idx_r)
            else:
                row_choices_assign = [
                    (
                        idx_r,
                        f"Rij {idx_r + 1}: {r.get('Date', '')} ({r.get('Time', '')}) - {r.get('Home Team', '')} vs {r.get('Away Team', '')}",
                    )
                    for idx_r, r in skp_df_ctrl.iterrows()
                ]
                if row_choices_assign:
                    sel_row_idx = st.selectbox(
                        "Kies wedstrijd:",
                        [r[0] for r in row_choices_assign],
                        format_func=lambda x: dict(row_choices_assign)[x],
                        key="sel_row_assign",
                    )
                    target_match_indices = [sel_row_idx]

            st.write("")
            max_daily_tasks = st.slider(
                "Maximaal aantal taken per speler per dag:",
                min_value=1,
                max_value=4,
                value=1,
                step=1,
                help="Kies hoeveel taken iemand maximaal op één dag mag uitvoeren.",
                key="slider_max_daily_tasks"
            )

            start_auto_btn = st.button("🚀 Start indeling voor selectie", key="btn_start_assign")

    if start_auto_btn:
        if not target_match_indices:
            st.warning("Geen wedstrijden geselecteerd. Controleer de datumvinkjes.")
            st.stop()

        sheets, warns = run_assignment_core(sheets, target_match_indices, max_daily_tasks=max_daily_tasks, preserve_manual=False)
        st.session_state["indeling_gedaan"] = True
        st.session_state["assignment_warnings"] = warns
        st.session_state["validation_results"] = None
        if warns:
            st.warning("Indeling voltooid, maar er zijn openstaande posities. Bekijk het overzicht hierboven.")
        else:
            st.success("Indeling succesvol en evenwichtig uitgevoerd (max. 16 punten per lid)!")
        st.rerun()

    # --- 5. MENU: ROOSTER VALIDEREN ---
    st.sidebar.divider()
    with st.sidebar.expander("🔍 5. Rooster Valideren", expanded=False):
        st.caption("Controleer of de huidige indeling voldoet aan alle regels (dubbele boekingen, licenties, speeltijden, max. punten).")
        if st.button("🔍 Valideer Rooster", key="btn_validate_rules"):
            val_issues = validate_schedule_rules(sheets)
            st.session_state["validation_results"] = val_issues
            st.rerun()

    # --- 6. MENU: ROOSTER WISSEN ---
    st.sidebar.divider()
    with st.sidebar.expander("🗑️ 6. Rooster Wissen", expanded=False):
        if skp_key in sheets:
            skp_df_clear = sheets[skp_key]
            clear_mode = st.radio(
                "Wat wil je wissen?",
                ["Hele rooster wissen", "Indeling wissen per dag", "Wissen per regel"],
                key="clear_mode_selection",
            )

            if clear_mode == "Hele rooster wissen":
                if st.button("🗑️ Wis het hele rooster (excl. 🔒)", key="btn_clear_all_grid"):
                    for t_c in TASK_COLS:
                        l_col = LOCK_MAP[t_c]
                        if t_c in skp_df_clear.columns:
                            for i in skp_df_clear.index:
                                is_locked = (
                                    bool(skp_df_clear.at[i, l_col])
                                    if l_col in skp_df_clear.columns
                                    else False
                                )
                                if not is_locked:
                                    skp_df_clear.at[i, t_c] = ""
                    sheets[skp_key] = make_arrow_compatible(skp_df_clear)
                    sheets = update_player_stats(sheets)
                    st.session_state["assignment_warnings"] = []
                    st.session_state["validation_results"] = None
                    st.session_state["division_change_logs"] = []
                    st.session_state["auto_assigned_cells"] = set()
                    st.rerun()

            elif clear_mode == "Indeling wissen per dag":
                if "Date" in skp_df_clear.columns:
                    clear_dates_list = list(skp_df_clear["Date"].dropna().unique())
                    selected_clear_date = st.selectbox(
                        "Selecteer dag:", clear_dates_list, key="sel_clear_date_box"
                    )
                    if st.button(f"🗑️ Wis {selected_clear_date} (excl. 🔒)", key="btn_clear_day_act"):
                        d_reset_norm = normalize_date_str(selected_clear_date)
                        mask_reset = skp_df_clear["Date"].apply(normalize_date_str) == d_reset_norm
                        for idx_r in skp_df_clear[mask_reset].index:
                            for t_c in TASK_COLS:
                                l_col = LOCK_MAP[t_c]
                                if t_c in skp_df_clear.columns:
                                    is_locked = (
                                        bool(skp_df_clear.at[idx_r, l_col])
                                        if l_col in skp_df_clear.columns
                                        else False
                                    )
                                    if not is_locked:
                                        skp_df_clear.at[idx_r, t_c] = ""
                        sheets[skp_key] = make_arrow_compatible(skp_df_clear)
                        sheets = update_player_stats(sheets)
                        st.session_state["assignment_warnings"] = []
                        st.session_state["validation_results"] = None
                        st.session_state["division_change_logs"] = []
                        st.rerun()

            else:
                row_choices_clear = [
                    (
                        idx_r,
                        f"Rij {idx_r + 1}: {r.get('Date', '')} ({r.get('Time', '')}) - {r.get('Home Team', '')} vs {r.get('Away Team', '')}",
                    )
                    for idx_r, r in skp_df_clear.iterrows()
                ]
                if row_choices_clear:
                    sel_row_to_clear = st.selectbox(
                        "Kies regel:",
                        [r[0] for r in row_choices_clear],
                        format_func=lambda x: dict(row_choices_clear)[x],
                        key="sel_row_clear_box",
                    )
                    if st.button(f"🗑️ Wis Rij {sel_row_to_clear + 1} (excl. 🔒)", key="btn_clear_row_act"):
                        for t_c in TASK_COLS:
                            l_col = LOCK_MAP[t_c]
                            if t_c in skp_df_clear.columns:
                                is_locked = (
                                    bool(skp_df_clear.at[sel_row_to_clear, l_col])
                                    if l_col in skp_df_clear.columns
                                    else False
                                )
                                if not is_locked:
                                    skp_df_clear.at[sel_row_to_clear, t_c] = ""
                        sheets[skp_key] = make_arrow_compatible(skp_df_clear)
                        sheets = update_player_stats(sheets)
                        st.session_state["assignment_warnings"] = []
                        st.session_state["validation_results"] = None
                        st.session_state["division_change_logs"] = []
                        st.rerun()

    # --- 7. MENU: OPSLAAN & DOWNLOADEN ---
    st.sidebar.divider()
    st.sidebar.header("💾 7. Opslaan & Downloaden")

    wb_download = openpyxl.load_workbook(
        io.BytesIO(st.session_state["file_bytes"])
    )

    for sheet_name, df in sheets.items():
        if sheet_name in wb_download.sheetnames:
            ws = wb_download[sheet_name]
            clean_df = df.copy()

            for col_l in ALL_LOCK_COLS:
                if col_l in clean_df.columns:
                    clean_df = clean_df.drop(columns=[col_l])

            header_row_idx = None
            col_name_to_col_idx = {}
            for r in range(1, min(15, ws.max_row + 1)):
                row_vals = [
                    str(ws.cell(row=r, column=c).value or "").strip().lower()
                    for c in range(1, ws.max_column + 1)
                ]
                if any(
                    x in row_vals
                    for x in ["referee 1", "scorer", "first name", "home team", "date"]
                ):
                    header_row_idx = r
                    for c in range(1, ws.max_column + 1):
                        val_str = str(ws.cell(row=r, column=c).value or "").strip()
                        if val_str:
                            col_name_to_col_idx[val_str.lower()] = c
                    break

            if header_row_idx is not None:
                max_r = max(ws.max_row, header_row_idx + len(clean_df) + 10)
                for r in range(header_row_idx + 1, max_r + 1):
                    for c in col_name_to_col_idx.values():
                        cell = ws.cell(row=r, column=c)
                        if not isinstance(cell, MergedCell):
                            cell.value = None

                for df_col in clean_df.columns:
                    col_key = str(df_col).strip().lower()
                    if col_key in col_name_to_col_idx:
                        c_idx = col_name_to_col_idx[col_key]
                        for row_offset, val in enumerate(clean_df[df_col]):
                            target_row = header_row_idx + 1 + row_offset
                            cell = ws.cell(row=target_row, column=c_idx)
                            if not isinstance(cell, MergedCell):
                                cell.value = (
                                    None
                                    if (
                                        pd.isna(val)
                                        or val == ""
                                        or str(val).lower() == "nan"
                                    )
                                    else val
                                )

    output_buffer = io.BytesIO()
    wb_download.save(output_buffer)
    output_buffer.seek(0)

    st.sidebar.download_button(
        label="📥 Download Geüpdatet Excel",
        data=output_buffer,
        file_name="SKP_Schedule_Updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("👈 Upload je Excel-bestand in het linker menu om te beginnen.")
