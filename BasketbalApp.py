import copy
from datetime import datetime, timedelta
import io
import json
import random
import re
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google.oauth2 import service_account
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

LOCK_COLS = ["Lock Ref 1", "Lock Ref 2",
             "Lock Scorer", "Lock Timer", "Lock 24s"]
TASK_COLS = ["Referee 1", "Referee 2", "Scorer", "Timer", "24 sec operator"]
LOCK_MAP = dict(zip(TASK_COLS, LOCK_COLS))
DRIVE_FILENAME = "SKP_Live_Database.xlsx"


# =========================================================================
# GOOGLE DRIVE HELPER FUNCTIES
# =========================================================================
def get_drive_service():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=["https://www.googleapis.com/auth/drive"]
        )
        return build("drive", "v3", credentials=creds)
    except Exception as e:
        st.error(f"Fout bij verbinden met Google Drive: {e}")
        return None


def get_drive_folder_id():
    return st.secrets.get("gdrive", {}).get("folder_id", None)


def load_file_from_gdrive():
    service = get_drive_service()
    folder_id = get_drive_folder_id()
    if not service or not folder_id:
        return None

    try:
        query = f"'{folder_id}' in parents and name = '{DRIVE_FILENAME}' and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get("files", [])
        if not files:
            return None

        file_id = files[0]["id"]
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return fh.getvalue()
    except Exception as e:
        st.sidebar.warning(f"Laden van Google Drive mislukt: {e}")
        return None


def upload_file_to_gdrive(file_bytes):
    service = get_drive_service()
    folder_id = get_drive_folder_id()
    if not service or not folder_id:
        return

    try:
        query = f"'{folder_id}' in parents and name = '{DRIVE_FILENAME}' and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get("files", [])

        media = MediaIoBaseUpload(
            io.BytesIO(file_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True,
        )

        if files:
            file_id = files[0]["id"]
            service.files().update(fileId=file_id, media_body=media).execute()
        else:
            file_metadata = {"name": DRIVE_FILENAME, "parents": [folder_id]}
            service.files().create(body=file_metadata, media_body=media, fields="id").execute()
    except Exception as e:
        st.sidebar.warning(f"Live opslaan naar Google Drive mislukt: {e}")


def ensure_lock_columns(df):
    df_res = df.copy()
    for col in LOCK_COLS:
        if col not in df_res.columns:
            df_res[col] = False
        else:
            df_res[col] = df_res[col].fillna(False).astype(bool)
    return df_res


def make_arrow_compatible(df):
    df_clean = df.copy()
    for col in df_clean.columns:
        if col in LOCK_COLS:
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
    s = str(val).strip()
    m = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})", s)
    if m:
        p1, p2, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        if p1 > 12:
            d, m_val = p1, p2
        else:
            d, m_val = p2, p1
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


def normalize_time_str(val):
    if pd.isna(val) or val is None or str(val).strip() in ["", "nan", "None"]:
        return "00:00"
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%H:%M")
    s = str(val).strip()
    m = re.search(r"(\d{1,2}):(\d{2})", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return s


def calculate_comm_points(comm_str, comm_points_map):
    if not comm_str or pd.isna(comm_str) or str(comm_str).strip() in ["", "nan", "None"]:
        return 0.0
    comms = [c.strip() for c in str(comm_str).split(",") if c.strip()]
    return sum(comm_points_map.get(c, 0.0) for c in comms)


def standardize_players_df(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_last = find_col(df, ["last name", "lastname", "achternaam",
                        "last"], fallback_index=1 if len(df.columns) > 1 else None)
    col_first = find_col(df, ["first name", "firstname", "voornaam",
                         "first"], fallback_index=0 if len(df.columns) > 0 else None)
    col_dip = find_col(df, ["diploma", "licentie", "certificaat",
                       "niveau"], fallback_index=2 if len(df.columns) > 2 else None)
    col_comm = find_col(df, ["committee", "commissie"])
    col_extra = find_col(df, ["extra", "opmerking", "status", "notes"])
    col_team = find_col(df, ["team", "teamnaam", "spelend team"])
    col_season = find_col(
        df, ["full/ half season", "full/half season", "season", "seizoen", "half season", "half"])
    col_extra_pts = find_col(
        df, ["extra points", "extra punten", "commissie punten", "comm points"])
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
        df["Extra points"] = pd.to_numeric(
            df["Extra points"], errors="coerce").fillna(0.0)

    if "Referee" not in df.columns:
        df["Referee"] = 0
    if "Table duty" not in df.columns:
        df["Table duty"] = 0
    if "Total points" not in df.columns:
        df["Total points"] = df["Extra points"]
    else:
        df["Total points"] = pd.to_numeric(
            df["Total points"], errors="coerce").fillna(df["Extra points"])

    return df


def standardize_skp_df(df, div_map=None):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_home = find_col(df, ["home team", "thuis team", "home", "thuis"])
    col_away = find_col(df, ["away team", "uit team", "away", "uit"])
    col_date = find_col(df, ["date", "datum"])
    col_time = find_col(df, ["time", "tijd"])
    col_div = find_col(
        df, ["division", "divisie", "poule", "klasse", "league", "div"])

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

    if renames:
        df = df.rename(columns=renames)

    if div_map is not None:
        div_vals = []
        for _, r in df.iterrows():
            h_team = str(r.get("Home Team", ""))
            d_num = determine_division_for_team(h_team, div_map)
            div_vals.append(f"Division {d_num}")
        df["Division"] = div_vals
    elif "Division" not in df.columns:
        df["Division"] = "Division 5"

    df = ensure_lock_columns(df)
    return df


def build_team_busy_slots(all_games_df):
    busy_slots = set()
    if all_games_df is not None and not all_games_df.empty:
        for _, row in all_games_df.iterrows():
            d_val = normalize_date_str(row.get("Date"))
            t_val = normalize_time_str(row.get("Time"))
            h_team = str(row.get("Home Team", "")).lower()
            a_team = str(row.get("Away Team", "")).lower()

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


def is_player_eligible_for_date(season_type, date_val):
    if not season_type:
        return True
    st_clean = str(season_type).strip().lower()
    if "full" in st_clean or "heel" in st_clean:
        return True

    d_obj = parse_date_obj(date_val)
    if not d_obj:
        return True

    is_first_half = d_obj.month in [8, 9, 10, 11, 12, 1]
    if "first" in st_clean or "1st" in st_clean or "1e" in st_clean:
        return is_first_half
    elif "second" in st_clean or "2nd" in st_clean or "2e" in st_clean:
        return not is_first_half
    return True


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
        col_p = committees_df.columns[-1] if len(
            committees_df.columns) > 1 else None
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
            actual_team = t_explicit if (t_explicit and t_explicit not in [
                                         "nan", "", "none", "None"]) else current_team
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
            actual = t if (t and t.lower() not in [
                           "nan", "none", ""]) else current_team
            if i == target_idx:
                return actual if actual else "Overig / Geen Team"
    return "Overig / Geen Team"


def save_persistent_state(sheets_dict):
    if "file_bytes" not in st.session_state or st.session_state["file_bytes"] is None:
        return

    try:
        wb = openpyxl.load_workbook(io.BytesIO(st.session_state["file_bytes"]))
        for sheet_name, df in sheets_dict.items():
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                clean_df = df.copy()

                for col_l in LOCK_COLS:
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
                            val_str = str(
                                ws.cell(row=r, column=c).value or "").strip()
                            if val_str:
                                col_name_to_col_idx[val_str.lower()] = c
                        break

                if header_row_idx is not None:
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

        buf = io.BytesIO()
        wb.save(buf)
        updated_bytes = buf.getvalue()

        st.session_state["file_bytes"] = updated_bytes
        # Schrijf live weg naar Google Drive
        upload_file_to_gdrive(updated_bytes)
    except Exception as e:
        st.warning(f"Live opslaan mislukt: {e}")


def run_assignment_core(sheets_dict, target_match_indices, max_daily_tasks=1, preserve_manual=False):
    skp_key = find_sheet(sheets_dict, ["SKP", "Rooster"]) or "SKP"
    players_key = find_sheet(
        sheets_dict, ["Players skp", "Players", "Spelers"]) or "Players"
    all_games_key = find_sheet(
        sheets_dict, ["All games", "all games", "ALL GAMES"]) or "all games"
    div_key = find_sheet(sheets_dict, ["Divisions", "Divisies"]) or "Divisions"

    divisions_df = sheets_dict.get(div_key, pd.DataFrame())
    tantalus_div_map = build_division_map(divisions_df)

    skp_df = standardize_skp_df(sheets_dict[skp_key], div_map=tantalus_div_map)
    all_games_df = sheets_dict.get(all_games_key, pd.DataFrame())
    if not all_games_df.empty:
        all_games_df = standardize_skp_df(all_games_df)

    players_df = standardize_players_df(
        sheets_dict.get(players_key, pd.DataFrame()))

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
            player_team_map[full_n] = t_explicit if (t_explicit and t_explicit not in [
                                                     "nan", "", "none", "None"]) else current_team

    valid_players_dict = {}
    excluded_board_coach = set()
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
        season_type = str(
            p_row.get("Full/ half season", "Full season")).strip()

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
            is_locked = bool(skp_df.at[idx, l_col]
                             ) if l_col in skp_df.columns else False
            curr_v = str(skp_df.at[idx, col_c]).strip()
            is_manual_entry = (preserve_manual and curr_v not in [
                               "", "nan", "None", "x"] and (idx, col_c) not in auto_assigned)

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
                player_day_counts[name][d_val] = player_day_counts[name].get(
                    d_val, 0) + 1
        for col in ["Scorer", "Timer", "24 sec operator"]:
            name = str(skp_df.at[idx_row, col]).strip()
            if name in valid_players_dict:
                table_tasks_counter[name] += 1
                duty_specific_counter[name][col] = duty_specific_counter[name].get(
                    col, 0) + 1
                player_busy_times[name].add((d_val, t_val))
                player_day_counts[name][d_val] = player_day_counts[name].get(
                    d_val, 0) + 1

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
                is_locked = bool(
                    skp_df.at[idx, l_col]) if l_col in skp_df.columns else False
                if not is_locked:
                    skp_df.at[idx, ref_col] = "x"
        else:
            for ref_col in ["Referee 1", "Referee 2"]:
                l_col = LOCK_MAP[ref_col]
                is_locked = bool(
                    skp_df.at[idx, l_col]) if l_col in skp_df.columns else False
                if is_locked:
                    continue

                curr_val = str(skp_df.at[idx, ref_col]).strip()
                if curr_val.lower() == "x" or (curr_val and curr_val not in ["", "None", "nan"]):
                    continue

                other_ref_col = "Referee 2" if ref_col == "Referee 1" else "Referee 1"
                other_ref = str(skp_df.at[idx, other_ref_col]).strip()
                other_ref_dip = valid_players_dict.get(
                    other_ref, {}).get("Diploma", "")
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
                        fallback_dips = {"BS2", "BS1"}
                        emergency_dips = {"BS1"}
                    elif other_ref_dip == "BS2":
                        primary_dips = {"BS3", "BS2"}
                        fallback_dips = {"BS2", "BS1"}
                        emergency_dips = {"BS1"}
                    else:
                        primary_dips = {"BS3", "BS2"}
                        fallback_dips = {"BS2", "BS1"}
                        emergency_dips = {"BS1"}
                elif div_num == 4:
                    primary_dips = {"BS2", "BS1"}
                    fallback_dips = {"BS1"}
                    emergency_dips = {"BS1"}
                else:
                    primary_dips = {"BS1", "BS2"}
                    fallback_dips = {"BS1", "BS2"}

                all_ref_pts = [
                    p["Base_Points"] + (ref_tasks_counter[n]
                                        * 2) + (table_tasks_counter[n] * 1)
                    for n, p in valid_players_dict.items() if p["Has_Diploma"]
                ]
                avg_ref_pts = sum(all_ref_pts) / \
                    len(all_ref_pts) if all_ref_pts else 0.0

                ref_cands = []
                for p_name, p_info in valid_players_dict.items():
                    if not p_info["Has_Diploma"]:
                        continue

                    if not is_physically_free(p_name):
                        continue

                    day_cnt = player_day_counts[p_name].get(d_norm, 0)
                    if day_cnt >= max_daily_tasks:
                        continue

                    curr_pts = p_info["Base_Points"] + (
                        ref_tasks_counter[p_name] * 2) + (table_tasks_counter[p_name] * 1)
                    if (curr_pts + 2.0) > 16.0:
                        continue

                    dip = p_info["Diploma"]
                    is_aurelie = p_info["Is_Aurelie"]

                    under_12 = 0 if curr_pts < 12.0 else 1
                    is_overloaded = (curr_pts > (avg_ref_pts + 6.0))

                    if div_num <= 3:
                        if dip in ["L4", "L3"]:
                            dip_rank = 1
                        elif dip == "BS3":
                            dip_rank = 2
                        elif dip == "BS2":
                            dip_rank = 3
                        else:
                            dip_rank = 4
                    else:
                        dip_rank = 0

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
                        "total_points": curr_pts,
                        "dip_rank": dip_rank,
                        "under_12": under_12,
                        "ref_tasks": ref_tasks_counter[p_name],
                    })

                ref_cands.sort(
                    key=lambda x: (
                        x["day_count"],
                        x["aurelie_prio"],
                        x["high_div_boost"],
                        x["tier"],
                        x["total_points"],
                        x["dip_rank"],
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
                    player_day_counts[chosen][d_norm] = player_day_counts[chosen].get(
                        d_norm, 0) + 1
                    ref_tasks_counter[chosen] += 1
                    auto_assigned.add((idx, ref_col))
                else:
                    reasons = []
                    for p_name, p_info in valid_players_dict.items():
                        if not p_info["Has_Diploma"]:
                            continue
                        p_team = p_info["Team"]
                        p_day_cnt = player_day_counts[p_name].get(d_norm, 0)
                        p_pts = p_info["Base_Points"] + (ref_tasks_counter[p_name] * 2) + (
                            table_tasks_counter[p_name] * 1)

                        if (p_pts + 2.0) > 16.0:
                            reasons.append(
                                f"**{p_name}** ({p_info['Diploma']}): Bereikt maximum van 16 punten ({p_pts} pnt).")
                        elif (d_norm, t_norm) in player_busy_times[p_name]:
                            reasons.append(
                                f"**{p_name}** ({p_info['Diploma']}): Heeft al een taak om {t_norm}.")
                        elif is_player_playing(p_team, m_date, m_time, home_team, away_team, busy_game_slots):
                            reasons.append(
                                f"**{p_name}** ({p_info['Diploma']}): Speelt zelf met team *{p_team}*.")
                        elif p_day_cnt >= max_daily_tasks:
                            reasons.append(
                                f"**{p_name}** ({p_info['Diploma']}): Daglimiet van {max_daily_tasks} ta(a)k(en) bereikt.")
                        elif not is_player_eligible_for_date(p_info["Season_Type"], m_date):
                            reasons.append(
                                f"**{p_name}**: Speelt halve seizoen ({p_info['Season_Type']}).")

                    for exc_name in excluded_board_coach:
                        reasons.append(
                            f"**{exc_name}**: Vrijgesteld van taken (Board / Coach).")

                    if not reasons:
                        reasons.append(
                            "Geen actieve gediplomeerde arbiters beschikbaar (of allen overschrijden 16 punten).")

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
            is_locked = bool(skp_df.at[idx, l_col]
                             ) if l_col in skp_df.columns else False
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

                curr_pts = p_info["Base_Points"] + \
                    (ref_tasks_counter[p_name] * 2) + \
                    (table_tasks_counter[p_name] * 1)
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
                player_day_counts[chosen][d_norm] = player_day_counts[chosen].get(
                    d_norm, 0) + 1
                table_tasks_counter[chosen] += 1
                duty_specific_counter[chosen][col] = duty_specific_counter[chosen].get(
                    col, 0) + 1
                auto_assigned.add((idx, col))

    st.session_state["auto_assigned_cells"] = auto_assigned
    sheets_dict[skp_key] = make_arrow_compatible(skp_df)
    sheets_dict = update_player_stats(sheets_dict)
    save_persistent_state(sheets_dict)
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
        res_sheets, res_warns = run_assignment_core(
            sheets_dict, target_indices, max_daily_tasks=current_max_tasks, preserve_manual=True)
        save_persistent_state(res_sheets)
        return res_sheets, res_warns
    return sheets_dict, []


# =========================================================================
# INITIALISATIE: BESTAND LADEN UIT GOOGLE DRIVE OF SESSIE
# =========================================================================
def parse_and_load_bytes(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    raw_sheets = {}
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        raw_sheets[sheet] = df.loc[:, ~df.columns.astype(
            str).str.contains("^Unnamed")]

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
                    df[col] = df[col].fillna("").astype(
                        str).replace({"nan": "", "None": ""})
        orig_sheets[sheet] = df
    return orig_sheets


# Check bij app-start of er al een database op Google Drive staat
if "sheets" not in st.session_state:
    drive_bytes = load_file_from_gdrive()
    if drive_bytes:
        st.session_state["file_bytes"] = drive_bytes
        loaded_sheets = parse_and_load_bytes(drive_bytes)
        st.session_state["original_sheets"] = copy.deepcopy(loaded_sheets)
        st.session_state["sheets"] = copy.deepcopy(loaded_sheets)
        st.session_state["indeling_gedaan"] = False
        st.session_state["assignment_warnings"] = []
        st.session_state["auto_assigned_cells"] = set()
        st.sidebar.success("☁️ Live-rooster geladen vanuit Google Drive!")

uploaded_file = st.sidebar.file_uploader(
    "Upload handmatig Excel (optioneel)", type=["xlsx"], key="main_file_uploader"
)

# Handmatige upload overschrijft Google Drive bestand
if uploaded_file is not None:
    f_bytes = uploaded_file.getvalue()
    curr_fn = getattr(uploaded_file, "name", "excel")

    if st.session_state.get("last_uploaded_filename") != curr_fn:
        st.session_state["file_bytes"] = f_bytes
        st.session_state["last_uploaded_filename"] = curr_fn

        new_sheets = parse_and_load_bytes(f_bytes)
        st.session_state["original_sheets"] = copy.deepcopy(new_sheets)
        st.session_state["sheets"] = copy.deepcopy(new_sheets)
        st.session_state["indeling_gedaan"] = False
        st.session_state["assignment_warnings"] = []
        st.session_state["auto_assigned_cells"] = set()

        upload_file_to_gdrive(f_bytes)
        st.rerun()


# =========================================================================
# HOOFDPROGRAMMA & ZIJBALK
# =========================================================================
if "sheets" in st.session_state:
    sheets = st.session_state["sheets"]

    players_key = find_sheet(
        sheets, ["Players skp", "Players", "Spelers"]) or "Players"
    skp_key = find_sheet(sheets, ["SKP", "Rooster"]) or "SKP"
    comm_key = find_sheet(sheets, ["Committees", "Commissies"]) or "Committees"
    all_games_key = find_sheet(
        sheets, ["All games", "all games", "ALL GAMES"]) or "all games"
    div_key = find_sheet(sheets, ["Divisions", "Divisies"]) or "Divisions"

    divisions_df = sheets.get(div_key, pd.DataFrame())
    tantalus_div_map = build_division_map(divisions_df)

    if skp_key in sheets:
        sheets[skp_key] = standardize_skp_df(
            sheets[skp_key], div_map=tantalus_div_map)

    st.caption("☁️ **Google Drive Gekoppeld**: Alle aanpassingen en indelingen worden realtime gesynchroniseerd met Google Drive.")

    tab_names = list(sheets.keys())
    col_sel_sheet, _ = st.columns([2, 1])
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
        for l_col in LOCK_COLS:
            column_config[l_col] = st.column_config.CheckboxColumn(
                f"🔒 {l_col.replace('Lock ', '')}",
                help="Vink aan om deze taak vast te zetten.",
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
        sheets[selected_tab] = edited_df
        sheets = update_player_stats(sheets)
        save_persistent_state(sheets)
        st.rerun()

    # --- MELDINGENVENSTER EN VERGELIJKEN ORIGINEEL VS GEÜPDATET (IN TABBLADEN) ---
    st.divider()
    if st.session_state.get("assignment_warnings"):
        with st.expander("⚠️ Meldingenoverzicht: Waarom scheidsrechterplekken openstaan", expanded=True):
            st.error(
                "Niet alle scheidsrechterposities konden automatisch worden ingedeeld. Hieronder staan de gedetailleerde redenen:")
            for w in st.session_state["assignment_warnings"]:
                st.markdown(f"**🏀 {w['match']}** - *{w['slot']}*:")
                for r_line in w["reasons"]:
                    st.write(f"- {r_line}")

    st.subheader("🔍 Rooster Vergelijking: Origineel vs Geüpdatet")
    tab_orig, tab_updated = st.tabs(
        ["📄 Origineel Rooster", "✨ Geüpdatet Rooster"])

    with tab_orig:
        orig_skp = st.session_state.get(
            "original_sheets", {}).get(skp_key, pd.DataFrame())
        if not orig_skp.empty:
            st.dataframe(make_arrow_compatible(orig_skp),
                         height=550, width="stretch")
        else:
            st.info("Geen origineel rooster beschikbaar.")

    with tab_updated:
        if skp_key in sheets:
            st.dataframe(make_arrow_compatible(
                sheets[skp_key]), height=550, width="stretch")
        else:
            st.info("Geen geüpdatet rooster beschikbaar.")

    # =========================================================================
    # ZIJBALK STRUCTUUR: EXACT IN VOLGORDE (1 T/M 5)
    # =========================================================================

    # --- 1. MENU: LEDENBEHEER ---
    st.sidebar.divider()
    with st.sidebar.expander("👤 1. Ledenbeheer", expanded=False):
        if players_key in sheets:
            p_df_manage = sheets[players_key]
            available_teams_list = get_all_available_teams(sheets)
            diploma_options = ["Geen", "BS1", "BS2", "BS3", "L3", "L4"]
            season_options = ["Full season",
                              "1st half season", "2nd half season"]

            # 1.1 TUSSENMENU: LID TOEVOEGEN
            with st.expander("➕ Lid toevoegen", expanded=False):
                with st.form("form_add_member_unified"):
                    new_first = st.text_input("Voornaam:")
                    new_last = st.text_input("Achternaam:")
                    new_dip = st.selectbox(
                        "Scheidsrechter Diploma:", diploma_options)
                    new_season = st.selectbox("Seizoenshelft:", season_options)
                    new_team = st.selectbox(
                        "Team:", available_teams_list, index=0 if available_teams_list else None)
                    btn_add = st.form_submit_button(
                        "➕ Lid toevoegen & Rooster updaten")

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
                                "Total points": 0.0
                            }
                            p_df_manage = pd.concat(
                                [p_df_manage, pd.DataFrame([new_row])], ignore_index=True)
                            sheets[players_key] = standardize_players_df(
                                p_df_manage)
                            sheets = update_player_stats(sheets)
                            save_persistent_state(sheets)
                            sheets, warns = auto_reassign_future_schedule(
                                sheets, days_ahead=7)
                            st.session_state["assignment_warnings"] = warns
                            st.success(
                                f"Lid {new_first} {new_last} toegevoegd! Rooster geüpdatet vanaf 7 dagen.")
                            st.rerun()

            # 1.2 TUSSENMENU: LID WIJZIGEN
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
                        valid_members_edit.append(
                            (i, f"{f_val} {l_val} ({member_team})"))

                if valid_members_edit:
                    chosen_idx = st.selectbox(
                        "Kies lid:",
                        [item[0] for item in valid_members_edit],
                        format_func=lambda x: dict(valid_members_edit)[x],
                        key="sel_edit_member_tab"
                    )

                    curr_dip = str(
                        p_df_manage.at[chosen_idx, "Diploma"]).strip()
                    dip_idx = diploma_options.index(
                        curr_dip) if curr_dip in diploma_options else 0
                    curr_season = str(
                        p_df_manage.at[chosen_idx, "Full/ half season"]).strip()
                    seas_idx = season_options.index(
                        curr_season) if curr_season in season_options else 0
                    curr_actual_team = get_member_actual_team(
                        p_df_manage, chosen_idx)

                    team_idx = 0
                    for t_i, t_name in enumerate(available_teams_list):
                        if clean_team_code(t_name) == clean_team_code(curr_actual_team):
                            team_idx = t_i
                            break

                    with st.form("form_edit_member_tab"):
                        edit_dip = st.selectbox(
                            "Scheidsrechter Diploma:", diploma_options, index=dip_idx)
                        edit_season = st.selectbox(
                            "Seizoenshelft:", season_options, index=seas_idx)
                        edit_team = st.selectbox(
                            "Nieuw Team:", available_teams_list, index=team_idx)
                        btn_update = st.form_submit_button(
                            "💾 Wijziging opslaan & Rooster updaten")

                        if btn_update:
                            p_df_manage.at[chosen_idx,
                                           "Diploma"] = "" if edit_dip == "Geen" else edit_dip
                            p_df_manage.at[chosen_idx,
                                           "Full/ half season"] = edit_season
                            p_df_manage.at[chosen_idx, "Team"] = edit_team
                            sheets[players_key] = standardize_players_df(
                                p_df_manage)
                            sheets = update_player_stats(sheets)
                            save_persistent_state(sheets)
                            sheets, warns = auto_reassign_future_schedule(
                                sheets, days_ahead=7)
                            st.session_state["assignment_warnings"] = warns
                            st.success(
                                "Gegevens gewijzigd! Rooster geüpdatet vanaf 7 dagen.")
                            st.rerun()
                else:
                    st.info("Geen leden gevonden voor dit team.")

            # 1.3 TUSSENMENU: LID VERWIJDEREN
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
                        members_to_del.append(
                            (i, f"{f_val} {l_val} ({m_team})"))

                if members_to_del:
                    del_idx = st.selectbox(
                        "Kies lid:",
                        [m[0] for m in members_to_del],
                        format_func=lambda x: dict(members_to_del)[x],
                        key="sel_member_del_tab"
                    )

                    del_type = st.radio(
                        "Type actie:",
                        ["Alleen uit team halen",
                            "Volledig uit vereniging verwijderen"],
                        key="radio_del_type_tab"
                    )

                    if st.button("❌ Verwijderen uit selectie", key="btn_exec_del_tab"):
                        if del_type == "Alleen uit team halen":
                            p_df_manage.at[del_idx, "Team"] = ""
                            msg = "Speler uit het team gehaald! Rooster direct opnieuw ingedeeld."
                        else:
                            p_df_manage = p_df_manage.drop(
                                index=del_idx).reset_index(drop=True)
                            msg = "Speler definitief verwijderd! Rooster direct opnieuw ingedeeld."

                        sheets[players_key] = standardize_players_df(
                            p_df_manage)
                        sheets = update_player_stats(sheets)
                        save_persistent_state(sheets)
                        sheets, warns = auto_reassign_future_schedule(
                            sheets, days_ahead=0)
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
                with st.expander(f"🏀 {team_name} ({len(player_indices)} spelers)"):
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
                        players_df.at[idx, "Committee"] = ", ".join(
                            selected_comms)

            if st.button("💾 Sla Commissies op & Synchroniseer", key="btn_save_comm_menu"):
                sheets[players_key] = players_df
                sheets = update_player_stats(sheets)
                save_persistent_state(sheets)
                st.session_state["action_feedback"] = (
                    "success",
                    "Commissies succesvol bijgewerkt!",
                )
                st.rerun()

    # --- 3. MENU: ROOSTER INDELEN ---
    st.sidebar.divider()
    target_match_indices = []
    start_auto_btn = False

    with st.sidebar.expander("🤖 3. Rooster Indelen", expanded=True):
        if skp_key in sheets:
            skp_df_ctrl = sheets[skp_key]

            assign_mode = st.radio(
                "Indelen per:",
                ["Per dag(en)", "Per regel (specifieke wedstrijd)"],
                key="assign_scope_mode",
            )

            if assign_mode == "Per dag(en)":
                if "Date" in skp_df_ctrl.columns:
                    unique_dates = [d for d in skp_df_ctrl["Date"].dropna(
                    ).unique() if str(d).strip() not in ["", "nan", "None"]]

                    def toggle_all_days_sync():
                        new_val = st.session_state["select_all_assign_days"]
                        for d in unique_dates:
                            st.session_state[f"assign_date_chk_{str(d).strip()}"] = new_val

                    st.checkbox(
                        "Selecteer Alle Dagen",
                        value=True,
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

                    chosen_dates_norm = {
                        normalize_date_str(d) for d in chosen_dates}
                    for idx_r in skp_df_ctrl.index:
                        d_row_norm = normalize_date_str(
                            skp_df_ctrl.at[idx_r, "Date"])
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

            start_auto_btn = st.button(
                "🚀 Start indeling voor selectie", key="btn_start_assign")

    if start_auto_btn:
        if not target_match_indices:
            st.warning(
                "Geen wedstrijden geselecteerd. Controleer de datumvinkjes.")
            st.stop()

        sheets, warns = run_assignment_core(
            sheets, target_match_indices, max_daily_tasks=max_daily_tasks, preserve_manual=False)
        st.session_state["indeling_gedaan"] = True
        st.session_state["assignment_warnings"] = warns
        save_persistent_state(sheets)
        if warns:
            st.warning(
                "Indeling voltooid, maar er zijn openstaande posities. Bekijk het overzicht hierboven.")
        else:
            st.success(
                "Indeling succesvol en evenwichtig uitgevoerd (max. 16 punten per lid)!")
        st.rerun()

    # --- 4. MENU: ROOSTER WISSEN ---
    st.sidebar.divider()
    with st.sidebar.expander("🗑️ 4. Rooster Wissen", expanded=False):
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
                    st.session_state["auto_assigned_cells"] = set()
                    save_persistent_state(sheets)
                    st.rerun()

            elif clear_mode == "Indeling wissen per dag":
                if "Date" in skp_df_clear.columns:
                    clear_dates_list = list(
                        skp_df_clear["Date"].dropna().unique())
                    selected_clear_date = st.selectbox(
                        "Selecteer dag:", clear_dates_list, key="sel_clear_date_box"
                    )
                    if st.button(f"🗑️ Wis {selected_clear_date} (excl. 🔒)", key="btn_clear_day_act"):
                        d_reset_norm = normalize_date_str(selected_clear_date)
                        mask_reset = skp_df_clear["Date"].apply(
                            normalize_date_str) == d_reset_norm
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
                        save_persistent_state(sheets)
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
                                    bool(
                                        skp_df_clear.at[sel_row_to_clear, l_col])
                                    if l_col in skp_df_clear.columns
                                    else False
                                )
                                if not is_locked:
                                    skp_df_clear.at[sel_row_to_clear, t_c] = ""
                        sheets[skp_key] = make_arrow_compatible(skp_df_clear)
                        sheets = update_player_stats(sheets)
                        st.session_state["assignment_warnings"] = []
                        save_persistent_state(sheets)
                        st.rerun()

    # --- 5. MENU: OPSLAAN & DOWNLOADEN ---
    st.sidebar.divider()
    st.sidebar.header("💾 5. Opslaan & Downloaden")

    wb_download = openpyxl.load_workbook(
        io.BytesIO(st.session_state["file_bytes"])
    )

    for sheet_name, df in sheets.items():
        if sheet_name in wb_download.sheetnames:
            ws = wb_download[sheet_name]
            clean_df = df.copy()

            for col_l in LOCK_COLS:
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
                        val_str = str(
                            ws.cell(row=r, column=c).value or "").strip()
                        if val_str:
                            col_name_to_col_idx[val_str.lower()] = c
                    break

            if header_row_idx is not None:
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
    st.info("👈 Upload eenmalig je Excel-bestand in het linker menu om te beginnen, of configureer Google Drive.")
