# logic.py
import pandas as pd
import random
from config import *

class RosterEngine:
    def __init__(self):
        self.df = None
        self.week_columns = []
        self.availability_map = {} 
        self.initial_roster = {}   
        self.all_members = {} 

    def load_file(self, filepath):
        try:
            df_raw = pd.read_excel(filepath, header=None, keep_default_na=False)
            header_row = -1
            for i, row in df_raw.iterrows():
                row_str = [str(x).strip() for x in row.values]
                if "Name" in row_str:
                    header_row = i
                    break
            
            if header_row == -1: return False, "Could not find 'Name' column."
            
            self.df = pd.read_excel(filepath, header=header_row, keep_default_na=False)
            self.df.columns = self.df.columns.astype(str).str.replace('\n', ' ').str.strip()
            
            self._process_data()
            return True, "File Loaded Successfully"
            
        except Exception as e:
            return False, str(e)

    def _process_data(self):
        cols = self.df.columns
        inst_col = next((c for c in cols if "INSTRUMENT" in str(c).upper() or ("PIANO" in str(c).upper() and "DRUM" in str(c).upper())), None)
        filled_col = next((c for c in cols if "FILLED" in str(c).upper() or "✅" in str(c)), None)
        
        fwt_col, fph_col, fmc_col, fut_col = None, None, None, None
        
        def check_col(name, keys):
            if any(k in str(name).upper() for k in keys): return True
            if not self.df.empty:
                val = str(self.df[name].iloc[0]).upper()
                if any(k in val for k in keys): return True
            return False

        for c in cols:
            if c == inst_col or c == filled_col: continue
            if check_col(c, ["FWT", "WORSHIP"]): fwt_col = c
            elif check_col(c, ["FPH", "PRODUCTION", "HUB"]): fph_col = c
            elif check_col(c, ["FMC", "MC"]): fmc_col = c
            elif check_col(c, ["FUT", "USHER"]): fut_col = c

        self.week_columns = [c for c in cols if "Week" in c]
        self.availability_map = {week: {role: [] for role in ROLES_ORDER} for week in self.week_columns}

        for week in self.week_columns:
            self.availability_map[week]["Cleanup 1"] = CLEANUP_OPTIONS.copy()
            self.availability_map[week]["Cleanup 2"] = CLEANUP_OPTIONS.copy()

        def is_active(val):
            s = str(val).upper()
            return "Y" in s or "TRUE" in s or "YES" in s or "1" in s

        def get_capabilities(row):
            raw = str(row[inst_col]).upper().replace("\n", ",").replace("/", ",").replace("(", "").replace(")", "")
            caps = []
            for code in [x.strip() for x in raw.split(',')]:
                if code in INSTRUMENT_MAP: caps.append(INSTRUMENT_MAP[code])
                elif "PPT" in code: caps.append("PPT")
                elif "SOUND" in code: caps.append("Sound")
                elif "OBS" in code or "LIGHT" in code: caps.append("Lighting/OBS")
            
            if fph_col and is_active(row[fph_col]):
                if "Sound" not in caps: caps.append("Sound")
                if "PPT" not in caps: caps.append("PPT")
                if "Lighting/OBS" not in caps: caps.append("Lighting/OBS")
            if fmc_col and is_active(row[fmc_col]):
                if "MC" not in caps: caps.append("MC")
            if fut_col and is_active(row[fut_col]):
                if "Usher" not in caps: caps.append("Usher")
            return caps

        self.all_members = {}

        for idx, row in self.df.iterrows():
            if filled_col:
                val = str(row[filled_col]).upper()
                if not ("✅" in val or "TRUE" in val or "Y" in val or "1" in val): continue 
            
            original_name = str(row['Name']).strip()
            raw_caps = get_capabilities(row)
            
            is_md = "MD" in raw_caps
            clean_caps = raw_caps # Keep MD in caps for tracking
            display_name = original_name
            
            avail_str = ""
            for week in self.week_columns:
                status = str(row[week]).upper()
                if "N/A" in status or "NA" in status:
                    avail_str += "X"
                else:
                    avail_str += "O"
            
            self.all_members[display_name] = {"Roles": clean_caps, "AvailString": avail_str}

            for w_idx, week in enumerate(self.week_columns):
                if avail_str[w_idx] == "O":
                    for r in ROLES_ORDER:
                        if r == "MD": continue
                        if "Usher" in r:
                            if "Usher" in clean_caps: self.availability_map[week][r].append(display_name)
                        elif r in clean_caps:
                            self.availability_map[week][r].append(display_name)

    def generate_draft(self):
        self.initial_roster = {week: {} for week in self.week_columns}
        burnout = {name: 0 for name in self.all_members.keys()}
        last_week_played = {name: -1 for name in self.all_members.keys()}
        
        for w_idx, week in enumerate(self.week_columns):
            assigned_this_week = set() 
            sorted_roles = sorted(ROLES_ORDER, key=lambda r: len(self.availability_map[week][r]))
            
            # 1. Assign Standard Roles
            for role in sorted_roles:
                if role == "MD": continue 

                candidates = [p for p in self.availability_map[week][role] if p not in assigned_this_week]
                
                if candidates:
                    random.shuffle(candidates)
                    if "Cleanup" in role:
                        winner = candidates[0]
                        self.initial_roster[week][role] = winner
                        assigned_this_week.add(winner)
                    else:
                        candidates.sort(key=lambda p: (burnout.get(p, 0) * 10) + (50 if last_week_played.get(p) == (w_idx - 1) else 0))
                        winner = candidates[0]
                        self.initial_roster[week][role] = winner
                        assigned_this_week.add(winner)
                        burnout[winner] = burnout.get(winner, 0) + 1
                        last_week_played[winner] = w_idx
                else:
                    self.initial_roster[week][role] = ""

            # 2. Logic: Lock Bass if No Keys
            if not self.initial_roster[week].get("Piano"):
                if self.initial_roster[week].get("Bass"):
                    bassist = self.initial_roster[week]["Bass"]
                    self.initial_roster[week]["Bass"] = ""
                    if bassist in burnout: burnout[bassist] -= 1

            # 3. Logic: Auto-Fill MD
            md_candidate = ""
            for role in BAND_ROLES:
                person = self.initial_roster[week].get(role, "")
                if person and "MD" in self.all_members.get(person, {}).get("Roles", []):
                    md_candidate = person
                    break
            self.initial_roster[week]["MD"] = md_candidate