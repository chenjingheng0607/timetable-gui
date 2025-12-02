import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import random
import functools

# ==========================================
# CONFIGURATION
# ==========================================
ROLES_ORDER = [
    "Lead", "Vocal", "Piano", "Drum/Cajon", "Bass", "Guitar", 
    "PPT", "Sound", 
    "MC", 
    "Usher 1", "Usher 2", "Usher 3", 
    "Cleanup 1", "Cleanup 2"
]

CLEANUP_OPTIONS = ["LHW", "UF", "LB", "YGSS", "SJS", "PK"]

INSTRUMENT_MAP = {
    "WL": "Lead", "V": "Vocal", "P": "Piano", "G": "Guitar", 
    "B": "Bass", "D": "Drum/Cajon", "PPT": "PPT", 
    "S": "Sound", "SOUND": "Sound", "MC": "MC", "USHER": "Usher"
}

# --- CATEGORY DEFINITIONS ---
CATEGORY_CONFIG = {
    "Praise & Worship": {
        "roles": ["Lead", "Vocal", "Piano", "Drum/Cajon", "Bass", "Guitar"],
        "color": "#c00000", # Red
        "text_col": "white"
    },
    "FPH": {
        "roles": ["PPT", "Sound"],
        "color": "#0070c0", # Blue
        "text_col": "white"
    },
    "MC": {
        "roles": ["MC"],
        "color": "#7030a0", # Purple
        "text_col": "white"
    },
    "Usher": {
        "roles": ["Usher 1", "Usher 2", "Usher 3"],
        "color": "#ffc000", # Orange/Yellow
        "text_col": "black"
    },
    "LG": {
        "roles": ["Cleanup 1", "Cleanup 2"],
        "color": "#00b050", # Green
        "text_col": "white"
    }
}

# Map Role -> Category Info
ROLE_TO_CAT_MAP = {}
for cat, data in CATEGORY_CONFIG.items():
    for r in data["roles"]:
        ROLE_TO_CAT_MAP[r] = {"cat": cat, "color": data["color"]}

# ==========================================
# LOGIC CLASS
# ==========================================
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
            
            if fph_col and is_active(row[fph_col]):
                if "Sound" not in caps: caps.append("Sound")
                if "PPT" not in caps: caps.append("PPT")
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
            
            name = row['Name']
            caps = get_capabilities(row)
            
            avail_str = ""
            for week in self.week_columns:
                status = str(row[week]).upper()
                if "N/A" in status or "NA" in status:
                    avail_str += "X"
                else:
                    avail_str += "O"
            
            self.all_members[name] = {"Roles": caps, "AvailString": avail_str}

            for w_idx, week in enumerate(self.week_columns):
                if avail_str[w_idx] == "O":
                    for r in ROLES_ORDER:
                        if "Usher" in r:
                            if "Usher" in caps: self.availability_map[week][r].append(name)
                        elif r in caps:
                            self.availability_map[week][r].append(name)

    def generate_draft(self):
        self.initial_roster = {week: {} for week in self.week_columns}
        burnout = {name: 0 for name in self.df['Name']}
        last_week_played = {name: -1 for name in self.df['Name']}
        
        for w_idx, week in enumerate(self.week_columns):
            assigned_this_week = [] 
            sorted_roles = sorted(ROLES_ORDER, key=lambda r: len(self.availability_map[week][r]))
            
            for role in sorted_roles:
                candidates = [p for p in self.availability_map[week][role] if p not in assigned_this_week]
                
                if candidates:
                    random.shuffle(candidates)
                    if "Cleanup" in role:
                        winner = candidates[0]
                        self.initial_roster[week][role] = winner
                        assigned_this_week.append(winner)
                    else:
                        candidates.sort(key=lambda p: (burnout.get(p, 0) * 10) + (50 if last_week_played.get(p) == (w_idx - 1) else 0))
                        winner = candidates[0]
                        self.initial_roster[week][role] = winner
                        assigned_this_week.append(winner)
                        burnout[winner] = burnout.get(winner, 0) + 1
                        last_week_played[winner] = w_idx
                else:
                    self.initial_roster[week][role] = ""

# ==========================================
# GUI CLASS
# ==========================================
class RosterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Auto-Roster Pro (Visual Dashboard)")
        self.geometry("1600x900")
        self.engine = RosterEngine()
        self.combos = {} 
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.map("TCombobox", fieldbackground=[("readonly", "white")])
        
        self._build_ui()

    def _build_ui(self):
        top_frame = tk.Frame(self, pady=5, bg="#ddd")
        top_frame.pack(fill=tk.X)
        
        tk.Button(top_frame, text="1. Load Excel", command=self.load_file, bg="white").pack(side=tk.LEFT, padx=10)
        self.lbl_status = tk.Label(top_frame, text="No file loaded", bg="#ddd", fg="red")
        self.lbl_status.pack(side=tk.LEFT)
        tk.Button(top_frame, text="3. Export", command=self.export_file, bg="#4CAF50", fg="white").pack(side=tk.RIGHT, padx=10)
        tk.Button(top_frame, text="2. Clear Grid", command=self.clear_grid, bg="#ff9999", fg="black").pack(side=tk.RIGHT, padx=10)

        paned = tk.PanedWindow(self, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # TOP: INPUT GRID
        self.roster_frame = tk.Frame(paned)
        paned.add(self.roster_frame, height=300)
        self.canvas_r = tk.Canvas(self.roster_frame)
        self.scroll_y_r = tk.Scrollbar(self.roster_frame, orient="vertical", command=self.canvas_r.yview)
        self.scroll_x_r = tk.Scrollbar(self.roster_frame, orient="horizontal", command=self.canvas_r.xview)
        self.grid_container = tk.Frame(self.canvas_r)
        self.grid_container.bind("<Configure>", lambda e: self.canvas_r.configure(scrollregion=self.canvas_r.bbox("all")))
        self.canvas_r.create_window((0, 0), window=self.grid_container, anchor="nw")
        self.canvas_r.configure(yscrollcommand=self.scroll_y_r.set, xscrollcommand=self.scroll_x_r.set)
        self.scroll_x_r.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas_r.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll_y_r.pack(side=tk.RIGHT, fill=tk.Y)

        # BOTTOM: DASHBOARD
        self.dash_frame = tk.Frame(paned, bg="white")
        paned.add(self.dash_frame, stretch="always")
        
        # Legend
        legend_frame = tk.Frame(self.dash_frame, bg="white")
        legend_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(legend_frame, text="LEGEND:", bg="white", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
        for cat, data in CATEGORY_CONFIG.items():
             tk.Label(legend_frame, text=f" ■ {cat} ", fg=data["color"], bg="white").pack(side=tk.LEFT)

        self.canvas_d = tk.Canvas(self.dash_frame, bg="white")
        self.scroll_y_d = tk.Scrollbar(self.dash_frame, orient="vertical", command=self.canvas_d.yview)
        self.scroll_x_d = tk.Scrollbar(self.dash_frame, orient="horizontal", command=self.canvas_d.xview)
        self.dash_container = tk.Frame(self.canvas_d, bg="white")
        self.dash_container.bind("<Configure>", lambda e: self.canvas_d.configure(scrollregion=self.canvas_d.bbox("all")))
        self.canvas_d.create_window((0, 0), window=self.dash_container, anchor="nw")
        self.canvas_d.configure(yscrollcommand=self.scroll_y_d.set, xscrollcommand=self.scroll_x_d.set)
        self.scroll_x_d.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas_d.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll_y_d.pack(side=tk.RIGHT, fill=tk.Y)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path: return
        success, msg = self.engine.load_file(path)
        if success:
            self.lbl_status.config(text=f"Loaded: {os.path.basename(path)}", fg="green")
            self.engine.generate_draft()
            self.render_roster_grid()
            self.update_dashboard()
        else:
            messagebox.showerror("Error", msg)

    def clear_grid(self):
        if not self.combos: return
        if messagebox.askyesno("Confirm", "Clear all selections?"):
            for cb in self.combos.values(): cb.set("")
            self.validate_all()
            self.update_dashboard()

    def render_roster_grid(self):
        for w in self.grid_container.winfo_children(): w.destroy()
        self.combos = {}

        tk.Label(self.grid_container, text="Week", font=("Arial", 9, "bold"), width=15, relief="solid").grid(row=0, column=0, padx=1)
        for i, role in enumerate(ROLES_ORDER):
            cat_data = ROLE_TO_CAT_MAP.get(role, {"color": "#ddd"})
            tk.Label(self.grid_container, text=role, font=("Arial", 9, "bold"), fg=cat_data["color"], width=12, relief="solid").grid(row=0, column=i+1, padx=1)

        for r, week in enumerate(self.engine.week_columns):
            row = r + 1
            tk.Label(self.grid_container, text=week, font=("Arial", 8), width=15, anchor="w").grid(row=row, column=0, padx=1, pady=5)
            for c, role in enumerate(ROLES_ORDER):
                cb = ttk.Combobox(self.grid_container, state="readonly", width=11)
                draft = self.engine.initial_roster[week].get(role, "")
                if draft: cb.set(draft)
                cb.bind('<Button-1>', functools.partial(self.update_dropdown_options, week=week, role=role, widget=cb))
                cb.bind('<<ComboboxSelected>>', self.on_selection_change)
                cb.grid(row=row, column=c+1, padx=2)
                self.combos[(week, role)] = cb
        
        self.validate_all()

    def update_dropdown_options(self, event, week, role, widget):
        capable = self.engine.availability_map[week][role]
        busy = []
        for r in ROLES_ORDER:
            if r == role: continue 
            val = self.combos[(week, r)].get()
            if val: busy.append(val)
        
        filtered = [p for p in capable if p not in busy]
        if "Cleanup" not in role: filtered.sort()
        filtered.insert(0, "")
        widget['values'] = filtered

    def on_selection_change(self, event):
        self.validate_all()
        self.update_dashboard()

    def validate_all(self):
        for week in self.engine.week_columns:
            seen = {}
            dupes = []
            for role in ROLES_ORDER:
                val = self.combos[(week, role)].get()
                if val:
                    if val in seen: dupes.append(val)
                    seen[val] = True
            for role in ROLES_ORDER:
                w = self.combos[(week, role)]
                if w.get() and w.get() in dupes: w.config(foreground="red")
                else: w.config(foreground="black")

    # ==========================================
    # VISUAL DASHBOARD (THE BIG UPDATE)
    # ==========================================
    def update_dashboard(self):
        for w in self.dash_container.winfo_children(): w.destroy()

        # 1. Map who is serving where for each week
        # Structure: assigned_map[week][name] = "RoleName"
        assigned_map = {week: {} for week in self.engine.week_columns}
        serve_counts = {name: 0 for name in self.engine.all_members}
        
        for week in self.engine.week_columns:
            for role in ROLES_ORDER:
                if (week, role) in self.combos:
                    name = self.combos[(week, role)].get()
                    if name:
                        assigned_map[week][name] = role
                        if name in serve_counts: serve_counts[name] += 1

        # 2. Build Columns based on CATEGORY_CONFIG
        col_idx = 0
        
        for cat_name, cat_data in CATEGORY_CONFIG.items():
            # A. Category Header
            tk.Label(self.dash_container, text=cat_name, bg=cat_data["color"], fg=cat_data["text_col"], 
                     font=("Arial", 10, "bold"), relief="flat").grid(
                         row=0, column=col_idx, columnspan=len(cat_data["roles"]), sticky="ew", padx=1, pady=(0, 0))
            
            # B. Role Sub-Headers & Member Lists
            start_col = col_idx
            
            for role in cat_data["roles"]:
                # Sub Header
                tk.Label(self.dash_container, text=role, bg="#eee", font=("Arial", 8, "bold"), relief="solid").grid(
                    row=1, column=col_idx, sticky="ew", padx=0)

                # Get Members capable of this role
                members = []
                
                if cat_name == "LG": # Cleanup is special
                    for opt in CLEANUP_OPTIONS:
                        members.append({"name": opt, "avail": "XXXX", "count": 0, "obj": None})
                else:
                    # Find relevant members
                    for name, data in self.engine.all_members.items():
                        # Check capability
                        has_cap = False
                        if "Usher" in role: has_cap = "Usher" in data["Roles"]
                        elif role in data["Roles"]: has_cap = True
                        
                        if has_cap:
                            members.append({
                                "name": name,
                                "avail": data["AvailString"],
                                "count": serve_counts.get(name, 0),
                                "obj": data
                            })
                    
                    # Sort: Count Desc, then Name
                    members.sort(key=lambda x: (-x["count"], x["name"]))

                # Render Members
                row_idx = 2
                for m in members:
                    self._render_member_cell(m, row_idx, col_idx, assigned_map)
                    row_idx += 1
                
                col_idx += 1
            
            # Divider Column (Gap)
            tk.Frame(self.dash_container, width=10, bg="white").grid(row=0, column=col_idx)
            col_idx += 1

    def _render_member_cell(self, m, row, col, assigned_map):
        container = tk.Frame(self.dash_container, bg="white", borderwidth=1, relief="solid")
        container.grid(row=row, column=col, sticky="ew", padx=0, pady=0)
        container.columnconfigure(1, weight=1) # Name expands

        # 1. Color Logic based on count
        bg_col = "white"
        if m["count"] >= 3: bg_col = "#ffcccc" # Red Burnout
        elif m["count"] >= 1: bg_col = "#ffeeb0" # Orange Active
        
        # 2. Name Label
        name_lbl = tk.Label(container, text=m["name"], bg=bg_col, font=("Arial", 8), anchor="w", width=10)
        name_lbl.pack(side=tk.LEFT, fill=tk.Y)

        # 3. Status Indicators (The O/X circles)
        if m["name"] in CLEANUP_OPTIONS:
            tk.Label(container, text="----", bg=bg_col, font=("Arial", 8)).pack(side=tk.LEFT)
        else:
            status_frame = tk.Frame(container, bg=bg_col)
            status_frame.pack(side=tk.LEFT, padx=2)
            
            for i, char in enumerate(m["avail"]):
                color = "black" # Default Available
                text = "O"
                week_key = self.engine.week_columns[i]
                
                if char == "X":
                    color = "#ccc" # Unavailable
                    text = "X"
                else:
                    # Check if serving somewhere
                    if m["name"] in assigned_map[week_key]:
                        role_assigned = assigned_map[week_key][m["name"]]
                        # Get Color of that role
                        cat_info = ROLE_TO_CAT_MAP.get(role_assigned)
                        if cat_info:
                            color = cat_info["color"]
                
                tk.Label(status_frame, text=text, fg=color, bg=bg_col, font=("Arial", 8, "bold"), width=1).pack(side=tk.LEFT)

        # 4. Count Label
        count_lbl = tk.Label(container, text=f"({m['count']})", bg=bg_col, font=("Arial", 8), width=3)
        count_lbl.pack(side=tk.RIGHT)

    def export_file(self):
        if not self.engine.week_columns: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not path: return

        data = []
        for week in self.engine.week_columns:
            row = {"Week": week}
            for role in ROLES_ORDER:
                row[role] = self.combos[(week, role)].get()
            
            has_drums = row["Drum/Cajon"] != ""
            has_keys = row["Piano"] != ""
            has_bass = row["Bass"] != ""
            mode = "INCOMPLETE"
            if has_bass: mode = "FULL BAND"
            elif has_drums and has_keys: mode = "ACOUSTIC SET"
            
            final_row = {"Week": week, "Band Mode": mode}
            final_row.update({k: v for k, v in row.items() if k != "Week"})
            data.append(final_row)

        pd.DataFrame(data).to_excel(path, index=False)
        messagebox.showinfo("Done", "Exported!")

if __name__ == "__main__":
    app = RosterApp()
    app.mainloop()