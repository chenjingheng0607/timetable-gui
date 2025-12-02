import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import random
import functools

# Image Generation Library
try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ==========================================
# CONFIGURATION
# ==========================================
ROLES_ORDER = [
    "Lead", "Vocal", "Piano", "Drum/Cajon", "Bass", "Guitar", 
    "PPT", "Sound", "Lighting/OBS", 
    "MC", 
    "Usher 1", "Usher 2", "Usher 3", 
    "Cleanup 1", "Cleanup 2"
]

CLEANUP_OPTIONS = ["LHW", "UF", "LB", "YGSS", "SJS", "PK"]

# MAPPING
INSTRUMENT_MAP = {
    "WL": "Lead", "V": "Vocal", "P": "Piano", "G": "Guitar", 
    "B": "Bass", "D": "Drum/Cajon", "PPT": "PPT", 
    "S": "Sound", "SOUND": "Sound", 
    "OBS": "Lighting/OBS", "LIGHT": "Lighting/OBS", "L": "Lighting/OBS",
    "MC": "MC", "USHER": "Usher"
}

# COLORS & CATEGORIES
CATEGORY_CONFIG = {
    "Praise & Worship": {
        "roles": ["Lead", "Vocal", "Piano", "Drum/Cajon", "Bass", "Guitar"],
        "color": "#c00000", "text_col": "white"
    },
    "FPH": {
        "roles": ["PPT", "Sound", "Lighting/OBS"], 
        "color": "#0070c0", "text_col": "white"
    },
    "MC": {
        "roles": ["MC"],
        "color": "#7030a0", "text_col": "white"
    },
    "Usher": {
        "roles": ["Usher 1", "Usher 2", "Usher 3"],
        "color": "#ffc000", "text_col": "black"
    },
    "LG": {
        "roles": ["Cleanup 1", "Cleanup 2"],
        "color": "#00b050", "text_col": "white"
    }
}

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

            if not self.initial_roster[week].get("Piano"):
                if self.initial_roster[week].get("Bass"):
                    bassist = self.initial_roster[week]["Bass"]
                    self.initial_roster[week]["Bass"] = ""
                    if bassist in burnout: burnout[bassist] -= 1

# ==========================================
# GUI CLASS
# ==========================================
class RosterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Auto-Roster Pro (Final Fix)")
        self.geometry("1600x900")
        self.engine = RosterEngine()
        self.combos = {} 
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.map("TCombobox", fieldbackground=[("readonly", "white")])
        
        if not HAS_PIL:
            messagebox.showwarning("Missing Library", "Pillow not found. Image export disabled.\nRun: pip install Pillow")

        self._build_ui()

    def _build_ui(self):
        # TOP BAR
        top_frame = tk.Frame(self, pady=5, bg="#ddd")
        top_frame.pack(fill=tk.X)
        tk.Button(top_frame, text="1. Load Excel", command=self.load_file, bg="white").pack(side=tk.LEFT, padx=5)
        self.lbl_status = tk.Label(top_frame, text="No file loaded", bg="#ddd", fg="red")
        self.lbl_status.pack(side=tk.LEFT)
        
        tk.Button(top_frame, text="4. Export Image", command=self.export_image_cmd, bg="#0078d7", fg="white", font=("Arial", 9, "bold")).pack(side=tk.RIGHT, padx=5)
        tk.Button(top_frame, text="3. Export Excel", command=self.export_excel, bg="#4CAF50", fg="white").pack(side=tk.RIGHT, padx=5)
        tk.Button(top_frame, text="2. Clear Grid", command=self.clear_grid, bg="#ff9999", fg="black").pack(side=tk.RIGHT, padx=5)

        paned = tk.PanedWindow(self, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

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

        self.dash_frame = tk.Frame(paned, bg="white")
        paned.add(self.dash_frame, stretch="always")
        
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
            self.on_selection_change(None)

    def render_roster_grid(self):
        for w in self.grid_container.winfo_children(): w.destroy()
        self.combos = {}

        total_rows = len(self.engine.week_columns) + 1

        tk.Label(self.grid_container, text="Week", font=("Arial", 9, "bold"), width=15, relief="solid").grid(row=0, column=0, padx=1, sticky="ns")
        sep_main = tk.Frame(self.grid_container, width=3, bg="#888")
        sep_main.grid(row=0, column=1, rowspan=total_rows, sticky="ns", padx=2)

        current_col = 2
        prev_cat = None

        for role in ROLES_ORDER:
            this_cat = ROLE_TO_CAT_MAP[role]["cat"]
            
            if prev_cat and this_cat != prev_cat:
                sep = tk.Frame(self.grid_container, width=3, bg="#888")
                sep.grid(row=0, column=current_col, rowspan=total_rows, sticky="ns", padx=2)
                current_col += 1
            
            cat_data = ROLE_TO_CAT_MAP.get(role, {"color": "#ddd"})
            tk.Label(self.grid_container, text=role, font=("Arial", 9, "bold"), fg=cat_data["color"], width=12, relief="solid").grid(row=0, column=current_col, padx=1)
            
            for r, week in enumerate(self.engine.week_columns):
                row_idx = r + 1
                if role == ROLES_ORDER[0]: 
                    tk.Label(self.grid_container, text=week, font=("Arial", 8), width=15, anchor="w").grid(row=row_idx, column=0, padx=1, pady=5)

                cb = ttk.Combobox(self.grid_container, state="readonly", width=11)
                draft = self.engine.initial_roster[week].get(role, "")
                if draft: cb.set(draft)
                cb.bind('<Button-1>', functools.partial(self.update_dropdown_options, week=week, role=role, widget=cb))
                cb.bind('<<ComboboxSelected>>', self.on_selection_change)
                cb.grid(row=row_idx, column=current_col, padx=2)
                self.combos[(week, role)] = cb

            prev_cat = this_cat
            current_col += 1
        
        self.on_selection_change(None) 

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
        self.update_locks()
        self.validate_all()
        self.update_dashboard()

    def update_locks(self):
        for week in self.engine.week_columns:
            if (week, "Piano") not in self.combos: continue
            piano_val = self.combos[(week, "Piano")].get()
            bass_combo = self.combos[(week, "Bass")]
            if not piano_val:
                bass_combo.set("")
                bass_combo.state(["disabled"])
            else:
                bass_combo.state(["!disabled"])
                bass_combo.config(state="readonly")

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
            
            if not self.combos[(week, "Piano")].get() and self.combos[(week, "Bass")].get():
                self.combos[(week, "Bass")].config(foreground="red")

    def update_dashboard(self):
        for w in self.dash_container.winfo_children(): w.destroy()

        assigned_map = {week: {} for week in self.engine.week_columns}
        serve_counts = {name: 0 for name in self.engine.all_members}
        cleanup_counts = {opt: 0 for opt in CLEANUP_OPTIONS}
        
        member_active_roles = {name: set() for name in self.engine.all_members}
        cleanup_active_roles = {opt: set() for opt in CLEANUP_OPTIONS}

        for week in self.engine.week_columns:
            for role in ROLES_ORDER:
                if (week, role) in self.combos:
                    name = self.combos[(week, role)].get()
                    if name:
                        assigned_map[week][name] = role
                        if "Cleanup" in role:
                            if name in cleanup_counts: 
                                cleanup_counts[name] += 1
                                cleanup_active_roles[name].add(role)
                        else:
                            if name in serve_counts: 
                                serve_counts[name] += 1
                                member_active_roles[name].add(role)

        col_idx = 0
        for cat_name, cat_data in CATEGORY_CONFIG.items():
            tk.Label(self.dash_container, text=cat_name, bg=cat_data["color"], fg=cat_data["text_col"], 
                     font=("Arial", 10, "bold"), relief="flat").grid(
                         row=0, column=col_idx, columnspan=len(cat_data["roles"]), sticky="ew", padx=1)
            
            for role in cat_data["roles"]:
                tk.Label(self.dash_container, text=role, bg="#eee", font=("Arial", 8, "bold"), relief="solid").grid(
                    row=1, column=col_idx, sticky="ew", padx=0)

                members = []
                if cat_name == "LG": 
                    for opt in CLEANUP_OPTIONS:
                        is_active = role in cleanup_active_roles.get(opt, set())
                        s_val = 2
                        if is_active:
                            if cleanup_counts[opt] >= 3: s_val = 0
                            elif cleanup_counts[opt] >= 1: s_val = 1
                        members.append({"name": opt, "avail": "XXXX", "count": cleanup_counts[opt], "is_active": is_active, "sort_val": s_val})
                else:
                    for name, data in self.engine.all_members.items():
                        has_cap = False
                        if "Usher" in role: has_cap = "Usher" in data["Roles"]
                        elif role in data["Roles"]: has_cap = True
                        
                        if has_cap:
                            is_active = role in member_active_roles.get(name, set())
                            cnt = serve_counts.get(name, 0)
                            s_val = 2
                            if is_active:
                                if cnt >= 3: s_val = 0
                                elif cnt >= 1: s_val = 1
                            
                            members.append({
                                "name": name,
                                "avail": data["AvailString"],
                                "count": cnt,
                                "is_active": is_active,
                                "sort_val": s_val
                            })
                
                members.sort(key=lambda x: (x["sort_val"], -x["count"], x["name"]))

                row_idx = 2
                for m in members:
                    self._render_member_cell(m, row_idx, col_idx, assigned_map)
                    row_idx += 1
                col_idx += 1
            
            tk.Frame(self.dash_container, width=15, bg="white").grid(row=0, column=col_idx)
            col_idx += 1

    def _render_member_cell(self, m, row, col, assigned_map):
        container = tk.Frame(self.dash_container, bg="white", borderwidth=1, relief="solid")
        container.grid(row=row, column=col, sticky="ew")
        container.columnconfigure(1, weight=1)

        bg_col = "white"
        if m["is_active"]:
            if m["count"] >= 3: bg_col = "#ffcccc"
            elif m["count"] >= 1: bg_col = "#ffeeb0"
        
        name_lbl = tk.Label(container, text=m["name"], bg=bg_col, font=("Arial", 8), anchor="w", width=10)
        name_lbl.pack(side=tk.LEFT, fill=tk.Y)

        if m["name"] in CLEANUP_OPTIONS:
            tk.Label(container, text="----", bg=bg_col, font=("Arial", 8)).pack(side=tk.LEFT)
        else:
            status_frame = tk.Frame(container, bg=bg_col)
            status_frame.pack(side=tk.LEFT, padx=2)
            for i, char in enumerate(m["avail"]):
                color = "black"
                text = "O"
                week_key = self.engine.week_columns[i]
                if char == "X":
                    color = "#ccc"; text = "X"
                else:
                    if m["name"] in assigned_map[week_key]:
                        role_assigned = assigned_map[week_key][m["name"]]
                        cat_info = ROLE_TO_CAT_MAP.get(role_assigned)
                        if cat_info: color = cat_info["color"]
                
                tk.Label(status_frame, text=text, fg=color, bg=bg_col, font=("Arial", 8, "bold"), width=1).pack(side=tk.LEFT)

        tk.Label(container, text=f"({m['count']})", bg=bg_col, font=("Arial", 8), width=3).pack(side=tk.RIGHT)

    def export_excel(self):
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
        messagebox.showinfo("Done", "Excel Exported!")

    # ==========================================
    # IMAGE EXPORT
    # ==========================================
    def export_image_cmd(self):
        if not HAS_PIL:
            messagebox.showerror("Error", "Pillow not installed.")
            return
        if not self.engine.week_columns: return
        
        path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG Image", "*.png")])
        if not path: return

        assigned_map = {week: {} for week in self.engine.week_columns}
        serve_counts = {name: 0 for name in self.engine.all_members}
        cleanup_counts = {opt: 0 for opt in CLEANUP_OPTIONS}
        
        member_active_roles = {name: set() for name in self.engine.all_members}
        cleanup_active_roles = {opt: set() for opt in CLEANUP_OPTIONS}

        roster_data = {}
        for week in self.engine.week_columns:
            roster_data[week] = {}
            for role in ROLES_ORDER:
                val = self.combos[(week, role)].get()
                roster_data[week][role] = val
                if val:
                    assigned_map[week][val] = role
                    if "Cleanup" in role:
                        if val in cleanup_counts: 
                            cleanup_counts[val] += 1
                            cleanup_active_roles[val].add(role)
                    else:
                        if val in serve_counts: 
                            serve_counts[val] += 1
                            member_active_roles[val].add(role)

        # --- DRAWING ---
        COL_W = 160
        ROW_H = 30
        MARGIN = 20
        FONT_SIZE = 12
        CAT_SPACER = 10
        
        roster_w = COL_W * (len(ROLES_ORDER) + 1)
        dash_w = 0
        for cat, data in CATEGORY_CONFIG.items():
            dash_w += (len(data["roles"]) * COL_W) + CAT_SPACER
        dash_w -= CAT_SPACER 
        
        img_w = max(roster_w, dash_w) + (MARGIN * 2)
        
        max_mem_rows = 0
        for cat, data in CATEGORY_CONFIG.items():
            for role in data["roles"]:
                c = 0
                if cat == "LG": c = len(CLEANUP_OPTIONS)
                else:
                    for _, d in self.engine.all_members.items():
                        if "Usher" in role and "Usher" in d["Roles"]: c+=1
                        elif role in d["Roles"]: c+=1
                max_mem_rows = max(max_mem_rows, c)
        
        roster_h = ROW_H * (len(self.engine.week_columns) + 2) 
        dash_h = ROW_H * (max_mem_rows + 3)
        img_h = roster_h + 60 + dash_h + (MARGIN * 2)
        
        img = Image.new("RGB", (img_w, img_h), "white")
        draw = ImageDraw.Draw(img)
        
        try:
            font = ImageFont.truetype("arial.ttf", FONT_SIZE)
            font_bold = ImageFont.truetype("arialbd.ttf", FONT_SIZE)
        except:
            font = ImageFont.load_default()
            font_bold = font

        # DRAW ROSTER
        roster_start_x = (img_w - roster_w) // 2
        y = MARGIN
        x = roster_start_x
        
        x_curs = x + COL_W
        for cat_name, cat_data in CATEGORY_CONFIG.items():
            w = len(cat_data["roles"]) * COL_W
            draw.rectangle([x_curs, y, x_curs+w, y+ROW_H], fill="white", outline="black")
            text_len = draw.textlength(cat_name, font=font_bold)
            draw.text((x_curs + (w-text_len)/2, y+5), cat_name, fill=cat_data["color"], font=font_bold)
            x_curs += w
        y += ROW_H
        
        draw.rectangle([x, y, x+COL_W, y+ROW_H], outline="black") 
        x_curs = x + COL_W
        for role in ROLES_ORDER:
            draw.rectangle([x_curs, y, x_curs+COL_W, y+ROW_H], outline="black", fill="#f0f0f0")
            draw.text((x_curs+5, y+5), role, fill="black", font=font_bold)
            x_curs += COL_W
        y += ROW_H
        
        for week in self.engine.week_columns:
            draw.rectangle([x, y, x+COL_W, y+ROW_H], outline="black")
            draw.text((x+5, y+5), week, fill="black", font=font)
            x_curs = x + COL_W
            for role in ROLES_ORDER:
                val = roster_data[week][role]
                draw.rectangle([x_curs, y, x_curs+COL_W, y+ROW_H], outline="black")
                if val: draw.text((x_curs+5, y+5), val, fill="black", font=font)
                x_curs += COL_W
            y += ROW_H

        # DRAW DASHBOARD
        y += 60
        x = (img_w - dash_w) // 2
        x_curs = x
        
        for cat_name, cat_data in CATEGORY_CONFIG.items():
            width = len(cat_data["roles"]) * COL_W
            draw.rectangle([x_curs, y, x_curs+width, y+ROW_H], fill=cat_data["color"], outline="black")
            draw.text((x_curs+5, y+5), cat_name, fill=cat_data["text_col"], font=font_bold)
            
            role_x = x_curs
            for role in cat_data["roles"]:
                draw.rectangle([role_x, y+ROW_H, role_x+COL_W, y+(ROW_H*2)], fill="#eee", outline="black")
                draw.text((role_x+5, y+ROW_H+5), role, fill="black", font=font_bold)
                
                members = []
                if cat_name == "LG":
                    for opt in CLEANUP_OPTIONS:
                        is_active = role in cleanup_active_roles.get(opt, set())
                        s_val = 2
                        if is_active:
                            if cleanup_counts[opt] >= 3: s_val = 0
                            elif cleanup_counts[opt] >= 1: s_val = 1
                        members.append({"name": opt, "avail": "XXXX", "count": cleanup_counts[opt], "is_active": is_active, "sort_val": s_val})
                else:
                    for name, d in self.engine.all_members.items():
                        has_cap = False
                        if "Usher" in role: has_cap = "Usher" in d["Roles"]
                        elif role in d["Roles"]: has_cap = True
                        if has_cap:
                            is_active = role in member_active_roles.get(name, set())
                            cnt = serve_counts.get(name, 0)
                            s_val = 2
                            if is_active:
                                if cnt >= 3: s_val = 0
                                elif cnt >= 1: s_val = 1
                            members.append({
                                "name": name, 
                                "avail": d["AvailString"], 
                                "count": cnt,
                                "is_active": is_active,
                                "sort_val": s_val
                            })
                
                members.sort(key=lambda x: (x["sort_val"], -x["count"], x["name"]))
                
                mem_y = y + (ROW_H*2)
                for m in members:
                    bg = "white"
                    if m["is_active"]:
                        if m["count"] >= 3: bg = "#ffcccc"
                        elif m["count"] >= 1: bg = "#ffeeb0"
                    
                    draw.rectangle([role_x, mem_y, role_x+COL_W, mem_y+ROW_H], fill=bg, outline="black")
                    
                    draw.text((role_x+5, mem_y+5), m["name"], fill="black", font=font)
                    
                    # RIGHT ALIGN COUNT
                    count_text = f"({m['count']})"
                    c_len = draw.textlength(count_text, font=font)
                    count_x = role_x + COL_W - c_len - 5
                    draw.text((count_x, mem_y+5), count_text, fill="black", font=font)
                    
                    # DOTS LEFT OF COUNT
                    if cat_name != "LG":
                        dot_block_w = 45
                        dot_start_x = count_x - dot_block_w
                        for i, char in enumerate(m["avail"]):
                            color = "black"
                            if char == "X": color = "#ccc"
                            else:
                                wk = self.engine.week_columns[i]
                                if m["name"] in assigned_map[wk]:
                                    assigned_role = assigned_map[wk][m["name"]]
                                    c_info = ROLE_TO_CAT_MAP.get(assigned_role)
                                    if c_info: color = c_info["color"]
                            draw.text((dot_start_x + (i*11), mem_y+5), char, fill=color, font=font_bold)
                    
                    mem_y += ROW_H
                role_x += COL_W
            x_curs += width + CAT_SPACER
        
        img.save(path)
        messagebox.showinfo("Success", "Image Exported!")

if __name__ == "__main__":
    app = RosterApp()
    app.mainloop()