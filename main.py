import sys
import os
import random
import functools
import traceback
import pandas as pd
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QLabel, QScrollArea, 
                               QFrame, QFileDialog, QMessageBox, QComboBox, 
                               QGridLayout, QSizePolicy, QSplitter)
from PySide6.QtCore import Qt, Signal, QTimer
from PySide6.QtGui import QColor, QFont, QIcon

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

# BASIC DATA STRUCTURE (Colors handled by Theme Manager)
CATEGORY_DATA = {
    "Praise & Worship": ["Lead", "Vocal", "Piano", "Drum/Cajon", "Bass", "Guitar"],
    "FPH": ["PPT", "Sound", "Lighting/OBS"],
    "MC": ["MC"],
    "Usher": ["Usher 1", "Usher 2", "Usher 3"],
    "LG": ["Cleanup 1", "Cleanup 2"]
}

# THEMES
THEMES = {
    "Dark": {
        "bg_main": "#1e1e1e", "bg_sec": "#252526", "fg_pri": "#ffffff", "fg_sec": "#cccccc",
        "input_bg": "#333333", "input_border": "#3e3e42", "input_sel": "#264f78",
        "btn_bg": "#0e639c", "btn_fg": "#ffffff",
        "cats": {
            "Praise & Worship": "#ff5252",
            "FPH": "#42a5f5",
            "MC": "#ab47bc",
            "Usher": "#ffa726",
            "LG": "#66bb6a"
        },
        "dash_bg_warn": "#4a0000", 
        "dash_bg_notice": "#4a3b00",
        "dash_text_avail": "white",
        "dash_text_unavail": "#666666",
        "active_cell_text": "#ffffff"
    },
    "Light": {
        "bg_main": "#f5f5f5", "bg_sec": "#ffffff", "fg_pri": "#000000", "fg_sec": "#333333",
        "input_bg": "#ffffff", "input_border": "#bdc3c7", "input_sel": "#0078d7",
        "btn_bg": "#0078d7", "btn_fg": "#ffffff",
        "cats": {
            "Praise & Worship": "#C62828",
            "FPH": "#1565C0",
            "MC": "#6A1B9A",
            "Usher": "#EF6C00",
            "LG": "#2E7D32"
        },
        "dash_bg_warn": "#FFEBEE",
        "dash_bg_notice": "#FFFDE7",
        "dash_text_avail": "black",
        "dash_text_unavail": "#999999",
        "active_cell_text": "#000000"
    }
}

# Global Map - Updated by App.apply_theme
ROLE_TO_CAT_MAP = {}
def build_role_map(theme_cats):
    rm = {}
    for cat, roles in CATEGORY_DATA.items():
        for r in roles:
            rm[r] = {"cat": cat, "color": theme_cats.get(cat, "#888")}
    return rm

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
            # First pass: find header
            df_raw = pd.read_excel(filepath, header=None, keep_default_na=False)
            header_row = -1
            for i, row in df_raw.iterrows():
                row_str = [str(x).strip() for x in row.values]
                if "Name" in row_str:
                    header_row = i
                    break
            
            if header_row == -1: return False, "Could not find 'Name' column."
            
            # Second pass: read data
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
        
        # Optimize column finding
        def check_col(name, keys):
            name_upper = str(name).upper()
            if any(k in name_upper for k in keys): return True
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
                        # Direct assignment optimization
                        if "Usher" in r:
                            if "Usher" in caps: self.availability_map[week][r].append(name)
                        elif r in caps:
                            self.availability_map[week][r].append(name)

    def generate_draft(self):
        self.initial_roster = {week: {} for week in self.week_columns}
        burnout = {name: 0 for name in self.df['Name']}
        last_week_played = {name: -1 for name in self.df['Name']}
        
        for w_idx, week in enumerate(self.week_columns):
            assigned_this_week = set() # Use set for O(1) lookups
            sorted_roles = sorted(ROLES_ORDER, key=lambda r: len(self.availability_map[week][r]))
            
            for role in sorted_roles:
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

            # Piano/Bass Logic
            if not self.initial_roster[week].get("Piano"):
                if self.initial_roster[week].get("Bass"):
                    bassist = self.initial_roster[week]["Bass"]
                    self.initial_roster[week]["Bass"] = ""
                    if bassist in burnout: burnout[bassist] -= 1

# ==========================================
# GUI CLASS
# ==========================================

class EnhancedComboBox(QComboBox):
    def __init__(self, callback, week, role, parent=None):
        super().__init__(parent)
        self.callback = callback
        self.week = week
        self.role = role
        
    def showPopup(self):
        self.callback(self.week, self.role, self)
        super().showPopup()

class RosterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Auto-Roster Pro")
        self.resize(1600, 900)
        
        self.engine = RosterEngine()
        self.combos = {} 
        self.current_theme = "Dark" 
        
        # Debounce Timer for Dashboard Updates to improve performance
        self.update_timer = QTimer()
        self.update_timer.setSingleShot(True)
        self.update_timer.setInterval(80) # 80ms delay
        self.update_timer.timeout.connect(self._perform_dashboard_update)

        # Style
        QApplication.setStyle("Fusion")
        
        if not HAS_PIL:
            QMessageBox.warning(self, "Missing Library", "Pillow not found. Image export disabled.\nRun: pip install Pillow")

        self.apply_theme(self.current_theme)

    def apply_theme(self, theme_name):
        # Save current state before rebuilding
        saved_state = {}
        if hasattr(self, 'combos') and self.combos:
            for key, cb in self.combos.items():
                saved_state[key] = cb.currentText()
        
        saved_status = "No file loaded"
        if hasattr(self, 'lbl_status'):
            saved_status = self.lbl_status.text()

        self.current_theme = theme_name
        t = THEMES[theme_name]
        
        global ROLE_TO_CAT_MAP
        ROLE_TO_CAT_MAP = build_role_map(t["cats"])

        # GENERATE STYLESHEET
        css = f"""
            QWidget {{ 
                background-color: {t['bg_main']}; 
                color: {t['fg_pri']}; 
                font-family: "Segoe UI", Arial, sans-serif;
            }}
            QFrame {{ 
                background-color: {t['bg_main']}; 
                border: none; 
            }}
            QScrollArea {{ 
                background-color: {t['bg_main']}; 
                border: none; 
            }}
            QLineEdit, QComboBox, QAbstractItemView {{ 
                background-color: {t['input_bg']}; 
                color: {t['fg_pri']}; 
                border: 1px solid {t['input_border']};
                selection-background-color: {t['input_sel']};
                selection-color: white;
            }}
            QComboBox::drop-down {{
                border: none;
            }}
            QComboBox:disabled {{
                background-color: {t['bg_sec']};
                color: {t['fg_sec']};
            }}
            QLabel {{
                color: {t['fg_pri']};
            }}
            QPushButton {{
                background-color: {t['input_bg']};
                border: 1px solid {t['input_border']};
                padding: 4px;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background-color: {t['input_sel']};
                color: white;
            }}
            QMessageBox QPushButton {{
                min-width: 100px; 
                padding: 6px;
            }}
            QSplitter::handle {{
                background-color: {t['input_border']};
            }}
            /* Header / Legend frames */
            .QFrame_Legend, .QFrame_Top {{
                border-bottom: 1px solid {t['input_border']};
            }}
        """
        self.setStyleSheet(css)
        self._build_ui() # Rebuild to apply new styles/colors to dynamic elements
        
        # Restore Status
        if hasattr(self, 'lbl_status'):
            self.lbl_status.setText(saved_status)
            if "Loaded:" in saved_status:
                self.lbl_status.setStyleSheet("color: #4CAF50; margin-left: 10px;")
            else:
                self.lbl_status.setStyleSheet("color: red; margin-left: 10px;")

        if self.engine.week_columns:
            self.render_roster_grid()
            
            # Restore state
            for key, val in saved_state.items():
                if key in self.combos:
                    cb = self.combos[key]
                    cb.blockSignals(True)
                    # Use setItemText or check if in model? 
                    # Simpler to just add it if missing (though it should be valid) and set it.
                    if val and cb.findText(val) == -1:
                        cb.addItem(val)
                    cb.setCurrentText(val)
                    cb.blockSignals(False)

            self.trigger_dashboard_update()

    def toggle_theme(self):
        new = "Light" if self.current_theme == "Dark" else "Dark"
        self.apply_theme(new)

    def _build_ui(self):
        # Clear central widget
        if self.centralWidget():
            self.centralWidget().deleteLater()

        t = THEMES[self.current_theme]

        # TOP BAR
        top_frame = QFrame()
        top_frame.setProperty("class", "QFrame_Top")
        top_layout = QHBoxLayout(top_frame)
        top_layout.setContentsMargins(10, 5, 10, 5)
        
        btn_load = QPushButton("1. Load Excel")
        btn_load.clicked.connect(self.load_file)
        
        self.lbl_status = QLabel("No file loaded")
        self.lbl_status.setStyleSheet("color: red; margin-left: 10px;")
        
        btn_theme = QPushButton(f"Theme: {self.current_theme}")
        btn_theme.clicked.connect(self.toggle_theme)

        export_img_btn = QPushButton("4. Export Image")
        export_img_btn.clicked.connect(self.export_image_cmd)
        
        export_xlsx_btn = QPushButton("3. Export Excel")
        export_xlsx_btn.clicked.connect(self.export_excel)
        
        clear_btn = QPushButton("2. Clear Grid")
        clear_btn.clicked.connect(self.clear_grid)

        top_layout.addWidget(btn_load)
        top_layout.addWidget(self.lbl_status)
        top_layout.addStretch()
        top_layout.addWidget(btn_theme)
        top_layout.addWidget(clear_btn)
        top_layout.addWidget(export_xlsx_btn)
        top_layout.addWidget(export_img_btn)
        

        # MAIN CONTENT
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(top_frame)

        splitter = QSplitter(Qt.Vertical)
        
        # ROSTER AREA
        self.roster_frame = QWidget()
        roster_layout = QVBoxLayout(self.roster_frame)
        roster_layout.setContentsMargins(0, 0, 0, 0)
        
        self.scroll_r = QScrollArea()
        self.scroll_r.setWidgetResizable(True)
        self.grid_container = QWidget()
        self.grid_layout = QGridLayout(self.grid_container)
        self.grid_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.scroll_r.setWidget(self.grid_container)
        
        roster_layout.addWidget(self.scroll_r)
        
        # DASHBOARD AREA
        self.dash_frame = QWidget()
        dash_layout = QVBoxLayout(self.dash_frame)
        dash_layout.setContentsMargins(0,0,0,0)
        
        legend_frame = QFrame()
        legend_frame.setProperty("class", "QFrame_Legend")
        legend_layout = QHBoxLayout(legend_frame)
        legend_layout.addWidget(QLabel("LEGEND:"))
        for cat, roles in CATEGORY_DATA.items():
             col = THEMES[self.current_theme]["cats"].get(cat, "#888")
             l = QLabel(f" ■ {cat} ")
             l.setStyleSheet(f"color: {col}; font-weight: bold;")
             legend_layout.addWidget(l)
        legend_layout.addStretch()
        
        dash_layout.addWidget(legend_frame)

        self.scroll_d = QScrollArea()
        self.scroll_d.setWidgetResizable(True)
        self.dash_container = QWidget()
        self.dash_container.setStyleSheet(f"background-color: {t['bg_main']};")
        self.dash_layout_grid = QGridLayout(self.dash_container)
        self.dash_layout_grid.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.scroll_d.setWidget(self.dash_container)
        
        dash_layout.addWidget(self.scroll_d)
        
        splitter.addWidget(self.roster_frame)
        splitter.addWidget(self.dash_frame)
        splitter.setSizes([400, 500])
        
        main_layout.addWidget(splitter)

    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel", "", "Excel Files (*.xlsx)")
        if not path: return
        success, msg = self.engine.load_file(path)
        if success:
            self.lbl_status.setText(f"Loaded: {os.path.basename(path)}")
            self.lbl_status.setStyleSheet("color: #4CAF50; margin-left: 10px;") # Green always ok
            self.engine.generate_draft()
            self.render_roster_grid()
            self.trigger_dashboard_update()
        else:
            QMessageBox.critical(self, "Error", msg)

    def clear_grid(self):
        if not self.combos: return
        ret = QMessageBox.question(self, "Confirm", "Clear all selections?", QMessageBox.Yes | QMessageBox.No)
        if ret == QMessageBox.Yes:
            for cb in self.combos.values(): 
                cb.blockSignals(True)
                cb.setCurrentIndex(-1)
                cb.blockSignals(False)
            self.on_selection_change()

    def render_roster_grid(self):
        # Clear layout
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
            
        self.combos = {}
        total_rows = len(self.engine.week_columns) + 1
        t = THEMES[self.current_theme]
        
        # Headers
        header_week = QLabel("Week")
        header_week.setStyleSheet(f"font-weight: bold; padding: 2px; color: {t['fg_pri']};")
        header_week.setFixedWidth(100)
        self.grid_layout.addWidget(header_week, 0, 0)
        
        # Vertical Sep
        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setLineWidth(2)
        sep.setStyleSheet(f"color: {t['input_border']};")
        self.grid_layout.addWidget(sep, 0, 1, total_rows, 1)

        current_col = 2
        prev_cat = None
        
        for role in ROLES_ORDER:
            this_cat_data = ROLE_TO_CAT_MAP[role]
            this_cat = this_cat_data["cat"]
            
            if prev_cat and this_cat != prev_cat:
                s = QFrame()
                s.setFrameShape(QFrame.VLine)
                s.setLineWidth(2)
                s.setStyleSheet(f"color: {t['input_border']};")
                self.grid_layout.addWidget(s, 0, current_col, total_rows, 1)
                current_col += 1
            
            lbl = QLabel(role)
            lbl.setStyleSheet(f"font-weight: bold; color: {this_cat_data['color']}; border: 1px solid {t['input_border']}; padding: 2px;")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setFixedWidth(110)
            self.grid_layout.addWidget(lbl, 0, current_col)
            
            for r, week in enumerate(self.engine.week_columns):
                row_idx = r + 1
                if role == ROLES_ORDER[0]: 
                    w_lbl = QLabel(week)
                    w_lbl.setStyleSheet(f"font-size: 11px; color: {t['fg_sec']};")
                    self.grid_layout.addWidget(w_lbl, row_idx, 0)

                cb = EnhancedComboBox(self.update_dropdown_options, week, role)
                cb.setEditable(False)
                cb.setFixedWidth(110)
                
                draft = self.engine.initial_roster[week].get(role, "")
                cb.addItem("")
                if draft:
                    cb.addItem(draft)
                    cb.setCurrentText(draft)
                
                cb.currentTextChanged.connect(self.on_selection_change)
                self.grid_layout.addWidget(cb, row_idx, current_col)
                self.combos[(week, role)] = cb

            prev_cat = this_cat
            current_col += 1
        
        self.trigger_dashboard_update()

    def update_dropdown_options(self, week, role, widget):
        capable = self.engine.availability_map[week][role]
        busy = []
        for r in ROLES_ORDER:
            if r == role: continue 
            if (week, r) in self.combos:
                val = self.combos[(week, r)].currentText()
                if val: busy.append(val)
        
        filtered = [p for p in capable if p not in busy]
        if "Cleanup" not in role: filtered.sort()
        
        current_val = widget.currentText()
        widget.blockSignals(True)
        widget.clear()
        widget.addItem("")
        widget.addItems(filtered)
        if current_val in filtered or current_val == "":
            widget.setCurrentText(current_val)
        else:
             widget.addItem(current_val)
             widget.setCurrentText(current_val)
             
        widget.blockSignals(False)
    
    def on_selection_change(self, text=None):
        self.update_locks()
        self.validate_all()
        # Debounced update
        self.trigger_dashboard_update()

    def trigger_dashboard_update(self):
        self.update_timer.start()

    def update_locks(self):
        for week in self.engine.week_columns:
            if (week, "Piano") not in self.combos: continue
            piano_val = self.combos[(week, "Piano")].currentText()
            bass_combo = self.combos[(week, "Bass")]
            if not piano_val:
                bass_combo.setCurrentIndex(-1)
                bass_combo.setEnabled(False)
            else:
                bass_combo.setEnabled(True)

    def validate_all(self):
        t = THEMES[self.current_theme]
        normal_col = t['fg_pri']
        bg = t['input_bg']
        for week in self.engine.week_columns:
            seen = {}
            dupes = []
            for role in ROLES_ORDER:
                val = self.combos[(week, role)].currentText()
                if val:
                    if val in seen: dupes.append(val)
                    seen[val] = True
            for role in ROLES_ORDER:
                w = self.combos[(week, role)]
                txt = w.currentText()
                if txt and txt in dupes: 
                     w.setStyleSheet(f"color: red; background-color: {bg};")
                else:
                     w.setStyleSheet(f"color: {normal_col}; background-color: {bg};")
            
            if not self.combos[(week, "Piano")].currentText() and self.combos[(week, "Bass")].currentText():
                 self.combos[(week, "Bass")].setStyleSheet(f"color: red; background-color: {bg};")

    def _perform_dashboard_update(self):
        # Actual Dashboard Logic
        # Clear
        while self.dash_layout_grid.count():
            item = self.dash_layout_grid.takeAt(0)
            if item.widget(): item.widget().deleteLater()

        assigned_map = {week: {} for week in self.engine.week_columns}
        serve_counts = {name: 0 for name in self.engine.all_members}
        cleanup_counts = {opt: 0 for opt in CLEANUP_OPTIONS}
        
        member_active_roles = {name: set() for name in self.engine.all_members}
        cleanup_active_roles = {opt: set() for opt in CLEANUP_OPTIONS}

        for week in self.engine.week_columns:
            for role in ROLES_ORDER:
                if (week, role) in self.combos:
                    name = self.combos[(week, role)].currentText()
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

        t = THEMES[self.current_theme]
        col_idx = 0
        for cat_name, roles in CATEGORY_DATA.items():
            cat_col = t["cats"].get(cat_name, "#888")
            
            # Category Header
            lbl = QLabel(cat_name)
            # Use fixed white text for header if background is dark/colored
            lbl.setStyleSheet(f"background-color: {cat_col}; color: white; font-weight: bold; padding: 4px;")
            lbl.setAlignment(Qt.AlignCenter)
            self.dash_layout_grid.addWidget(lbl, 0, col_idx, 1, len(roles))
            
            current_role_col = col_idx
            for role in roles:
                role_lbl = QLabel(role)
                role_lbl.setStyleSheet(f"background-color: {t['bg_sec']}; color: {t['fg_pri']}; font-weight: bold; border: 1px solid {t['input_border']};")
                role_lbl.setAlignment(Qt.AlignCenter)
                self.dash_layout_grid.addWidget(role_lbl, 1, current_role_col)

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
                    cell = self._create_member_cell(m, assigned_map)
                    self.dash_layout_grid.addWidget(cell, row_idx, current_role_col)
                    row_idx += 1
                
                current_role_col += 1
            
            # Spacer col
            spacer = QFrame()
            spacer.setFixedWidth(15)
            self.dash_layout_grid.addWidget(spacer, 0, current_role_col)
            
            col_idx = current_role_col + 1

    def _create_member_cell(self, m, assigned_map):
        t = THEMES[self.current_theme]
        container = QFrame()
        bg_col = t['bg_sec'] # Default inactive bg
        
        # Determine BG
        # Active color logic
        if m["is_active"]:
            if m["count"] >= 3: bg_col = t['dash_bg_warn']
            elif m["count"] >= 1: bg_col = t['dash_bg_notice']
        
        border_col = t['input_border']
        container.setStyleSheet(f"background-color: {bg_col}; border: 1px solid {border_col};")
        layout = QHBoxLayout(container)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(2)
        
        text_col = t['active_cell_text'] if m["is_active"] else t['fg_pri']

        name_lbl = QLabel(m["name"])
        name_lbl.setStyleSheet(f"border: none; color: {text_col};")
        layout.addWidget(name_lbl)
        
        layout.addStretch()

        if m["name"] in CLEANUP_OPTIONS:
            l = QLabel("----")
            l.setStyleSheet(f"border: none; color: {text_col};")
            layout.addWidget(l)
        else:
            for i, char in enumerate(m["avail"]):
                color = text_col
                text = "O"
                week_key = self.engine.week_columns[i]
                if char == "X":
                    color = t['dash_text_unavail']; text = "X"
                else:
                    if m["name"] in assigned_map[week_key]:
                        role_assigned = assigned_map[week_key][m["name"]]
                        cat_info = ROLE_TO_CAT_MAP.get(role_assigned)
                        if cat_info: color = cat_info["color"]
                
                dot = QLabel(text)
                dot.setStyleSheet(f"color: {color}; font-weight: bold; border: none; background: transparent;")
                layout.addWidget(dot)

        cnt_lbl = QLabel(f"({m['count']})")
        cnt_lbl.setStyleSheet(f"border: none; font-size: 10px; color: {text_col};")
        layout.addWidget(cnt_lbl)
        
        return container

    def export_excel(self):
        if not self.engine.week_columns: return
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "roster.xlsx", "Excel Files (*.xlsx)")
        if not path: return

        data = []
        for week in self.engine.week_columns:
            row = {"Week": week}
            for role in ROLES_ORDER:
                row[role] = self.combos[(week, role)].currentText()
            
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
        QMessageBox.information(self, "Done", "Excel Exported!")

    # ==========================================
    # IMAGE EXPORT
    # ==========================================
    def export_image_cmd(self):
        if not HAS_PIL:
            QMessageBox.critical(self, "Error", "Pillow not installed.")
            return
        if not self.engine.week_columns: return
        
        path, _ = QFileDialog.getSaveFileName(self, "Save Image", "roster.png", "PNG Image (*.png)")
        if not path: return

        # Gathers Data
        assigned_map = {week: {} for week in self.engine.week_columns}
        serve_counts = {name: 0 for name in self.engine.all_members}
        cleanup_counts = {opt: 0 for opt in CLEANUP_OPTIONS}
        
        member_active_roles = {name: set() for name in self.engine.all_members}
        cleanup_active_roles = {opt: set() for opt in CLEANUP_OPTIONS}

        roster_data = {}
        for week in self.engine.week_columns:
            roster_data[week] = {}
            for role in ROLES_ORDER:
                val = self.combos[(week, role)].currentText()
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
        # USE LIGHT THEME COLORS FOR EXPORT for consistency
        EXPORT_CATS = THEMES["Light"]["cats"]

        COL_W = 160
        ROW_H = 30
        MARGIN = 20
        FONT_SIZE = 12
        CAT_SPACER = 10
        
        roster_w = COL_W * (len(ROLES_ORDER) + 1)
        dash_w = 0
        for cat, roles in CATEGORY_DATA.items():
            dash_w += (len(roles) * COL_W) + CAT_SPACER
        dash_w -= CAT_SPACER 
        
        img_w = max(roster_w, dash_w) + (MARGIN * 2)
        
        max_mem_rows = 0
        for cat, roles in CATEGORY_DATA.items():
            for role in roles:
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
        for cat_name, roles in CATEGORY_DATA.items():
            w = len(roles) * COL_W
            color = EXPORT_CATS.get(cat_name, "black")
            draw.rectangle([x_curs, y, x_curs+w, y+ROW_H], fill="white", outline="black")
            text_len = draw.textlength(cat_name, font=font_bold)
            draw.text((x_curs + (w-text_len)/2, y+5), cat_name, fill=color, font=font_bold)
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
        
        for cat_name, roles in CATEGORY_DATA.items():
            width = len(roles) * COL_W
            color = EXPORT_CATS.get(cat_name, "black")
            draw.rectangle([x_curs, y, x_curs+width, y+ROW_H], fill=color, outline="black")
            draw.text((x_curs+5, y+5), cat_name, fill="white", font=font_bold)
            
            role_x = x_curs
            for role in roles:
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
                        if m["count"] >= 3: bg = "#ffcccc" # Light Red
                        elif m["count"] >= 1: bg = "#ffeeb0" # Light Yellow
                    
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
                                    
                                    # Lookup category color for exported image (Light Theme)
                                    c_data = CATEGORY_DATA
                                    found_cat = None
                                    for ccat, rroles in c_data.items():
                                        if assigned_role in rroles: found_cat = ccat; break
                                    
                                    if found_cat: color = EXPORT_CATS.get(found_cat, "black")

                            draw.text((dot_start_x + (i*11), mem_y+5), char, fill=color, font=font_bold)
                    
                    mem_y += ROW_H
                role_x += COL_W
            x_curs += width + CAT_SPACER
        
        img.save(path)
        QMessageBox.information(self, "Success", "Image Exported!")

def exception_hook(exctype, value, tb):
    """
    Global function to catch unhandled exceptions.
    Prints to console and displays a popup message.
    """
    traceback_formated = ''.join(traceback.format_exception(exctype, value, tb))
    print(traceback_formated, file=sys.stderr)
    
    if QApplication.instance():
        error_msg = f"An unexpected error occurred:\n{exctype.__name__}: {value}\n\n{traceback_formated}"
        QMessageBox.critical(None, "Critical Error", error_msg)
    else:
        sys.__excepthook__(exctype, value, tb)

if __name__ == "__main__":
    sys.excepthook = exception_hook
    app = QApplication(sys.argv)
    window = RosterApp()
    window.show()
    sys.exit(app.exec())