# gui.py
import os
import functools
import pandas as pd
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                               QPushButton, QLabel, QScrollArea, QFrame, 
                               QFileDialog, QMessageBox, QComboBox, QGridLayout, QSplitter)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QFont

from config import *
from logic import RosterEngine

try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

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
        self.update_timer = QTimer()
        self.update_timer.setSingleShot(True)
        self.update_timer.setInterval(80) 
        self.update_timer.timeout.connect(self._perform_dashboard_update)
        if not HAS_PIL:
            QMessageBox.warning(self, "Missing Library", "Pillow not found. Image export disabled.")
        self.apply_theme(self.current_theme)

    def apply_theme(self, theme_name):
        saved_state = {}
        if hasattr(self, 'combos') and self.combos:
            for key, cb in self.combos.items(): saved_state[key] = cb.currentText()
        
        saved_status = self.lbl_status.text() if hasattr(self, 'lbl_status') else "No file loaded"
        self.current_theme = theme_name
        t = THEMES[theme_name]
        
        global ROLE_TO_CAT_MAP
        ROLE_TO_CAT_MAP = build_role_map(t["cats"])

        css = f"""
            QWidget {{ background-color: {t['bg_main']}; color: {t['fg_pri']}; font-family: "Segoe UI", Arial; }}
            QFrame, QScrollArea {{ background-color: {t['bg_main']}; border: none; }}
            QLineEdit, QComboBox, QAbstractItemView {{ 
                background-color: {t['input_bg']}; color: {t['fg_pri']}; border: 1px solid {t['input_border']};
                selection-background-color: {t['input_sel']}; selection-color: white; }}
            QComboBox::drop-down {{ border: none; }}
            QComboBox:disabled {{ background-color: {t['bg_sec']}; color: {t['fg_sec']}; }}
            QPushButton {{ background-color: {t['input_bg']}; border: 1px solid {t['input_border']}; padding: 6px 12px; border-radius: 4px; min-width: 70px; }}
            QPushButton:hover {{ background-color: {t['input_sel']}; color: white; }}
            QSplitter::handle {{ background-color: {t['input_border']}; }}
        """
        self.setStyleSheet(css)
        self._build_ui() 
        
        self.lbl_status.setText(saved_status)
        if "Loaded:" in saved_status:
            self.lbl_status.setStyleSheet("color: #4CAF50; margin-left: 10px;")
        if self.engine.week_columns:
            self.render_roster_grid()
            for key, val in saved_state.items():
                if key in self.combos:
                    self.combos[key].blockSignals(True)
                    if val and self.combos[key].findText(val) == -1: self.combos[key].addItem(val)
                    self.combos[key].setCurrentText(val)
                    self.combos[key].blockSignals(False)
            self.update_locks()
            self.validate_all()
            self.trigger_dashboard_update()

    def toggle_theme(self):
        self.apply_theme("Light" if self.current_theme == "Dark" else "Dark")

    def _build_ui(self):
        if self.centralWidget(): self.centralWidget().deleteLater()
        t = THEMES[self.current_theme]
        
        top = QFrame()
        top_l = QHBoxLayout(top)
        
        btn_load = QPushButton("1. Load Excel"); btn_load.clicked.connect(self.load_file)
        self.lbl_status = QLabel("No file loaded"); self.lbl_status.setStyleSheet("color: red; margin-left: 10px;")
        
        top_l.addWidget(btn_load); top_l.addWidget(self.lbl_status); top_l.addStretch()
        
        btn_theme = QPushButton(f"Theme: {self.current_theme}"); btn_theme.clicked.connect(self.toggle_theme)
        btn_clear = QPushButton("2. Clear"); btn_clear.clicked.connect(self.clear_grid)
        btn_ex_xl = QPushButton("3. Export Excel"); btn_ex_xl.clicked.connect(self.export_excel)
        btn_ex_img = QPushButton("4. Export Image"); btn_ex_img.clicked.connect(self.export_image_cmd)
        
        for b in [btn_theme, btn_clear, btn_ex_xl, btn_ex_img]: top_l.addWidget(b)

        central = QWidget(); self.setCentralWidget(central)
        main_l = QVBoxLayout(central); main_l.setContentsMargins(0,0,0,0); main_l.addWidget(top)
        
        splitter = QSplitter(Qt.Vertical)
        
        self.grid_c = QWidget(); self.grid_l = QGridLayout(self.grid_c); self.grid_l.setAlignment(Qt.AlignTop|Qt.AlignLeft)
        scroll_r = QScrollArea(); scroll_r.setWidgetResizable(True); scroll_r.setWidget(self.grid_c)
        splitter.addWidget(scroll_r)
        
        dash_frame = QWidget(); dash_l = QVBoxLayout(dash_frame)
        self.dash_c = QWidget(); self.dash_l = QGridLayout(self.dash_c); self.dash_l.setAlignment(Qt.AlignTop|Qt.AlignLeft)
        scroll_d = QScrollArea(); scroll_d.setWidgetResizable(True); scroll_d.setWidget(self.dash_c)
        
        # Legend
        leg = QFrame(); leg_l = QHBoxLayout(leg)
        leg_l.addWidget(QLabel("LEGEND:"))
        for c, d in CATEGORY_CONFIG.items():
            l = QLabel(f" ■ {c} "); l.setStyleSheet(f"color: {d['color']}; font-weight: bold;")
            leg_l.addWidget(l)
        leg_l.addStretch()
        dash_l.addWidget(leg); dash_l.addWidget(scroll_d)
        splitter.addWidget(dash_frame)
        splitter.setSizes([400, 500])
        main_l.addWidget(splitter)

    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel", "", "Excel Files (*.xlsx)")
        if not path: return
        success, msg = self.engine.load_file(path)
        if success:
            self.lbl_status.setText(f"Loaded: {os.path.basename(path)}")
            self.lbl_status.setStyleSheet("color: #4CAF50; margin-left: 10px;")
            self.engine.generate_draft()
            self.render_roster_grid()
            self.trigger_dashboard_update()
        else:
            QMessageBox.critical(self, "Error", msg)

    def clear_grid(self):
        if not self.combos: return
        if QMessageBox.question(self, "Confirm", "Clear all?") == QMessageBox.Yes:
            for cb in self.combos.values(): cb.blockSignals(True); cb.setCurrentIndex(-1); cb.blockSignals(False)
            self.on_selection_change()

    def render_roster_grid(self):
        while self.grid_l.count():
            item = self.grid_l.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        
        self.combos = {}
        rows = len(self.engine.week_columns) + 1
        t = THEMES[self.current_theme]
        
        # Headers
        self.grid_l.addWidget(QLabel("Week"), 0, 0)
        sep = QFrame(); sep.setFrameShape(QFrame.VLine); sep.setStyleSheet(f"color: {t['input_border']}")
        self.grid_l.addWidget(sep, 0, 1, rows, 1)
        
        col, prev_cat = 2, None
        for role in ROLES_ORDER:
            this_cat = ROLE_TO_CAT_MAP[role]["cat"]
            if prev_cat and this_cat != prev_cat:
                s = QFrame(); s.setFrameShape(QFrame.VLine); s.setStyleSheet(f"color: {t['input_border']}")
                self.grid_l.addWidget(s, 0, col, rows, 1); col += 1
            
            l = QLabel(role); l.setStyleSheet(f"font-weight: bold; color: {ROLE_TO_CAT_MAP[role]['color']}; padding: 2px;")
            l.setAlignment(Qt.AlignCenter); l.setFixedWidth(110)
            self.grid_l.addWidget(l, 0, col)
            
            for r, week in enumerate(self.engine.week_columns):
                if role == ROLES_ORDER[0]: self.grid_l.addWidget(QLabel(week), r+1, 0)
                cb = EnhancedComboBox(self.update_dropdown_options, week, role)
                cb.setEditable(False); cb.setFixedWidth(110)
                
                draft = self.engine.initial_roster[week].get(role, "")
                display_val = ""
                if draft:
                    display_val = draft
                    md_draft = self.engine.initial_roster[week].get("MD", "")
                    # Note: md_draft is just a name "John".
                    if draft == md_draft and md_draft:
                        display_val += " (MD)"
                
                cb.addItem(""); 
                if display_val: cb.addItem(display_val); cb.setCurrentText(display_val)
                cb.currentTextChanged.connect(self.on_selection_change)
                self.grid_l.addWidget(cb, r+1, col)
                self.combos[(week, role)] = cb
            prev_cat = this_cat; col += 1
        self.update_locks()
        self.trigger_dashboard_update()

    def update_dropdown_options(self, week, role, widget):
        curr_text = widget.currentText()
        curr_clean = curr_text.replace(" (MD)", "").strip()
        
        if role == "MD":
            # Show ANYONE assigned to a band role this week who has MD capability
            potentials = []
            for br in BAND_ROLES:
                if (week, br) in self.combos:
                    val = self.combos[(week, br)].currentText()
                    val_clean = val.replace(" (MD)", "").strip()
                    if val_clean:
                        roles = self.engine.all_members.get(val_clean, {}).get("Roles", [])
                        if "MD" in roles: potentials.append(val_clean)
            filtered = list(set(potentials))
            filtered.sort()
        else:
            # FIX: Filter out anyone busy in ANY other column (except MD)
            capable = self.engine.availability_map[week][role]
            busy = []
            for r in ROLES_ORDER:
                if r == role: continue
                if r == "MD": continue # MD is allowed to overlap
                if (week, r) in self.combos:
                    val = self.combos[(week, r)].currentText()
                    val_clean = val.replace(" (MD)", "").strip()
                    if val_clean: busy.append(val_clean)
            
            # Allow current user to stay selected (don't filter self out if re-opening box)
            filtered = [p for p in capable if p not in busy or p == curr_clean]
            if "Cleanup" not in role: filtered.sort()
        
        widget.blockSignals(True)
        widget.clear(); widget.addItem("")
        
        # Check active MD for this week to apply suffix
        md_name = ""
        if (week, "MD") in self.combos:
            md_name = self.combos[(week, "MD")].currentText().replace(" (MD)", "").strip()

        # Add items with suffix logic
        for p in filtered:
            display = p
            # Logic: If this person IS the active MD for the week, add suffix
            if p == md_name:
                display = p + " (MD)"
            widget.addItem(display)
            
        # Restore selection if possible
        target = curr_clean + " (MD)" if curr_clean == md_name and md_name else curr_clean
        if widget.findText(target) != -1:
            widget.setCurrentText(target)
        elif widget.findText(curr_text) != -1:
             widget.setCurrentText(curr_text)
            
        widget.blockSignals(False)

    def on_selection_change(self, _=None):
        sender = self.sender()
        if isinstance(sender, EnhancedComboBox):
            self.update_week_visuals(sender.week)

        self.update_locks()
        self.validate_all()
        self.trigger_dashboard_update()

    def update_week_visuals(self, week):
        # Identify current MD for the week
        md_name = ""
        if (week, "MD") in self.combos:
            md_name = self.combos[(week, "MD")].currentText().replace(" (MD)", "").strip()
        
        # Update suffix for all roles in this week
        for role in ROLES_ORDER:
            if (week, role) in self.combos:
                cb = self.combos[(week, role)]
                curr_txt = cb.currentText()
                curr_clean = curr_txt.replace(" (MD)", "").strip()
                
                if not curr_clean: continue
                
                # Only the active MD gets the suffix
                should_have_suffix = (curr_clean == md_name) and md_name != ""
                new_txt = curr_clean + " (MD)" if should_have_suffix else curr_clean
                
                if curr_txt != new_txt:
                    cb.blockSignals(True)
                    idx = cb.currentIndex()
                    if idx >= 0:
                        cb.setItemText(idx, new_txt)
                    cb.blockSignals(False)

    def update_locks(self):
        for week in self.engine.week_columns:
            if (week, "Piano") in self.combos:
                p = self.combos[(week, "Piano")].currentText()
                b = self.combos[(week, "Bass")]
                if not p: b.setCurrentIndex(-1); b.setEnabled(False)
                else: b.setEnabled(True)

    def validate_all(self):
        bg = THEMES[self.current_theme]['input_bg']
        for week in self.engine.week_columns:
            seen, dupes = {}, []
            for role in ROLES_ORDER:
                if role == "MD": continue
                val = self.combos[(week, role)].currentText().replace(" (MD)", "")
                if val:
                    if val in seen: dupes.append(val)
                    seen[val] = True
            
            for role in ROLES_ORDER:
                w = self.combos[(week, role)]
                
                # If disabled (e.g. Bass locked), don't override style
                if not w.isEnabled():
                    w.setStyleSheet("")
                    continue

                txt = w.currentText()
                val = txt.replace(" (MD)", "")
                style = f"color: {THEMES[self.current_theme]['fg_pri']}; background-color: {bg};"
                
                # Highlight MD if not in band
                if role == "MD" and val:
                    in_band = any(self.combos[(week, br)].currentText().replace(" (MD)", "") == val for br in BAND_ROLES if (week, br) in self.combos)
                    if not in_band: style = f"color: red; background-color: {bg};"
                # Highlight Dupes
                elif val and val in dupes: style = f"color: red; background-color: {bg};"
                # Highlight Bass without Piano
                elif role == "Bass" and not self.combos[(week, "Piano")].currentText() and val:
                    style = f"color: red; background-color: {bg};"
                
                w.setStyleSheet(style)

    def trigger_dashboard_update(self): self.update_timer.start()

    def _perform_dashboard_update(self):
        while self.dash_l.count(): 
            item = self.dash_l.takeAt(0)
            if item.widget(): item.widget().deleteLater()

        assigned_map = {w: {} for w in self.engine.week_columns}
        counts = {n: 0 for n in self.engine.all_members}
        cl_counts = {o: 0 for o in CLEANUP_OPTIONS}
        mem_active = {n: set() for n in self.engine.all_members}
        cl_active = {o: set() for o in CLEANUP_OPTIONS}

        for w in self.engine.week_columns:
            for r in ROLES_ORDER:
                if (w, r) in self.combos:
                    val = self.combos[(w, r)].currentText().replace(" (MD)", "")
                    if val:
                        assigned_map[w][val] = r
                        if "Cleanup" in r:
                            if val in cl_counts: 
                                cl_counts[val] += 1; cl_active[val].add(r)
                        else:
                            if r != "MD": # MD doesn't increment "Serving Load" count
                                if val in counts: 
                                    counts[val] += 1; mem_active[val].add(r)

        t = THEMES[self.current_theme]
        col = 0
        for cat, data in CATEGORY_CONFIG.items():
            roles = data["roles"]
            l = QLabel(cat); l.setStyleSheet(f"background-color: {data['color']}; color: white; font-weight: bold;")
            l.setAlignment(Qt.AlignCenter)
            self.dash_l.addWidget(l, 0, col, 1, len(roles))
            
            cur_r_col = col
            for role in roles:
                rl = QLabel(role); rl.setStyleSheet(f"background-color: {t['bg_sec']}; font-weight: bold; border: 1px solid {t['input_border']};")
                rl.setAlignment(Qt.AlignCenter)
                self.dash_l.addWidget(rl, 1, cur_r_col)
                
                members = []
                if cat == "LG":
                    for o in CLEANUP_OPTIONS:
                        act = role in cl_active.get(o, set())
                        sv = 2
                        if act: sv = 0 if cl_counts[o]>=3 else 1
                        members.append({"name": o, "av": "XXXX", "c": cl_counts[o], "act": act, "sv": sv})
                else:
                    for n, d in self.engine.all_members.items():
                        # Capable?
                        can = False
                        if "Usher" in role: can = "Usher" in d["Roles"]
                        elif role in d["Roles"]: can = True
                        
                        if can:
                            act = role in mem_active.get(n, set())
                            cnt = counts.get(n, 0)
                            sv = 2
                            if act: sv = 0 if cnt>=3 else 1
                            members.append({"name": n, "av": d["AvailString"], "c": cnt, "act": act, "sv": sv})
                
                members.sort(key=lambda x: (x["sv"], -x["c"], x["name"]))
                
                ridx = 2
                for m in members:
                    self.dash_l.addWidget(self._create_mem_cell(m, assigned_map), ridx, cur_r_col)
                    ridx += 1
                cur_r_col += 1
            
            sp = QFrame(); sp.setFixedWidth(15)
            self.dash_l.addWidget(sp, 0, cur_r_col)
            col = cur_r_col + 1

    def _create_mem_cell(self, m, assigned_map):
        t = THEMES[self.current_theme]
        f = QFrame()
        bg = t['bg_sec']
        if m["act"]: bg = t['dash_bg_warn'] if m["c"]>=3 else t['dash_bg_notice']
        
        f.setStyleSheet(f"background-color: {bg}; border: 1px solid {t['input_border']};")
        l = QHBoxLayout(f); l.setContentsMargins(2,2,2,2); l.setSpacing(2)
        
        tc = t['active_cell_text'] if m["act"] else t['fg_pri']
        nm = QLabel(m["name"]); nm.setStyleSheet(f"border: none; color: {tc};"); l.addWidget(nm)
        l.addStretch()
        
        if m["name"] in CLEANUP_OPTIONS:
            l.addWidget(QLabel("----"))
        else:
            for i, c in enumerate(m["av"]):
                col = tc
                txt = "O"
                if c == "X": col = t['dash_text_unavail']; txt = "X"
                else:
                    # Check assignment color
                    wk = self.engine.week_columns[i]
                    if m["name"] in assigned_map[wk]:
                        rc = assigned_map[wk][m["name"]]
                        col = ROLE_TO_CAT_MAP[rc]["color"]
                dt = QLabel(txt); dt.setStyleSheet(f"color: {col}; font-weight: bold; border:none; background:transparent;"); l.addWidget(dt)
        
        ct = QLabel(f"({m['c']})"); ct.setStyleSheet(f"border:none; font-size:10px; color:{tc};"); l.addWidget(ct)
        return f

    def export_excel(self):
        if not self.engine.week_columns: return
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "roster.xlsx", "Excel Files (*.xlsx)")
        if not path: return
        
        data = []
        for w in self.engine.week_columns:
            r = {"Week": w}
            for role in ROLES_ORDER: 
                if role == "MD": continue
                r[role] = self.combos[(w, role)].currentText()
            
            # Band Mode
            hd = r.get("Drum/Cajon", "") != ""; hk = r.get("Piano", "") != ""; hb = r.get("Bass", "") != ""
            mode = "INCOMPLETE"
            if hb: mode = "FULL BAND"
            elif hd and hk: mode = "ACOUSTIC SET"
            
            fr = {"Week": w, "Band Mode": mode}
            fr.update({k: v for k, v in r.items() if k != "Week"})
            data.append(fr)
        
        pd.DataFrame(data).to_excel(path, index=False)
        QMessageBox.information(self, "Done", "Exported!")

    def export_image_cmd(self):
        if not HAS_PIL: return
        if not self.engine.week_columns: return
        path, _ = QFileDialog.getSaveFileName(self, "Save Img", "roster.png", "PNG (*.png)")
        if not path: return

        # Gather Data
        assigned_map = {w: {} for w in self.engine.week_columns}
        counts = {n: 0 for n in self.engine.all_members}
        cl_counts = {o: 0 for o in CLEANUP_OPTIONS}
        act_roles = {n: set() for n in self.engine.all_members}
        cl_act = {o: set() for o in CLEANUP_OPTIONS}
        
        r_data = {}
        for w in self.engine.week_columns:
            r_data[w] = {}
            for role in ROLES_ORDER:
                val = self.combos[(w, role)].currentText()
                r_data[w][role] = val
                clean_val = val.replace(" (MD)", "")
                if clean_val:
                    assigned_map[w][clean_val] = role
                    if "Cleanup" in role:
                        cl_counts[clean_val] += 1; cl_act[clean_val].add(role)
                    else:
                        if role != "MD":
                            # Defensive check for name existing in data
                            if clean_val in counts:
                                counts[clean_val] += 1; act_roles[clean_val].add(role)

        # Drawing
        COL_W = 160; ROW_H = 30; MARGIN = 20; SP = 10
        EC = THEMES["Light"]["cats"]
        
        EXPORT_ROLES = [r for r in ROLES_ORDER if r != "MD"]
        
        rw = COL_W * (len(EXPORT_ROLES)+1)
        dw = 0
        cat_widths = {}
        for cat, data in CATEGORY_CONFIG.items():
            # Count only non-MD roles for this category
            valid_roles = [r for r in data["roles"] if r != "MD"]
            w = len(valid_roles)*COL_W
            if w > 0: # Only count if category has remaining roles
                dw += w + SP
                cat_widths[cat] = w
        dw -= SP
        iw = max(rw, dw) + (MARGIN*2)
        
        # Height calc
        mx_rows = 0
        for _, data in CATEGORY_CONFIG.items():
            roles = data["roles"]
            for r in roles:
                c = len(CLEANUP_OPTIONS) if "Cleanup" in r else 0
                if "Cleanup" not in r:
                    for n, d in self.engine.all_members.items():
                        can = False
                        if "Usher" in r: can = "Usher" in d["Roles"]
                        elif r in d["Roles"]: can = True
                        if can: c += 1
                mx_rows = max(mx_rows, c)
        
        ih = ROW_H*(len(self.engine.week_columns)+2) + 60 + ROW_H*(mx_rows+3) + MARGIN*2
        
        img = Image.new("RGB", (iw, ih), "white")
        draw = ImageDraw.Draw(img)
        try: font = ImageFont.truetype("arial.ttf", 12); fontb = ImageFont.truetype("arialbd.ttf", 12)
        except: font = ImageFont.load_default(); fontb = font

        # Draw Roster
        y = MARGIN; x = (iw - rw)//2
        cur_x = x + COL_W
        for cat, data in CATEGORY_CONFIG.items():
            w = cat_widths.get(cat, 0)
            if w > 0:
                draw.rectangle([cur_x, y, cur_x+w, y+ROW_H], fill="white", outline="black")
                tl = draw.textlength(cat, fontb)
                draw.text((cur_x+(w-tl)/2, y+5), cat, fill=EC.get(cat, "black"), font=fontb)
                cur_x += w
        y += ROW_H
        
        y += ROW_H
        
        draw.rectangle([x, y, x+COL_W, y+ROW_H], outline="black")
        cur_x = x + COL_W
        for r in EXPORT_ROLES:
            draw.rectangle([cur_x, y, cur_x+COL_W, y+ROW_H], outline="black", fill="#f0f0f0")
            draw.text((cur_x+5, y+5), r, fill="black", font=fontb)
            cur_x += COL_W
        y += ROW_H
        
        for w in self.engine.week_columns:
            draw.rectangle([x, y, x+COL_W, y+ROW_H], outline="black")
            draw.text((x+5, y+5), w, fill="black", font=font)
            cur_x = x + COL_W
            for r in EXPORT_ROLES:
                v = r_data[w][r]
                draw.rectangle([cur_x, y, cur_x+COL_W, y+ROW_H], outline="black")
                if v: draw.text((cur_x+5, y+5), v, fill="black", font=font)
                cur_x += COL_W
            y += ROW_H

        # Draw Dash
        y += 60; x = (iw - dw)//2; cur_x = x
        for cat, data in CATEGORY_CONFIG.items():
            roles = [r for r in data["roles"] if r != "MD"] # Filter MD out
            if not roles: continue
            
            w = len(roles)*COL_W
            draw.rectangle([cur_x, y, cur_x+w, y+ROW_H], fill=EC.get(cat, "black"), outline="black")
            draw.text((cur_x+5, y+5), cat, fill="white", font=fontb)
            
            rx = cur_x
            for r in roles:
                draw.rectangle([rx, y+ROW_H, rx+COL_W, y+ROW_H*2], fill="#eee", outline="black")
                draw.text((rx+5, y+ROW_H+5), r, fill="black", font=fontb)
                
                mems = []
                if cat == "LG":
                    for o in CLEANUP_OPTIONS:
                        act = r in cl_act.get(o, set())
                        sv = 2
                        if act: sv = 0 if cl_counts[o]>=3 else 1
                        mems.append({"n": o, "av": "XXXX", "c": cl_counts[o], "act": act, "sv": sv})
                else:
                    for n, d in self.engine.all_members.items():
                        # Capable?
                        can = False
                        if "Usher" in r: can = "Usher" in d["Roles"]
                        elif r in d["Roles"]: can = True
                        
                        if can:
                            act = r in act_roles.get(n, set())
                            cnt = counts.get(n, 0)
                            sv = 2
                            if act: sv = 0 if cnt>=3 else 1
                            mems.append({"n": n, "av": d["AvailString"], "c": cnt, "act": act, "sv": sv})
                
                mems.sort(key=lambda x: (x["sv"], -x["c"], x["n"]))
                
                my = y + ROW_H*2
                for m in mems:
                    bg = "white"
                    if m["act"]: bg = "#ffcccc" if m["c"]>=3 else "#ffeeb0"
                    
                    draw.rectangle([rx, my, rx+COL_W, my+ROW_H], fill=bg, outline="black")
                    draw.text((rx+5, my+5), m["n"], fill="black", font=font)
                    
                    ctxt = f"({m['c']})"
                    cln = draw.textlength(ctxt, font=font)
                    cx = rx + COL_W - cln - 5
                    draw.text((cx, my+5), ctxt, fill="black", font=font)
                    
                    if cat != "LG":
                        sx = cx - 45
                        for i, c in enumerate(m["av"]):
                            col = "black"
                            if c == "X": col = "#ccc"
                            else:
                                wk = self.engine.week_columns[i]
                                if m["n"] in assigned_map[wk]:
                                    ar = assigned_map[wk][m["n"]]
                                    # Find color
                                    for ccat, data in CATEGORY_CONFIG.items():
                                        if ar in data["roles"]: col = EC.get(ccat, "black"); break
                            draw.text((sx + i*11, my+5), c, fill=col, font=fontb)
                    my += ROW_H
                rx += COL_W
            cur_x += w + SP
        
        img.save(path)
        QMessageBox.information(self, "Success", "Image Saved!")