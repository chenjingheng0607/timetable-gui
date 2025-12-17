# config.py

# ROLES CONFIGURATION
ROLES_ORDER = [
    "MD", "Lead", "Vocal", "Piano", "Drum/Cajon", "Bass", "Guitar", 
    "PPT", "Sound", "Lighting/OBS", 
    "MC", 
    "Usher 1", "Usher 2", "Usher 3", 
    "Cleanup 1", "Cleanup 2"
]

# Roles that are eligible to be MD
BAND_ROLES = ["Piano", "Bass", "Guitar"]

# Fixed Options for Cleanup
CLEANUP_OPTIONS = ["LHW", "UF", "LB", "YGSS", "SJS", "PK"]

# Excel Code Mapping
INSTRUMENT_MAP = {
    "WL": "Lead", "V": "Vocal", "P": "Piano", "G": "Guitar", 
    "B": "Bass", "D": "Drum/Cajon", "PPT": "PPT", 
    "S": "Sound", "SOUND": "Sound", 
    "OBS": "Lighting/OBS", "LIGHT": "Lighting/OBS", "L": "Lighting/OBS",
    "MC": "MC", "USHER": "Usher",
    "MD": "MD", "M.D.": "MD"
}

# DASHBOARD CATEGORIES & COLORS
CATEGORY_CONFIG = {
    "Praise & Worship": {
        "roles": ["MD", "Lead", "Vocal", "Piano", "Drum/Cajon", "Bass", "Guitar"],
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

# Auto-generate Map
def build_role_map(cat_colors):
    mapping = {}
    for cat, data in CATEGORY_CONFIG.items():
        color = cat_colors.get(cat, data["color"])
        for r in data["roles"]:
            mapping[r] = {"cat": cat, "color": color}
    return mapping

ROLE_TO_CAT_MAP = {}
for cat, data in CATEGORY_CONFIG.items():
    for r in data["roles"]:
        ROLE_TO_CAT_MAP[r] = {"cat": cat, "color": data["color"]}

# UI THEMES
THEMES = {
    "Dark": {
        "bg_main": "#1e1e1e", "bg_sec": "#252526", "fg_pri": "#ffffff", "fg_sec": "#cccccc",
        "input_bg": "#333333", "input_border": "#3e3e42", "input_sel": "#264f78",
        "btn_bg": "#0e639c", "btn_fg": "#ffffff",
        "dash_bg_warn": "#4a0000", 
        "dash_bg_notice": "#4a3b00",
        "dash_text_avail": "white",
        "dash_text_unavail": "#666666",
        "active_cell_text": "#ffffff",
        "cats": {
            "Praise & Worship": "#ff6666",
            "FPH": "#66b3ff",
            "MC": "#d9b3ff",
            "Usher": "#ffdf80",
            "LG": "#66ff66"
        }
    },
    "Light": {
        "bg_main": "#fafafa", "bg_sec": "#e0e0e0", "fg_pri": "#000000", "fg_sec": "#333333",
        "input_bg": "#ffffff", "input_border": "#bdc3c7", "input_sel": "#0078d7",
        "btn_bg": "#0078d7", "btn_fg": "#ffffff",
        "dash_bg_warn": "#FFEBEE",
        "dash_bg_notice": "#FFFDE7",
        "dash_text_avail": "black",
        "dash_text_unavail": "#999999",
        "active_cell_text": "#000000",
        "cats": {
            "Praise & Worship": "#c00000",
            "FPH": "#0070c0",
            "MC": "#7030a0",
            "Usher": "#ffc000",
            "LG": "#00b050"
        }
    }
}