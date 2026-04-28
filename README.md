# Auto-Roster (Timetable GUI)

Auto-Roster is a desktop application designed to automate and manage the scheduling of rosters for teams, specifically tailored for church worship teams or similar volunteer groups. It provides a visual interface to generate, edit, and export rosters while managing member availability and preventing burnout.

## Features

* **Automated Drafting**: Generates a draft roster based on member availability, past usage, and role capabilities to minimize burnout.
* **Smart MD Handling**:
    * Automatically identifies MDs based on the selected band members.
    * Visually tags the active MD with `(MD)` in the grid.
    * Excludes the redundant "MD" column from exports to keep the output clean.
* **Logic Constraints & Validation**:
    * **Bass/Piano dependencies**: Automatically locks the Bass role if no Piano is assigned (Acoustic measurement).
    * **Availability Filtering**: Dropdowns strictly filter for available members for that specific week.
    * **Validation & Highlighting**: Highlights duplicate assignments, MD not in band, and Bass without Piano for easy correction.
* **State Management**:
    * **Save/Load State**: Save your current roster state to a file and reload it later to continue editing.
* **Visual Dashboard**: Real-time dashboard shows all members, their roles, availability, assignment status, and serving load.
* **Theming**: Distinct **Light** and **Dark** modes with visual cues for disabled fields.
* **Exports**:
    * **Excel**: Clean table export without redundant columns.
    * **Image**: Beautifully rendered PNG with category headers and alignment (requires Pillow).

## Prerequisites

*   `uv`

## Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/chenjingheng0607/timetable-gui.git
    cd timetable-gui
    ```

2.  **Install dependencies:**
    ```bash
    uv sync
    ```

## Usage

### Running the Application

To start the application, run the `main.py` script:

```bash
uv run main.py
```

### Workflow

1. **Load Excel**: Click **"Load Excel"** to select your source data file.
2. **Review & Edit**:
    * The **Input Grid** (top) shows the generated roster. You can manually change assignments using the dropdowns.
    * The **Dashboard** (bottom) updates in real-time to show who is serving, their total load, and availability.
    * Validation highlights issues (e.g., duplicate assignments, MD not in band, Bass without Piano) for correction.
3. **Save/Load State**:
    * Use **"Save State"** to save your current roster and assignments to a file.
    * Use **"Load State"** to reload a previously saved roster state and continue editing.
4. **Smart Logic**:
    * To assign an MD, select a band member (Piano, Bass, Guitar). If they are eligible, they will appear in the MD dropdown options.
    * Selecting an MD in the MD row will tag their name in their instrument row with `(MD)`.
5. **Exports**:
    * **"Export Excel"**: Saves the roster data (excluding the internal MD column).
    * **"Export Image"**: Generates a polished PNG of the roster (requires Pillow).

## Configuration

Configuration is handled in `config.py` (not main.py). You can adjust:

*   **ROLES_ORDER**: The display order of columns.
*   **BAND_ROLES**: Which instruments are eligible to be MD.
*   **CATEGORY_CONFIG**: Colors and grouping for teams (Praise & Worship, Production, etc.).
*   **THEMES**: Color palettes for Light and Dark modes.
*   **CLEANUP_OPTIONS**: Fixed options for cleanup roles.
*   **INSTRUMENT_MAP**: Mapping for instrument codes in Excel files.
