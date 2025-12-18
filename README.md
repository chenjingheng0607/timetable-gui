# Auto-Roster (Timetable GUI)

Auto-Roster is a desktop application designed to automate and manage the scheduling of rosters for teams, specifically tailored for church worship teams or similar volunteer groups. It provides a visual interface to generate, edit, and export rosters while managing member availability and preventing burnout.

## Features

*   **Automated Drafting**: Automatically generates a draft roster based on member availability, past usage, and role capabilities to minimize burnout.
*   **Smart MD Handling**: 
    *   Automatically identifies MDs based on the selected band members.
    *   Visually tags the active MD with `(MD)` in the grid.
    *   Excludes the redundant "MD" column from exports to keep the output clean.
*   **Logic Constraints**:
    *   **Bass/Piano dependencies**: Automatically locks the Bass role if no Piano is assigned (Acoustic measurement).
    *   **Availability Filtering**: Dropdowns strictly filter for available members for that specific week.
*   **Visual Dashboard**: A comprehensive view of all members, their roles, availability, and current assignment status.
*   **Theming**: distinct **Light** and **Dark** modes with proper visual cues for disabled fields.
*   **Exports**: 
    *   **Excel**: Clean table export without redundant columns.
    *   **Image**: Beautifully rendered PNG with category headers and alignment.

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

1.  **Load Excel**: Click the **"1. Load Excel"** button to select your source data file.
2.  **Review & Edit**:
    *   The **Input Grid** (top) shows the generated roster. You can manually change assignments using the dropdowns.
    *   The **Dashboard** (bottom) updates in real-time to show who is serving, their total load, and availability.
3.  **Smart Logic**:
    *   To assign an MD, select a band member (Piano, Bass, Guitar). If they are eligible, they will appear in the MD dropdown options.
    *   Selecting an MD in the MD row will tag their name in their instrument row with `(MD)`.
4.  **Exports**:
    *   **"3. Export Excel"**: Saves the roster data (excluding the internal MD column).
    *   **"4. Export Image"**: Generates a polished PNG of the roster.

## Configuration

Configuration is handled in `config.py` (not main.py). You can adjust:

*   **ROLES_ORDER**: The display order of columns.
*   **BAND_ROLES**: Which instruments are eligible to be MD.
*   **CATEGORY_CONFIG**: Colors and grouping for teams (Praise & Worship, Production, etc.).
*   **THEMES**: Color palettes for Light and Dark modes.

## Building Executable

To build a standalone `.exe` using `nuitka`:

```bash
nuitka --onefile --enable-plugin=pyside6 --windows-console-mode=disable --windows-icon-from-ico=FirelightLogo.png --lto=yes --remove-output --include-qt-plugins=platforms,styles --nofollow-import-to=PySide6.QtNetwork --nofollow-import-to=PySide6.QtSql main.py
