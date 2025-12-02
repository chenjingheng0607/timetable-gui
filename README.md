# Auto-Roster Pro (Timetable GUI)

Auto-Roster Pro is a desktop application designed to automate and manage the scheduling of rosters for teams, specifically tailored for church worship teams or similar volunteer groups. It provides a visual interface to generate, edit, and export rosters while managing member availability and preventing burnout.

## Features

*   **Automated Drafting**: Automatically generates a draft roster based on member availability and past usage to minimize burnout.
*   **Visual Dashboard**: A comprehensive view of all members, their roles, availability, and current assignment status.
*   **Interactive Grid**: Easy-to-use grid for manual adjustments with smart dropdowns that only show available members for a specific role and week.
*   **Conflict Detection**: Highlights duplicate assignments to prevent scheduling errors.
*   **Excel Integration**: Seamlessly loads member data from Excel and exports the final roster to Excel.
*   **Image Export**: Generate a shareable PNG image of the roster and dashboard for easy distribution.
*   **Role Management**: Supports various roles including Lead, Vocal, Piano, Drums, Bass, Guitar, PPT, Sound, MC, Ushers, and Cleanup crews.

## Prerequisites

*   Python 3.13 or higher
*   `uv` (recommended for dependency management) or `pip`

## Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/chenjingheng0607/timetable-gui.git
    cd timetable-gui
    ```

2.  **Install dependencies:**
    Using `uv` (Recommended):
    ```bash
    uv sync
    ```
    
    Or using `pip`:
    ```bash
    pip install -r requirements.txt
    ```
    *(Note: If `requirements.txt` is missing, install manually: `pip install pandas openpyxl pyinstaller Pillow`)*

## Usage

### Running the Application

To start the application, run the `main.py` script:

```bash
uv run main.py
# or
python main.py
```

### Workflow

1.  **Load Excel**: Click the **"1. Load Excel"** button to select your source data file.
2.  **Review & Edit**:
    *   The **Input Grid** (top) shows the generated roster. You can manually change assignments using the dropdowns.
    *   The **Dashboard** (bottom) updates in real-time to show who is serving, their total load, and availability.
3.  **Clear Grid**: Use **"2. Clear Grid"** if you want to start fresh or reset selections.
4.  **Export Excel**: Click **"3. Export Excel"** to save the finalized roster as a new Excel file.
5.  **Export Image**: Click **"4. Export Image"** to save a visual snapshot of the roster and dashboard as a PNG file.

### Input File Format

The application expects an Excel file (`.xlsx`) with specific columns:

*   **Name**: The name of the member.
*   **Instrument/Role Columns**: Columns identifying capabilities (e.g., "Instrument", "Piano", "Drum").
*   **Availability Columns**: Columns representing weeks (must contain "Week" in the header).
    *   Values should be "O" (Available) or "N/A" (Not Available).
*   **Group Columns** (Optional): Columns like "FWT" (Worship), "FPH" (Production), "FMC" (MC), "FUT" (Usher) to help auto-detect roles.
*   **Filled/Active**: A column containing "Filled" or "✅" to mark active members.

## Executable

The standalone `.exe` file for Windows is located in the `root` folder.

## Configuration

The application allows some configuration within the code (`main.py`) for:
*   **Roles Order**: The order in which roles appear in the grid.
*   **Category Colors**: Colors used for different teams (Praise & Worship, Production, etc.).
*   **Cleanup Options**: List of groups available for cleanup duties.
