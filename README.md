# Time Tracker Application

A simple yet powerful desktop application for tracking work time across multiple projects with detailed metadata and Excel export functionality.

## Features

### Core Functionality
- **Project Management**: Create multiple projects with associated metadata
  - Project names
  - Cost centers
  - Project managers
  - Working places

### Time Tracking
- **Live Timer**: Start and stop timers for active work sessions
- **Manual Entries**: Add time entries with custom start and end times
- **Today's View**: Display all time entries for the current day

### Metadata & Reporting
- **Project Metadata**: Store cost center, project manager, and working place for each project
- **Excel Export**: Export daily time logs to Excel files with:
  - Detailed time entries
  - Duration calculations (in minutes and hours)
  - Summary statistics (total time, number of entries)

## Installation

### Requirements
- Python 3.8 or later
- pip (Python package manager)

### Setup

1. Clone or navigate to the project directory:
```bash
cd Time-Tracker
```

2. Create a virtual environment:
```bash
python -m venv venv
```

3. Activate the virtual environment:
   - **Windows**:
   ```bash
   venv\Scripts\activate
   ```
   - **MacOS/Linux**:
   ```bash
   source venv/bin/activate
   ```

4. Install dependencies:
```bash
pip install -r requirements.txt
```

## Running the Application

### Option 1: Using the Executable (Easiest) ⭐

**Desktop Shortcut (Recommended)**
- A shortcut named **"Time Tracker"** has been created on your desktop
- Simply double-click it to launch the application
- No terminal or Python knowledge required!

**Or run the batch file:**
- Double-click `run_timetracker.bat` in the project folder

### Option 2: From Command Line

```bash
venv\Scripts\python main.py
```

### Option 3: Direct Executable

Navigate to the `dist` folder and double-click `TimeTracker.exe`

## Usage

### Adding a Project
1. Click the **"+ New Project"** button
2. Enter project details:
   - Project Name (required)
   - Cost Center
   - Project Manager
   - Working Place
3. Click **OK**

### Tracking Time

#### Using the Timer
1. Select a project from the dropdown
2. Click **"Start Timer"** to begin tracking
3. The active timer will be displayed and updated every second
4. Click **"Stop Timer"** to end the session

#### Adding Manual Entries
1. Select a project
2. Set start and end times using the time pickers
3. Click **"Add Manual Entry"**

### Exporting Data

#### Export Today's Entries
1. Click **"Export to Excel"**
2. A file named `time_tracker_YYYY-MM-DD.xlsx` will be created
3. The Excel file contains:
   - **Time Entries sheet**: Detailed list of all entries with project metadata
   - **Summary sheet**: Total duration, number of entries, export timestamp

#### End Day
1. Click **"End Day"**
2. Choose to export entries or skip export
3. The application will confirm the end of your work day

## Data Storage

- All data is stored locally in `time_tracker.db` (SQLite database)
- No cloud synchronization or remote storage
- Database is created automatically on first run
- Excel files are exported to the current working directory

## File Structure

```
Time-Tracker/
├── main.py                 # Application entry point
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── time_tracker.db        # SQLite database (created on first run)
└── src/
    ├── models/
    │   └── database.py    # Database management
    ├── ui/
    │   └── main_window.py # PyQt5 UI
    └── utils/
        └── export.py      # Excel export functionality
```

## Troubleshooting

### "No module named 'PyQt5'"
Make sure you've activated your virtual environment and installed dependencies:
```bash
pip install -r requirements.txt
```

### Excel export not working
Ensure the working directory has write permissions for creating the .xlsx file.

### Timer not updating
Try restarting the application.

## Future Enhancements

Potential features for future versions:
- Edit existing time entries
- Multiple date filtering and reporting
- Project statistics and visualizations
- Recurring projects
- Time entry notes/descriptions
- Import from CSV
- Backup and restore functionality

## License

This is a personal project for time tracking purposes.
