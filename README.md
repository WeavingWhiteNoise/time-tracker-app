# Time Tracker Application

Desktop app (PyQt5 + SQLite) for tracking working time across projects and exporting logs to Excel.

## Current Features

### Project Management
- Add and edit projects
- Optional project metadata:
   - Recipient (Empfaenger)
   - Process (Vorgang)
- Mark projects as favorites and keep favorites sorted to the top
- Filter projects by name or recipient

### Time Tracking
- Start/stop live timer for the selected project
- Add manual entries with selected date + start/end time
- Optional metadata per entry:
   - Service type (Leistungsart code)
   - Workplace (Arbeitsplatz code)
   - Notes/comment
- Double-click an existing row to load project/comment and start a new timer
- Edit existing finished entries
- Delete entries

### Entry View
- Date picker to view entries for the selected day
- Table columns:
   - Zeiten von bis
   - Projekt Name
   - Kommentar

### Excel Export
- "Export to Excel" exports entries to an .xlsx file:
   - Sheet 1: Time Entries
   - Sheet 2: Summary
- Default filename:
   - time_tracker_YYYY-MM-DD.xlsx

## Installation

### Requirements
- Python 3.8+
- pip

Note:
- The pinned versions in requirements.txt are older. If you use very new Python versions, dependency resolution may fail. In that case, use a project venv and install compatible package versions.

### Setup

1. Open a terminal in the project folder.
2. Create a virtual environment:

```bash
python -m venv .venv
```

3. Activate it:

Windows (PowerShell):

```powershell
.\.venv\Scripts\Activate.ps1
```

Windows (cmd):

```bat
.venv\Scripts\activate.bat
```

4. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

### From Source

```bash
python main.py
```

### From Existing EXE
- Run dist/TimeTracker.exe

### Convenience Script
- Run run_timetracker.bat

## Build EXE (PyInstaller)

```bash
pyinstaller --noconfirm TimeTracker.spec
```

Output:
- dist/TimeTracker.exe

## Data Storage

- Primary database location:
   - %APPDATA%/TimeTracker/time_tracker.db
- On startup, the app can seed this DB from local copies if needed.
- The app mirrors DB changes back to a local project copy for convenience.

## Project Structure

```
time-tracker-app/
├── main.py
├── TimeTracker.spec
├── requirements.txt
├── run_timetracker.bat
├── src/
│   ├── models/
│   │   └── database.py
│   ├── ui/
│   │   └── main_window.py
│   └── utils/
│       └── export.py
└── dist/
      └── TimeTracker.exe
```

## Notes

- The former "End Day" button/workflow is removed.
- Export currently uses all entries from the database.

## License

Personal project.
