<!-- Use this file to provide workspace-specific custom instructions to Copilot. -->

# Time Tracker Application

This is a PyQt5-based desktop application for tracking work time across multiple projects with metadata and Excel export functionality.

## Project Structure
- `src/models/database.py` - SQLite database management
- `src/ui/main_window.py` - PyQt5 UI components
- `src/utils/export.py` - Excel export functionality
- `main.py` - Application entry point
- `requirements.txt` - Python dependencies

## Setup Instructions

1. Install Python 3.8 or later
2. Create and activate a virtual environment:
   ```
   python -m venv venv
   venv\Scripts\activate
   ```
3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
4. Run the application:
   ```
   python main.py
   ```

## Features
- Create and manage work projects with metadata (cost center, project manager, working place)
- Start/stop timers for active work sessions
- Add manual time entries with custom start/end times
- View today's time entries
- Export daily time logs to Excel with summary data
- SQLite database for local data storage
