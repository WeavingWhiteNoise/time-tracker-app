"""Database models for Time Tracker application."""
import os
import shutil
import sqlite3
import sys
from datetime import datetime
from pathlib import Path


class Database:
    """SQLite database manager for time tracking."""

    @staticmethod
    def _resolve_db_path(db_path: str) -> str:
        """Return a single shared DB path for both script and EXE runs.

        If no central DB exists yet, copy the newest existing local DB
        (project root or dist folder) as initial seed.
        """
        input_path = Path(db_path)
        if input_path.is_absolute():
            return str(input_path)

        appdata_root = Path(os.getenv("APPDATA", str(Path.home())))
        target_dir = appdata_root / "TimeTracker"
        target_dir.mkdir(parents=True, exist_ok=True)
        target_path = target_dir / db_path

        if not target_path.exists():
            candidates = [Path.cwd() / db_path, Path.cwd() / "dist" / db_path]
            existing = [p for p in candidates if p.exists()]
            if existing:
                newest = max(existing, key=lambda p: p.stat().st_mtime)
                shutil.copy2(newest, target_path)

        return str(target_path)
    
    def __init__(self, db_path: str = "time_tracker.db"):
        """Initialize database connection."""
        self.db_path = self._resolve_db_path(db_path)
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller EXE (dist\TimeTracker.exe) → go one level up
            local_dir = Path(sys.executable).parent.parent
        else:
            local_dir = Path.cwd()
        self.local_path = local_dir / Path(db_path).name
        self.init_database()
    
    def init_database(self):
        """Create tables if they don't exist."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Projects table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                recipient TEXT,
                service_type TEXT,
                process TEXT,
                workplace TEXT,
                cost_center TEXT,
                project_manager TEXT,
                working_place TEXT,
                is_favorite INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Backward-compatible migration for existing databases.
        cursor.execute("PRAGMA table_info(projects)")
        project_columns = [column[1] for column in cursor.fetchall()]
        if "is_favorite" not in project_columns:
            cursor.execute("ALTER TABLE projects ADD COLUMN is_favorite INTEGER DEFAULT 0")
        if "recipient" not in project_columns:
            cursor.execute("ALTER TABLE projects ADD COLUMN recipient TEXT")
        if "service_type" not in project_columns:
            cursor.execute("ALTER TABLE projects ADD COLUMN service_type TEXT")
        if "process" not in project_columns:
            cursor.execute("ALTER TABLE projects ADD COLUMN process TEXT")
        if "workplace" not in project_columns:
            cursor.execute("ALTER TABLE projects ADD COLUMN workplace TEXT")
        
        # Time entries table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS time_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                start_time TIMESTAMP NOT NULL,
                end_time TIMESTAMP,
                duration_minutes INTEGER,
                service_type TEXT,
                workplace TEXT,
                notes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (project_id) REFERENCES projects(id)
            )
        """)

        # Add missing columns to time_entries if they don't exist
        cursor.execute("PRAGMA table_info(time_entries)")
        time_entries_columns = [column[1] for column in cursor.fetchall()]
        if "service_type" not in time_entries_columns:
            cursor.execute("ALTER TABLE time_entries ADD COLUMN service_type TEXT")
        if "workplace" not in time_entries_columns:
            cursor.execute("ALTER TABLE time_entries ADD COLUMN workplace TEXT")
        
        conn.commit()
        conn.close()
        self._sync_to_local()
    
    def _sync_to_local(self):
        """Mirror the AppData DB to the local project folder after every write."""
        try:
            if str(self.local_path.resolve()) != str(Path(self.db_path).resolve()):
                shutil.copy2(self.db_path, self.local_path)
        except OSError:
            pass

    def add_project(
        self,
        name: str,
        recipient: str = "",
        process: str = "",
    ) -> int:
        """Add a new project."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute("""
                INSERT INTO projects (name, recipient, process)
                VALUES (?, ?, ?)
            """, (name, recipient, process))
            conn.commit()
            project_id = cursor.lastrowid
            self._sync_to_local()
            return project_id
        finally:
            conn.close()

    def update_project(
        self,
        project_id: int,
        name: str,
        recipient: str = "",
        process: str = "",
    ) -> bool:
        """Update an existing project."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute(
            """
            UPDATE projects
            SET name = ?, recipient = ?, process = ?
            WHERE id = ?
            """,
            (name, recipient, process, project_id),
        )
        conn.commit()
        success = cursor.rowcount > 0
        conn.close()
        self._sync_to_local()
        return success
    
    def get_all_projects(self, filter_text: str = "") -> list:
        """Get all projects, optionally filtered by name or recipient.

        Favorites are ordered first, then by project name.
        """
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        if filter_text:
            cursor.execute(
                """
                SELECT id, name, recipient, service_type, process, workplace, is_favorite
                FROM projects
                WHERE name LIKE ? OR recipient LIKE ?
                ORDER BY is_favorite DESC, name COLLATE NOCASE ASC
                """,
                (f"%{filter_text}%", f"%{filter_text}%"),
            )
        else:
            cursor.execute(
                """
                SELECT id, name, recipient, service_type, process, workplace, is_favorite
                FROM projects
                ORDER BY is_favorite DESC, name COLLATE NOCASE ASC
                """
            )

        projects = cursor.fetchall()
        conn.close()
        return projects

    def get_project_by_id(self, project_id: int) -> tuple:
        """Get project by ID."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT id, name, recipient, service_type, process, workplace, is_favorite
            FROM projects
            WHERE id = ?
            """,
            (project_id,),
        )
        project = cursor.fetchone()
        conn.close()
        return project

    def set_project_favorite(self, project_id: int, is_favorite: bool) -> bool:
        """Set or unset favorite flag for a project."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE projects SET is_favorite = ? WHERE id = ?",
            (1 if is_favorite else 0, project_id),
        )
        conn.commit()
        success = cursor.rowcount > 0
        conn.close()
        self._sync_to_local()
        return success
    
    def start_time_entry(self, project_id: int, start_time: str, notes: str = "", 
                         service_type: str = "", workplace: str = "") -> int:
        """Start a new time entry."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("""
            INSERT INTO time_entries (project_id, start_time, notes, service_type, workplace)
            VALUES (?, ?, ?, ?, ?)
        """, (project_id, start_time, notes, service_type, workplace))
        conn.commit()
        entry_id = cursor.lastrowid
        conn.close()
        self._sync_to_local()
        return entry_id
    
    def end_time_entry(self, entry_id: int, end_time: str, notes: str = None) -> bool:
        """End a time entry, calculate duration, and optionally update notes."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Get start time
        cursor.execute("SELECT start_time FROM time_entries WHERE id = ?", (entry_id,))
        result = cursor.fetchone()
        
        if not result:
            conn.close()
            return False
        
        start_time = datetime.fromisoformat(result[0])
        end_time_dt = datetime.fromisoformat(end_time)
        duration = int((end_time_dt - start_time).total_seconds() / 60)
        
        if notes is None:
            cursor.execute("""
                UPDATE time_entries 
                SET end_time = ?, duration_minutes = ?
                WHERE id = ?
            """, (end_time, duration, entry_id))
        else:
            cursor.execute("""
                UPDATE time_entries 
                SET end_time = ?, duration_minutes = ?, notes = ?
                WHERE id = ?
            """, (end_time, duration, notes, entry_id))
        
        conn.commit()
        conn.close()
        self._sync_to_local()
        return True
    
    def get_time_entries_by_date(self, project_id: int = None, 
                                date_str: str = None) -> list:
        """Get time entries for a specific date (or all if no date specified)."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        if date_str:
            query = """
                SELECT te.id, te.project_id, p.name, te.start_time, te.end_time, 
                       te.duration_minutes, p.recipient, te.service_type, p.process, te.workplace, te.notes
                FROM time_entries te
                JOIN projects p ON te.project_id = p.id
                WHERE DATE(te.start_time) = ?
            """
            params = [date_str]
            
            if project_id:
                query += " AND te.project_id = ?"
                params.append(project_id)
        else:
            query = """
                SELECT te.id, te.project_id, p.name, te.start_time, te.end_time, 
                       te.duration_minutes, p.recipient, te.service_type, p.process, te.workplace, te.notes
                FROM time_entries te
                JOIN projects p ON te.project_id = p.id
            """
            params = []
            
            if project_id:
                query += " WHERE te.project_id = ?"
                params.append(project_id)
        
        query += " ORDER BY te.start_time DESC"
        cursor.execute(query, params)
        entries = cursor.fetchall()
        conn.close()
        return entries

    def get_time_entry_by_id(self, entry_id: int) -> tuple:
        """Get one time entry with joined project metadata."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT te.id, te.project_id, p.name, te.start_time, te.end_time,
                   te.duration_minutes, p.recipient, te.service_type, p.process, te.workplace, te.notes
            FROM time_entries te
            JOIN projects p ON te.project_id = p.id
            WHERE te.id = ?
            """,
            (entry_id,),
        )
        entry = cursor.fetchone()
        conn.close()
        return entry

    def update_time_entry(
        self,
        entry_id: int,
        project_id: int,
        start_time: str,
        end_time: str,
        service_type: str,
        workplace: str,
        notes: str,
    ) -> bool:
        """Update an existing time entry and recalculate duration."""
        start_dt = datetime.fromisoformat(start_time)
        end_dt = datetime.fromisoformat(end_time)
        duration = int((end_dt - start_dt).total_seconds() / 60)

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            """
            UPDATE time_entries
            SET project_id = ?,
                start_time = ?,
                end_time = ?,
                duration_minutes = ?,
                service_type = ?,
                workplace = ?,
                notes = ?
            WHERE id = ?
            """,
            (
                project_id,
                start_time,
                end_time,
                duration,
                service_type,
                workplace,
                notes,
                entry_id,
            ),
        )
        conn.commit()
        success = cursor.rowcount > 0
        conn.close()
        self._sync_to_local()
        return success
    
    def delete_time_entry(self, entry_id: int) -> bool:
        """Delete a time entry."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM time_entries WHERE id = ?", (entry_id,))
        conn.commit()
        success = cursor.rowcount > 0
        conn.close()
        self._sync_to_local()
        return success

    def get_time_entries_by_month(self, date_str: str, project_id: int = None) -> list:
        """Get all time entries for a specific month.
        
        Args:
            date_str: Date string in format 'yyyy-MM-dd' (any day in the desired month)
            project_id: Optional project filter
            
        Returns:
            List of time entries for the entire month
        """
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        query = """
            SELECT te.id, te.project_id, p.name, te.start_time, te.end_time, 
                   te.duration_minutes, p.recipient, te.service_type, p.process, te.workplace, te.notes
            FROM time_entries te
            JOIN projects p ON te.project_id = p.id
            WHERE strftime('%Y-%m', te.start_time) = strftime('%Y-%m', ?)
        """
        params = [date_str]
        
        if project_id:
            query += " AND te.project_id = ?"
            params.append(project_id)
        
        query += " ORDER BY te.start_time ASC"
        cursor.execute(query, params)
        entries = cursor.fetchall()
        conn.close()
        return entries
