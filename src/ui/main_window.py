"""Main application window and UI."""
import sys
from datetime import datetime
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QComboBox, QPushButton, QLabel, QTimeEdit, QDateEdit,
    QListWidget, QListWidgetItem, QMessageBox, QFileDialog,
    QDialog, QLineEdit, QFormLayout, QSpinBox, QTextEdit,
    QCheckBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QDateTimeEdit
)
from PyQt5.QtCore import Qt, QDateTime, QDate, QTime, QTimer
from PyQt5.QtGui import QFont, QColor, QStandardItemModel, QStandardItem

from src.models.database import Database
from src.utils.export import export_to_excel


class AddProjectDialog(QDialog):
    """Dialog for adding a new project."""
    
    def __init__(self, parent=None, initial_data=None, dialog_title="Add New Project"):
        """Initialize the dialog."""
        super().__init__(parent)
        self.setWindowTitle(dialog_title)
        self.setGeometry(100, 100, 400, 200)
        
        layout = QFormLayout()
        
        self.project_name = QLineEdit()
        self.recipient = QLineEdit()
        self.process = QLineEdit()
        
        layout.addRow("Projekt Name:", self.project_name)
        layout.addRow("Empfaenger:", self.recipient)
        layout.addRow("Vorgang:", self.process)
        
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Cancel")
        
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        
        layout.addRow(button_layout)
        self.setLayout(layout)

        if initial_data:
            self.project_name.setText(initial_data.get("name", ""))
            self.recipient.setText(initial_data.get("recipient", ""))
            self.process.setText(initial_data.get("process", ""))
    
    def get_data(self):
        """Return the entered data."""
        return {
            'name': self.project_name.text(),
            'recipient': self.recipient.text(),
            'process': self.process.text()
        }


class EditEntryDialog(QDialog):
    """Dialog for editing an existing time entry and its metadata."""

    SERVICE_TYPES = [
        "114000 | Projektmanagement",
        "114009 | Reisezeit PM",
        "114012 | Software Engineering",
        "114013 | IBN Software ext.",
        "114019 | Reisezeit ENG",
        "114023 | Montage Elektr. int.",
        "114026 | Montage Elektr. ext.",
        "120003 | Std. Systemingenieur",
        "120200 | Reisezeit",
    ]

    WORKPLACES = [
        "D330000 | 4180 Software Engineering TAS DA",
        "D330010 | 4180 Software Software Applikation TAS",
        "D330020 | 4180 Software Engineering TAS NH",
        "P230000 | 4180 Montage TAS NH",
    ]

    def __init__(self, parent, db, entry_data):
        super().__init__(parent)
        self.db = db
        self.entry_data = entry_data
        self.setWindowTitle("Edit Entry")
        self.setGeometry(120, 120, 650, 420)

        (
            _entry_id,
            project_id,
            _name,
            start_time,
            end_time,
            _duration,
            _recipient,
            service_type,
            _process,
            workplace,
            notes,
        ) = entry_data

        layout = QFormLayout()

        self.project_combo = QComboBox()
        for project in self.db.get_all_projects(""):
            p_id, p_name, p_recipient, _p_service_type, p_process, _p_workplace, p_fav = project
            star = "★ " if p_fav else ""
            display = ", ".join([part for part in [star + (p_name or ""), p_recipient or "", p_process or ""] if part])
            self.project_combo.addItem(display, p_id)
        self._set_project(project_id)

        self.start_datetime = QDateTimeEdit()
        self.start_datetime.setCalendarPopup(True)
        self.start_datetime.setDisplayFormat("dd.MM.yyyy HH:mm")

        start_dt = datetime.fromisoformat(start_time)
        self.start_datetime.setDateTime(
            QDateTime(
                QDate(start_dt.year, start_dt.month, start_dt.day),
                QTime(start_dt.hour, start_dt.minute, start_dt.second),
            )
        )

        self.end_datetime = QDateTimeEdit()
        self.end_datetime.setCalendarPopup(True)
        self.end_datetime.setDisplayFormat("dd.MM.yyyy HH:mm")
        if end_time:
            end_dt = datetime.fromisoformat(end_time)
            self.end_datetime.setDateTime(
                QDateTime(
                    QDate(end_dt.year, end_dt.month, end_dt.day),
                    QTime(end_dt.hour, end_dt.minute, end_dt.second),
                )
            )
        else:
            self.end_datetime.setDateTime(self.start_datetime.dateTime().addSecs(3600))

        self.leistungsart_combo = QComboBox()
        self.leistungsart_combo.addItems(self.SERVICE_TYPES)
        self._set_combo_code(self.leistungsart_combo, service_type or "")

        self.arbeitsplatz_combo = QComboBox()
        self.arbeitsplatz_combo.addItems(self.WORKPLACES)
        self._set_combo_code(self.arbeitsplatz_combo, workplace or "")

        self.comment_edit = QTextEdit()
        self.comment_edit.setMaximumHeight(90)
        self.comment_edit.setPlainText(notes or "")

        layout.addRow("Projekt:", self.project_combo)
        layout.addRow("Start:", self.start_datetime)
        layout.addRow("Ende:", self.end_datetime)
        layout.addRow("Leistungsart:", self.leistungsart_combo)
        layout.addRow("Arbeitsplatz:", self.arbeitsplatz_combo)
        layout.addRow("Kommentar:", self.comment_edit)

        button_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        save_btn.clicked.connect(self._on_save)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        layout.addRow(button_layout)

        self.setLayout(layout)

    def _set_combo_code(self, combo: QComboBox, code: str):
        """Select first combo item matching a metadata code."""
        for idx in range(combo.count()):
            if combo.itemText(idx).startswith(f"{code} |"):
                combo.setCurrentIndex(idx)
                return

    def _set_project(self, project_id: int):
        """Select a project by ID in the dialog combo."""
        for idx in range(self.project_combo.count()):
            if self.project_combo.itemData(idx) == project_id:
                self.project_combo.setCurrentIndex(idx)
                return

    def _on_save(self):
        """Validate data before closing dialog."""
        start_dt = self.start_datetime.dateTime().toPyDateTime()
        end_dt = self.end_datetime.dateTime().toPyDateTime()
        if end_dt <= start_dt:
            QMessageBox.warning(self, "Warning", "End time must be after start time.")
            return
        self.accept()

    def get_data(self):
        """Return edited entry data from the dialog."""
        return {
            "project_id": self.project_combo.currentData(),
            "start_time": self.start_datetime.dateTime().toPyDateTime().isoformat(),
            "end_time": self.end_datetime.dateTime().toPyDateTime().isoformat(),
            "service_type": self.leistungsart_combo.currentText().split(" | ")[0],
            "workplace": self.arbeitsplatz_combo.currentText().split(" | ")[0],
            "notes": self.comment_edit.toPlainText(),
        }


class TimeTrackerWindow(QMainWindow):
    """Main application window."""
    
    def __init__(self):
        """Initialize the main window."""
        super().__init__()
        self.setWindowTitle("Time Tracker")
        self.setGeometry(100, 100, 800, 600)
        
        # Initialize database
        self.db = Database("time_tracker.db")
        self.current_entry = None  # Currently running timer entry
        self._updating_projects = False
        
        # Setup UI
        self.setup_ui()
        
        # Timer for live updates
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_running_entry)
    
    def setup_ui(self):
        """Setup the user interface."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        
        # Title
        title = QLabel("Time Tracker")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        main_layout.addWidget(title)
        
        # Formatted controls section with aligned inputs
        controls_layout = QVBoxLayout()
        controls_layout.setSpacing(8)
        
        # Row 1: Filter
        filter_row = QHBoxLayout()
        filter_label = QLabel("Filter:")
        filter_label.setFixedWidth(120)
        filter_row.addWidget(filter_label)
        self.project_filter = QLineEdit()
        self.project_filter.setPlaceholderText("Filter...")
        self.project_filter.setFixedWidth(300)
        self.project_filter.textChanged.connect(self.on_project_filter_changed)
        filter_row.addWidget(self.project_filter)
        filter_row.addStretch()
        controls_layout.addLayout(filter_row)
        
        # Row 2: Project dropdown
        project_row = QHBoxLayout()
        project_label = QLabel("Projekt:")
        project_label.setFixedWidth(120)
        project_row.addWidget(project_label)
        self.project_combo = QComboBox()
        self.project_combo.setModel(QStandardItemModel(self.project_combo))
        self.project_combo.setFixedWidth(300)
        self.project_combo.currentIndexChanged.connect(self.on_project_combo_changed)
        project_row.addWidget(self.project_combo)

        self.favorite_checkbox = QCheckBox("★ Favorite")
        self.favorite_checkbox.clicked.connect(self.on_favorite_checkbox_clicked)
        project_row.addWidget(self.favorite_checkbox)

        add_project_btn = QPushButton("+ New Project")
        add_project_btn.clicked.connect(self.add_new_project)
        project_row.addWidget(add_project_btn)

        edit_project_btn = QPushButton("Edit Project")
        edit_project_btn.clicked.connect(self.edit_selected_project)
        project_row.addWidget(edit_project_btn)

        project_row.addStretch()
        controls_layout.addLayout(project_row)
        
        # Row 3: Leistungsart dropdown
        leistungsart_row = QHBoxLayout()
        leistungsart_label = QLabel("Leistungsart:")
        leistungsart_label.setFixedWidth(120)
        leistungsart_row.addWidget(leistungsart_label)
        self.leistungsart_combo = QComboBox()
        self.leistungsart_combo.addItems([
            "114000 | Projektmanagement",
            "114009 | Reisezeit PM",
            "114012 | Software Engineering",
            "114013 | IBN Software ext.",
            "114019 | Reisezeit ENG",
            "114023 | Montage Elektr. int.",
            "114026 | Montage Elektr. ext.",
            "120003 | Std. Systemingenieur",
            "120200 | Reisezeit"
        ])
        self.leistungsart_combo.setFixedWidth(300)
        leistungsart_row.addWidget(self.leistungsart_combo)
        leistungsart_row.addStretch()
        controls_layout.addLayout(leistungsart_row)
        
        # Row 4: Arbeitsplatz dropdown
        arbeitsplatz_row = QHBoxLayout()
        arbeitsplatz_label = QLabel("Arbeitsplatz:")
        arbeitsplatz_label.setFixedWidth(120)
        arbeitsplatz_row.addWidget(arbeitsplatz_label)
        self.arbeitsplatz_combo = QComboBox()
        self.arbeitsplatz_combo.addItems([
            "D330000 | 4180 Software Engineering TAS DA",
            "D330010 | 4180 Software Software Applikation TAS",
            "D330020 | 4180 Software Engineering TAS NH",
            "P230000 | 4180 Montage TAS NH"
        ])
        self.arbeitsplatz_combo.setFixedWidth(300)
        arbeitsplatz_row.addWidget(self.arbeitsplatz_combo)
        arbeitsplatz_row.addStretch()
        controls_layout.addLayout(arbeitsplatz_row)
        
        # Row 5: Start Time
        start_time_row = QHBoxLayout()
        start_time_label = QLabel("Start Time:")
        start_time_label.setFixedWidth(120)
        start_time_row.addWidget(start_time_label)
        self.start_time = QTimeEdit()
        self.start_time.setDateTime(QDateTime.currentDateTime())
        self.start_time.setFixedWidth(300)
        start_time_row.addWidget(self.start_time)
        start_time_row.addStretch()
        controls_layout.addLayout(start_time_row)
        
        # Row 6: End Time
        end_time_row = QHBoxLayout()
        end_time_label = QLabel("End Time:")
        end_time_label.setFixedWidth(120)
        end_time_row.addWidget(end_time_label)
        self.end_time = QTimeEdit()
        self.end_time.setDateTime(QDateTime.currentDateTime())
        self.end_time.setFixedWidth(300)
        end_time_row.addWidget(self.end_time)
        end_time_row.addStretch()
        controls_layout.addLayout(end_time_row)
        
        # Row 7: Date selector
        date_row = QHBoxLayout()
        date_label = QLabel("Date:")
        date_label.setFixedWidth(120)
        date_row.addWidget(date_label)
        self.date_edit = QDateEdit()
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setFixedWidth(300)
        self.date_edit.dateChanged.connect(self.refresh_entries)
        date_row.addWidget(self.date_edit)
        date_row.addStretch()
        controls_layout.addLayout(date_row)
        
        main_layout.addLayout(controls_layout)
        
        # Comment field
        main_layout.addWidget(QLabel("Comment (optional):"))
        self.comment_edit = QTextEdit()
        self.comment_edit.setMaximumHeight(60)
        self.comment_edit.setPlaceholderText("Enter details about your work...")
        main_layout.addWidget(self.comment_edit)
        
        # Action buttons
        action_layout = QHBoxLayout()
        
        start_btn = QPushButton("Start Timer")
        start_btn.clicked.connect(self.start_timer)
        start_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        action_layout.addWidget(start_btn)
        
        stop_btn = QPushButton("Stop Timer")
        stop_btn.clicked.connect(self.stop_timer)
        stop_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold;")
        action_layout.addWidget(stop_btn)
        
        add_manual_btn = QPushButton("Add Manual Entry")
        add_manual_btn.clicked.connect(self.add_manual_entry)
        action_layout.addWidget(add_manual_btn)
        
        main_layout.addLayout(action_layout)
        
        # Current entry display
        self.current_entry_label = QLabel("No active timer")
        self.current_entry_label.setStyleSheet("background-color: #e3f2fd; padding: 10px; border-radius: 5px;")
        main_layout.addWidget(self.current_entry_label)
        
        # Time entries table
        main_layout.addWidget(QLabel("Entries for selected date (double-click row to load & start):"))
        self.entries_table = QTableWidget()
        self.entries_table.setColumnCount(3)
        self.entries_table.setHorizontalHeaderLabels(["Zeiten von bis", "Projekt Name", "Kommentar"])
        self.entries_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.entries_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.entries_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.entries_table.verticalHeader().setVisible(False)
        self.entries_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.entries_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.entries_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.entries_table.itemDoubleClicked.connect(self.on_entry_double_clicked)
        main_layout.addWidget(self.entries_table)
        
        # Export and end day buttons
        end_day_layout = QHBoxLayout()
        
        export_btn = QPushButton("Export to Excel")
        export_btn.clicked.connect(self.export_to_excel)
        export_btn.setStyleSheet("background-color: #2196F3; color: white;")
        end_day_layout.addWidget(export_btn)
        
        end_day_btn = QPushButton("End Day")
        end_day_btn.clicked.connect(self.end_day)
        end_day_btn.setStyleSheet("background-color: #FF9800; color: white;")
        end_day_layout.addWidget(end_day_btn)

        edit_entry_btn = QPushButton("Edit Entry")
        edit_entry_btn.clicked.connect(self.edit_selected_entry)
        edit_entry_btn.setStyleSheet("background-color: #607D8B; color: white;")
        end_day_layout.addWidget(edit_entry_btn)

        delete_entry_btn = QPushButton("Delete Entry")
        delete_entry_btn.clicked.connect(self.delete_selected_entry)
        delete_entry_btn.setStyleSheet("background-color: #9E9E9E; color: white;")
        end_day_layout.addWidget(delete_entry_btn)
        
        main_layout.addLayout(end_day_layout)
        
        central_widget.setLayout(main_layout)
        
        # Load initial data
        self.refresh_projects()
        self.refresh_entries()
    
    def refresh_projects(self, selected_project_id: int = None):
        """Refresh the project list in the combo box."""
        self._updating_projects = True

        if selected_project_id is None:
            selected_project_id = self.project_combo.currentData(Qt.UserRole)

        model = self.project_combo.model()
        model.clear()

        projects = self.db.get_all_projects(self.project_filter.text().strip())
        selected_index = -1

        for index, project in enumerate(projects):
            project_id, name, recipient, service_type, process, workplace, is_favorite = project
            
            # Build display name with optional favorite star and process
            star = "★ " if is_favorite else ""
            parts = [star + name, recipient, process]
            # Filter out empty parts
            display_name = ", ".join([p for p in parts if p])

            item = QStandardItem(display_name)
            item.setData(project_id, Qt.UserRole)
            model.appendRow(item)

            if project_id == selected_project_id:
                selected_index = index

        if selected_index >= 0:
            self.project_combo.setCurrentIndex(selected_index)
        elif model.rowCount() > 0:
            self.project_combo.setCurrentIndex(0)

        self._updating_projects = False
        self.on_project_combo_changed()  # Update checkbox state

    def on_project_filter_changed(self):
        """Filter projects while typing."""
        self.refresh_projects()

    def on_project_combo_changed(self):
        """Update the favorite checkbox when project selection changes."""
        if self._updating_projects:
            return
        
        project_id = self.project_combo.currentData(Qt.UserRole)
        if project_id is None:
            self.favorite_checkbox.setChecked(False)
            return
        
        project = self.db.get_project_by_id(project_id)
        if project:
            is_favorite = project[6]  # is_favorite is the 7th column (index 6)
            self._updating_projects = True
            self.favorite_checkbox.setChecked(is_favorite)
            self._updating_projects = False

    def on_favorite_checkbox_clicked(self):
        """Toggle favorite status for the selected project."""
        if self._updating_projects:
            return
        
        project_id = self.project_combo.currentData(Qt.UserRole)
        if project_id is None:
            return
        
        is_favorite = self.favorite_checkbox.isChecked()
        self.db.set_project_favorite(project_id, is_favorite)
        self.refresh_projects(selected_project_id=project_id)

    def select_project_by_id(self, project_id: int) -> bool:
        """Select a project by ID in the combo box, returns True on success."""
        for i in range(self.project_combo.count()):
            if self.project_combo.itemData(i, Qt.UserRole) == project_id:
                self.project_combo.setCurrentIndex(i)
                return True
        return False
    
    def refresh_entries(self):
        """Refresh the table of entries for the selected date."""
        self.entries_table.setRowCount(0)
        selected_day = self.date_edit.date().toString("yyyy-MM-dd")
        entries = self.db.get_time_entries_by_date(date_str=selected_day)
        
        for row_index, entry in enumerate(entries):
            entry_id, project_id, name, start_time, end_time, duration, recipient, service_type, process, workplace, notes = entry
            # Parse ISO format times (YYYY-MM-DDTHH:MM:SS.ffffff)
            start_hour_min = start_time.split('T')[1][:5] if 'T' in start_time else start_time[:5]

            if end_time:
                end_hour_min = end_time.split('T')[1][:5] if 'T' in end_time else end_time[:5]
                times_text = f"{start_hour_min} - {end_hour_min}"
            else:
                times_text = f"{start_hour_min} - RUNNING"

            self.entries_table.insertRow(row_index)

            times_item = QTableWidgetItem(times_text)
            project_item = QTableWidgetItem(name if name else "")
            notes_item = QTableWidgetItem(notes if notes else "")

            # Store metadata on first column item.
            times_item.setData(Qt.UserRole, entry_id)
            times_item.setData(Qt.UserRole + 1, project_id)

            if not end_time:
                times_item.setForeground(QColor("#4CAF50"))
                project_item.setForeground(QColor("#4CAF50"))
                notes_item.setForeground(QColor("#4CAF50"))

            self.entries_table.setItem(row_index, 0, times_item)
            self.entries_table.setItem(row_index, 1, project_item)
            self.entries_table.setItem(row_index, 2, notes_item)
    
    def on_entry_double_clicked(self, item):
        """Load an old entry and start a new timer with its metadata."""
        row = item.row()
        meta_item = self.entries_table.item(row, 0)
        notes_item = self.entries_table.item(row, 2)

        if meta_item is None:
            QMessageBox.warning(self, "Error", "Could not load project data.")
            return

        project_id = meta_item.data(Qt.UserRole + 1)
        notes = notes_item.text() if notes_item else ""
        
        if project_id is None:
            QMessageBox.warning(self, "Error", "Could not load project data.")
            return
        
        if self.current_entry:
            QMessageBox.warning(self, "Warning", "A timer is already running! Stop it first.")
            return

        # Ensure filtered list does not hide the target project.
        if self.project_filter.text().strip():
            self.project_filter.clear()
            self.refresh_projects(selected_project_id=project_id)

        if not self.select_project_by_id(project_id):
            QMessageBox.warning(self, "Error", "Das Projekt zum ausgewaehlten Eintrag wurde nicht gefunden.")
            return
        
        # Load the notes/comment
        self.comment_edit.setPlainText(notes)
        
        # Start the timer with the loaded data
        self.start_timer()

    def delete_selected_entry(self):
        """Delete the selected entry from the entries table."""
        selected_items = self.entries_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select an entry to delete.")
            return

        row = selected_items[0].row()
        meta_item = self.entries_table.item(row, 0)
        if meta_item is None:
            QMessageBox.warning(self, "Error", "Selected entry could not be read.")
            return

        entry_id = meta_item.data(Qt.UserRole)
        if entry_id is None:
            QMessageBox.warning(self, "Error", "Selected entry has no ID.")
            return

        if self.current_entry and self.current_entry[0] == entry_id:
            QMessageBox.warning(self, "Warning", "The running entry cannot be deleted. Stop timer first.")
            return

        reply = QMessageBox.question(
            self,
            "Delete Entry",
            "Delete selected entry?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        if self.db.delete_time_entry(entry_id):
            QMessageBox.information(self, "Success", "Entry deleted.")
            self.refresh_entries()
        else:
            QMessageBox.warning(self, "Warning", "Entry could not be deleted.")

    def edit_selected_entry(self):
        """Open dialog to edit the selected entry and its metadata."""
        selected_items = self.entries_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select an entry to edit.")
            return

        row = selected_items[0].row()
        meta_item = self.entries_table.item(row, 0)
        if meta_item is None:
            QMessageBox.warning(self, "Error", "Selected entry could not be read.")
            return

        entry_id = meta_item.data(Qt.UserRole)
        if entry_id is None:
            QMessageBox.warning(self, "Error", "Selected entry has no ID.")
            return

        if self.current_entry and self.current_entry[0] == entry_id:
            QMessageBox.warning(self, "Warning", "Stop the running timer before editing this entry.")
            return

        entry = self.db.get_time_entry_by_id(entry_id)
        if not entry:
            QMessageBox.warning(self, "Error", "Selected entry could not be loaded.")
            return

        if not entry[4]:
            QMessageBox.warning(self, "Warning", "Running entries cannot be edited here. Stop timer first.")
            return

        dialog = EditEntryDialog(self, self.db, entry)
        if dialog.exec_() != QDialog.Accepted:
            return

        data = dialog.get_data()
        updated = self.db.update_time_entry(
            entry_id=entry_id,
            project_id=data["project_id"],
            start_time=data["start_time"],
            end_time=data["end_time"],
            service_type=data["service_type"],
            workplace=data["workplace"],
            notes=data["notes"],
        )

        if updated:
            QMessageBox.information(self, "Success", "Entry updated.")
            self.refresh_entries()
        else:
            QMessageBox.warning(self, "Warning", "Entry could not be updated.")
    
    def add_new_project(self):
        """Open dialog to add a new project."""
        dialog = AddProjectDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            if data['name']:
                try:
                    self.db.add_project(
                        data['name'],
                        data['recipient'],
                        data['process']
                    )
                    self.refresh_projects()
                    QMessageBox.information(self, "Success", f"Project '{data['name']}' added successfully!")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to add project: {str(e)}")

    def edit_selected_project(self):
        """Open dialog to edit the currently selected project."""
        project_id = self.project_combo.currentData(Qt.UserRole)
        if project_id is None:
            QMessageBox.warning(self, "Warning", "Please select a project first!")
            return

        project = self.db.get_project_by_id(project_id)
        if not project:
            QMessageBox.warning(self, "Warning", "Selected project could not be loaded.")
            return

        initial_data = {
            "name": project[1] or "",
            "recipient": project[2] or "",
            "process": project[4] or "",
        }

        dialog = AddProjectDialog(self, initial_data=initial_data, dialog_title="Edit Project")
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            if not data["name"].strip():
                QMessageBox.warning(self, "Warning", "Project Name is required!")
                return

            try:
                updated = self.db.update_project(
                    project_id,
                    data["name"],
                    data["recipient"],
                    data["process"]
                )
                if updated:
                    self.refresh_projects(selected_project_id=project_id)
                    QMessageBox.information(self, "Success", "Project updated successfully!")
                else:
                    QMessageBox.warning(self, "Warning", "Project was not updated.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to update project: {str(e)}")
    
    def start_timer(self):
        """Start a new timer."""
        if not self.project_combo.currentText():
            QMessageBox.warning(self, "Warning", "Please select a project first!")
            return
        
        if self.current_entry:
            QMessageBox.warning(self, "Warning", "A timer is already running!")
            return
        
        project_id = self.project_combo.currentData(Qt.UserRole)
        start_time = datetime.now().isoformat()
        
        # Get dropdown values (extract the code before the pipe)
        leistungsart = self.leistungsart_combo.currentText().split(" | ")[0]
        arbeitsplatz = self.arbeitsplatz_combo.currentText().split(" | ")[0]
        
        # Notes are intentionally saved on stop, not on start.
        entry_id = self.db.start_time_entry(project_id, start_time, "", leistungsart, arbeitsplatz)
        self.current_entry = (entry_id, project_id)
        
        self.timer.start(1000)  # Update every second
        self.update_running_entry()
        self.refresh_entries()
    
    def stop_timer(self):
        """Stop the running timer."""
        if not self.current_entry:
            QMessageBox.warning(self, "Warning", "No timer is running!")
            return
        
        entry_id = self.current_entry[0]
        end_time = datetime.now().isoformat()
        comment = self.comment_edit.toPlainText()
        
        self.db.end_time_entry(entry_id, end_time, comment)
        self.timer.stop()
        self.current_entry = None
        self.current_entry_label.setText("No active timer")
        
        QMessageBox.information(self, "Success", "Timer stopped!")
        self.refresh_entries()
    
    def update_running_entry(self):
        """Update the display of the running entry."""
        if self.current_entry:
            entry_id = self.current_entry[0]
            project_id = self.current_entry[1]
            project = self.db.get_project_by_id(project_id)
            
            # Calculate elapsed time
            entries = self.db.get_time_entries_by_date()
            for entry in entries:
                if entry[0] == entry_id:
                    start_str = entry[3]
                    start = datetime.fromisoformat(start_str)
                    elapsed = datetime.now() - start
                    minutes = int(elapsed.total_seconds() / 60)
                    seconds = int(elapsed.total_seconds() % 60)
                    
                    text = f"Running: {project[1]} - {minutes:02d}:{seconds:02d}"
                    self.current_entry_label.setText(text)
                    break
    
    def add_manual_entry(self):
        """Add a manual time entry."""
        if not self.project_combo.currentText():
            QMessageBox.warning(self, "Warning", "Please select a project first!")
            return
        
        # Build start/end datetime from selected date + selected times.
        selected_date = self.date_edit.date().toPyDate()
        start_time_str = datetime.combine(selected_date, self.start_time.time().toPyTime())
        end_time_str = datetime.combine(selected_date, self.end_time.time().toPyTime())
        
        if start_time_str >= end_time_str:
            QMessageBox.warning(self, "Warning", "End time must be after start time!")
            return
        
        project_id = self.project_combo.currentData(Qt.UserRole)
        comment = self.comment_edit.toPlainText()
        
        # Get dropdown values (extract the code before the pipe)
        leistungsart = self.leistungsart_combo.currentText().split(" | ")[0]
        arbeitsplatz = self.arbeitsplatz_combo.currentText().split(" | ")[0]
        
        # Create time entry
        entry_id = self.db.start_time_entry(project_id, start_time_str.isoformat(), comment, leistungsart, arbeitsplatz)
        self.db.end_time_entry(entry_id, end_time_str.isoformat())
        
        # Clear comment field after adding entry
        self.comment_edit.clear()
        
        QMessageBox.information(self, "Success", "Time entry added!")
        self.refresh_entries()
    
    def export_to_excel(self):
        """Export all entries across all days to Excel."""
        entries = self.db.get_time_entries_by_date()
        
        if not entries:
            QMessageBox.warning(self, "Warning", "No entries to export!")
            return
        
        try:
            filename = export_to_excel(entries)
            QMessageBox.information(self, "Success", f"Exported to {filename}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Export failed: {str(e)}")
    
    def end_day(self):
        """End the work day and prompt for export."""
        reply = QMessageBox.question(
            self, 
            "End Day", 
            "Do you want to export today's entries before ending the day?",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
        )
        
        if reply == QMessageBox.Cancel:
            return
        
        if reply == QMessageBox.Yes:
            try:
                entries = self.db.get_time_entries_by_date()
                if entries:
                    export_to_excel(entries)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed: {str(e)}")
        
        QMessageBox.information(self, "End Day", "Have a great day!")
        self.timer.stop()
