"""Export utilities for generating Excel reports."""
import pandas as pd
from datetime import datetime


def export_to_excel(time_entries: list, filename: str = None) -> str:
    """
    Export time entries to Excel file.
    
    Args:
        time_entries: List of time entries with project metadata
        filename: Output filename (default: time_tracker_{date}.xlsx)
    
    Returns:
        Path to the created Excel file
    """
    if not time_entries:
        raise ValueError("No time entries to export")
    
    # Create DataFrame
    data = []
    for entry in time_entries:
        entry_id, project_id, project_name, start_time, end_time, duration, recipient, service_type, process, workplace, notes = entry

        start_dt = pd.to_datetime(start_time, errors='coerce')

        data.append({
            'Date': start_dt.normalize() if pd.notna(start_dt) else None,
            'Start Time': start_time,
            'End Time': end_time,
            'Stunden': round((duration or 0) / 60, 2),
            'Projekt Name': project_name,
            'Empfaenger': recipient or '',
            'Leistungsart': service_type or '',
            'Vorgang': process or '',
            'Arbeitsplatz': workplace or '',
            'Kommentar': notes or ''
        })

    df = pd.DataFrame(data)
    df['_sort_start'] = pd.to_datetime(df['Start Time'], errors='coerce')
    df = df.sort_values(by=['_sort_start', 'End Time'], ascending=[True, True], na_position='last')
    df = df.drop(columns=['_sort_start'])
    
    # Generate filename if not provided
    if filename is None:
        today = datetime.now().strftime("%Y-%m-%d")
        filename = f"time_tracker_{today}.xlsx"
    
    # Create Excel writer
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Time Entries', index=False)
        
        # Format columns
        workbook = writer.book
        worksheet = writer.sheets['Time Entries']
        
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        # Date column: true Excel date with DD.MM.YYYY display format.
        for cell in worksheet['A'][1:]:
            if cell.value is not None:
                cell.number_format = 'DD.MM.YYYY'

        # Stunden column (D): prefix each value with apostrophe to force text.
        for cell in worksheet['D'][1:]:
            if cell.value is not None:
                cell.value = f"{cell.value}"

        # Add summary sheet
        summary_data = {
            'Metric': ['Total Stunden', 'Number of Entries', 'Export Date'],
            'Value': [round(df['Stunden'].sum(), 2),
                     len(df),
                     datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    return filename
