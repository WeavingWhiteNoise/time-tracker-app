"""Export utilities for generating Excel reports."""
import hashlib
import os
import getpass
import pandas as pd
from datetime import datetime
from openpyxl.styles import Font, PatternFill


def get_user_hash() -> str:
    """Return a short deterministic hash based on the current Windows username."""
    username = getpass.getuser()
    return hashlib.sha256(username.encode()).hexdigest()[:12]


def export_to_excel(time_entries: list, filename: str = None) -> str:
    """
    Export time entries to Excel file with monthly filtering.
    
    Args:
        time_entries: List of time entries with project metadata
        filename: Output filename (default: time_tracker_YYYY-MM.xlsx)
    
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
            'Stunden': f"{round((duration or 0) / 60, 2)}",
            'Projekt Name': project_name,
            'Empfaenger': recipient or '',
            'Leistungsart': service_type or '',
            'Vorgang': process or '',
            'Arbeitsplatz': workplace or '',
            'Kommentar': notes or ''
        })

    df = pd.DataFrame(data)
    df['_sort_start'] = pd.to_datetime(df['Start Time'], errors='coerce')
    df = df.sort_values(by=['Date', '_sort_start', 'End Time'], ascending=[True, True, True], na_position='last')
    df = df.drop(columns=['_sort_start'])
    
    # Generate filename with user hash + month if not provided
    if filename is None:
        if not df.empty and df['Date'].notna().any():
            first_date = df[df['Date'].notna()]['Date'].iloc[0]
            month_str = pd.Timestamp(first_date).strftime("%Y-%m")
        else:
            month_str = datetime.now().strftime("%Y-%m")
        user_hash = get_user_hash()
        filename = f"{user_hash}_{month_str}.xlsx"
    
    # Create Excel writer
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Time Entries', index=False)
        
        # Format columns
        workbook = writer.book
        worksheet = writer.sheets['Time Entries']
        
        # Define styles
        bold_font = Font(bold=True)
        light_red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        
        # Get header row
        header_row = worksheet[1]
        
        # Format headers: make bold and add autofilter
        for cell in header_row:
            cell.font = bold_font
        
        # Add autofilter to headers
        worksheet.auto_filter.ref = f"A1:{chr(64 + len(df.columns))}{len(df) + 1}"
        
        # Adjust column widths and mark empty cells with red
        columns_mapping = {i + 1: col_name for i, col_name in enumerate(df.columns)}
        
        for col_idx, column in enumerate(worksheet.columns, 1):
            max_length = 0
            column_letter = column[0].column_letter
            
            for row_idx, cell in enumerate(column):
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
                
                # Mark empty cells (except header) with light red
                if row_idx > 0 and (cell.value is None or str(cell.value).strip() == ''):
                    cell.fill = light_red_fill
            
            adjusted_width = min(max_length + 6, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        # Date column (A): true Excel date with DD.MM.YYYY display format.
        for cell in worksheet['A'][1:]:
            if cell.value is not None:
                cell.number_format = 'DD.MM.YYYY'

        # Stunden column (D): format as number
        for cell in worksheet['D'][1:]:
            if cell.value is not None:
                cell.value = str(cell.value)

        # Add summary sheet
        summary_data = {
            'Metric': ['Total Stunden', 'Number of Entries', 'Export Date'],
            'Value': [round(sum(float(x) for x in df['Stunden']), 2),
                     len(df),
                     datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Format summary sheet headers
        summary_sheet = writer.sheets['Summary']
        for cell in summary_sheet[1]:
            cell.font = bold_font
    
    return filename
