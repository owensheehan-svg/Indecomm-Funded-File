#!/usr/bin/env python3
"""
Airtable to Excel Daily Export
Exports a specific Airtable view to an Excel file with formatting
"""

import os
import requests
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Configuration
AIRTABLE_TOKEN = os.environ.get('AIRTABLE_TOKEN')
BASE_ID = 'appgBl5EHB3qFtOPl'
TABLE_NAME = 'Data Input'
VIEW_NAME = 'Indecomm Funded File'

def get_airtable_records():
    """Gather all records from the Airtable view"""
    url = f'https://api.airtable.com/v0/{BASE_ID}/{TABLE_NAME}'
    headers = {
        'Authorization': f'Bearer {AIRTABLE_TOKEN}',
        'Content-Type': 'application/json'
    }
    
    params = {
        'view': VIEW_NAME
    }
    
    all_records = []
    offset = None
    
    print("📥 Fetching records from Airtable...")
    
    while True:
        if offset:
            params['offset'] = offset
            
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        
        data = response.json()
        records = data.get('records', [])
        all_records.extend(records)
        
        offset = data.get('offset')
        if not offset:
            break
    
    print(f"✅ Fetched {len(all_records)} records")
    return all_records

def records_to_dataframe(records):
    """Convert Airtable records to a pandas DataFrame"""
    if not records:
        return pd.DataFrame()
    
    # Extract fields from each record
    data = [record['fields'] for record in records]
    df = pd.DataFrame(data)
    
    return df

def format_excel(filename):
    """Apply nice formatting to the Excel file"""
    wb = load_workbook(filename)
    ws = wb.active
    
    # Style the header row
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        
        adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Freeze the header row
    ws.freeze_panes = 'A2'
    
    wb.save(filename)

def main():
    """Main execution function"""
    if not AIRTABLE_TOKEN:
        raise ValueError("❌ AIRTABLE_TOKEN environment variable not set")
    
    # Fetch records
    records = get_airtable_records()
    
    if not records:
        print("⚠️  No records found in view")
        return
    
    # Convert to DataFrame
    df = records_to_dataframe(records)
    
    # Generate filename with timestamp
    timestamp = datetime.now().strftime('%Y-%m-%d')
    filename = f'airtable_export_{timestamp}.xlsx'
    
    # Export to Excel
    print(f"💾 Exporting to {filename}...")
    df.to_excel(filename, index=False, sheet_name='Export')
    
    # Apply formatting
    print("🎨 Applying formatting...")
    format_excel(filename)
    
    print(f"✅ Export complete! File saved as: {filename}")
    print(f"📊 Exported {len(df)} rows and {len(df.columns)} columns")

if __name__ == '__main__':
    main()
