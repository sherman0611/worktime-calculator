import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Change file name
input_file = 'Register_Timekeeping_Report_2024-10-30_00_00GMT_to_2024-11-26_23_59GMT.xlsx'
output_file = 'employee_work_hours.xlsx'

def calculate_work_hours(input_file, output_file):
    df = pd.read_excel(input_file, header=None)
    
    # Select required columns for calculation
    columns_to_select = [0, 1, 5, 6]
    df = df.iloc[:, columns_to_select]
    
    # Add column names
    df.columns = ['Date', 'Time', 'Employee', 'Event']
    
    # Strip (UTC) string from time
    df['Time'] = df['Time'].str.replace(r"\(UTC\)", "", regex=True).str.strip()
    
    # Combine 'Date' and 'Time' columns to create a datetime column
    df['Datetime'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Time'], errors='coerce')
    
    # Filter out invalid rows
    df = df[df['Datetime'].notna()]
    
    # Calculate work hours
    work_hours = []

    for i in range(1, len(df)):
        if df.iloc[i-1]['Event'] == 'Clock In' and df.iloc[i]['Event'] == 'Clock Out':
            work_time = (df.iloc[i]['Datetime'] - df.iloc[i-1]['Datetime']).total_seconds() / 3600
            work_hours.append(round(work_time, 2))
        else:
            work_hours.append(None)

    # Add the work hours to the DataFrame
    df['Work Hours'] = [None] + work_hours

    df = df.drop(columns=['Datetime'])

    # Group by Employee and calculate total hours
    employee_groups = df.groupby('Employee')
    
    # Create a new DataFrame with empty rows and total hours inserted
    new_df = pd.DataFrame()
    for _, group in employee_groups:
        # Calculate total hours for the employee, ignoring None values
        total_hours = group['Work Hours'].sum(skipna=True)
        
        # Create a total hours row with the same number of columns as the original DataFrame
        total_hours_row = group.iloc[0:1].copy()  # Copy the first row's structure
        total_hours_row.iloc[0] = np.nan  # Use np.nan instead of '' to avoid dtype conflict
        total_hours_row['Event'] = 'Total Hours'
        total_hours_row['Work Hours'] = total_hours
        
        # Concatenate the group with its total hours row and an empty row
        employee_df = pd.concat([group, 
                                 total_hours_row, 
                                 pd.DataFrame([[''] * len(group.columns)], columns=group.columns)
                                ])
        
        # Add to the main DataFrame
        new_df = pd.concat([new_df, employee_df])
    
    new_df.to_excel(output_file, index=False)
    
    ### Set excel styles ################################################################
    wb = load_workbook(output_file)
    ws = wb.active
    
    # Adjust column widths based on the maximum length of the content in each column
    for col in range(1, len(new_df.columns) + 1):
        column = get_column_letter(col)
        max_length = 0
        for row in ws.iter_rows(min_col=col, max_col=col):
            for cell in row:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width
    
    # Lock the top row (header) for scrolling
    ws.freeze_panes = 'A2'
    
    # Apply background colour
    bg_color = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    for row in ws.iter_rows():
        if row[0].value and row[0].value != "Date":
            for cell in row:
                cell.fill = bg_color
        if row[3].value == 'Total Hours':
            row[4].fill = yellow
    
    wb.save(output_file)
    
    print(f"File saved successfully: {output_file}")

calculate_work_hours(input_file, output_file)