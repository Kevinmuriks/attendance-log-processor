import pandas as pd
from datetime import datetime, timedelta

def split_attendance_sheets(input_file, output_file):
    # Load the Excel file
    xls = pd.ExcelFile(input_file)
    df = pd.read_excel(xls, sheet_name="Sheet1")
    
    # Convert Date/Time to datetime format
    df['Date/Time'] = pd.to_datetime(df['Date/Time'])
    
    # Extract date and create separate columns for check-in and check-out
    df['Date'] = df['Date/Time'].dt.date
    df['Time'] = df['Date/Time'].dt.time
    
    # Sort data to ensure correct order
    df = df.sort_values(by=['Name', 'Date/Time'])
    
    # Create a new DataFrame to store processed data
    processed_data = []
    
    # Group by employee name or number if name is missing
    for key, group in df.groupby(df['Name'].fillna(df['No.']).astype(str)):
        daily_records = {}
        
        for _, row in group.iterrows():
            date = row['Date']
            if date not in daily_records:
                daily_records[date] = {'Check-In': row['Date/Time'], 'Check-Out': None}
            else:
                daily_records[date]['Check-Out'] = row['Date/Time']
        
        # Convert to DataFrame
        records_df = pd.DataFrame([
            {
                "Date": d,
                "Check-In": v['Check-In'].time() if v['Check-In'] else None,
                "Check-Out": v['Check-Out'].time() if v['Check-Out'] else None,
                "Overtime (Hours)": max(((v['Check-Out'] - v['Check-In']).total_seconds() / 3600) - 8, 0) if v['Check-In'] and v['Check-Out'] else None
            } 
            for d, v in daily_records.items()
        ])
        
        processed_data.append((key, records_df))
    
    # Create a new Excel writer
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for key, records_df in processed_data:
            records_df.to_excel(writer, sheet_name=key[:31], index=False)  # Sheet names max length is 31
        
        print(f"Processed {len(processed_data)} employees into separate sheets.")

# Example usage
input_path = "Er-Attendance logs_24 to 01 2025.xlsx"
output_path = "Processed_Attendance_Logs.xlsx"
split_attendance_sheets(input_path, output_path)
