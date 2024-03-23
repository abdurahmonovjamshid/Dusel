import pandas as pd

# Read the Excel file and parse the date column
xls = pd.ExcelFile('file.xlsx')
date_format = "%d.%m.%Y"  # Specify the date format
parse_dates = ['date']  # Replace 'date' with the actual column name

# Create a new ExcelWriter object
with pd.ExcelWriter('sorted_file_new.xlsx', mode='w', engine='openpyxl') as writer:
    # Sort and save each sheet separately
    for sheet_name in xls.sheet_names:
        # Read the sheet and parse the date column
        df = pd.read_excel(xls, sheet_name, parse_dates=parse_dates, date_format=date_format)
        
        # Sort the data by the date column
        df_sorted = df.sort_values('date')
        
        # Format the date column without time component
        df_sorted['date'] = df_sorted['date'].dt.strftime(date_format)
        
        # Save the sorted data to a new sheet in the same Excel file
        df_sorted.to_excel(writer, sheet_name=sheet_name, index=False)