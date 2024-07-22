#green color formatting is working (status and new ticket condition)
#red color formatting for missing items
#is sorted by project

import pandas as pd
from datetime import datetime

# Excel file path
file_path = r'C:\Users\michaeljohn.roguel\Desktop\jenkins_mini_project\sample_file.xlsx'

# Read Excel file
df = pd.read_excel(file_path)

# Convert 'Outlook Completion' column to the desired format
df['Outlook Completion'] = pd.to_datetime(df['Outlook Completion']).dt.strftime('%m-%d-%Y')

# Get today's date
today_date = datetime.today().strftime('%m-%d-%Y')

# Get the previous latest date in the 'Date' column
previous_latest_date = (pd.to_datetime(df['Date']).max() - pd.Timedelta(days=1)).strftime('%m-%d-%Y')

# Filter df based on today's date and previous latest date
df_today = df[df['Date'] == today_date]
df_previous = df[df['Date'] == previous_latest_date]

# Sort df_today by the 'Project' column
df_today = df_today.sort_values(by='Project')

# Identify missing tickets
missing_tickets_condition = (df_previous['Status'].isin(['On-hold', 'On-going'])) & (~df_previous['Ticket'].isin(df_today['Ticket']))
missing_tickets = df_previous[missing_tickets_condition][['Ticket', 'Employee']]

# Identify new tickets
new_tickets = df_today[~df_today['Ticket'].isin(df_previous['Ticket'])]

# Prepare the formatted text for "Status Report"
def format_status_report(df_today, new_tickets):
    status_report = []
    status_list = []
    from_today_new = []
    current_project = None
    project_counter = 1
    
    for _, row in df_today.iterrows():
        if row['Project'] != current_project:
            current_project = row['Project']
            #status_report.append(f"{current_project}")
            status_report.append(f"{project_counter}. {current_project}")
            status_list.append(None)  # No status for project titles
            from_today_new.append(False)  # No color for project titles
            project_counter += 1
        
        if row['Status'] in ['On-going', 'On-hold', 'Not yet started']:
            formatted_text = f"{row['Ticket']}: {row['Comment']}, Outlook Completion: {row['Outlook Completion']} [{row['Employee']}]"
        elif row['Status'] in ['Resolved', 'Closed']:
            formatted_text = f"[{row['Status']}] {row['Ticket']}: {row['Comment']}"
        
        status_report.append(formatted_text)
        status_list.append(row['Status'])
        from_today_new.append(row['Ticket'] in new_tickets['Ticket'].values)
    
    return status_report, status_list, from_today_new

status_report, status_list, from_today_new = format_status_report(df_today, new_tickets)

# Missing data
missing_data_text = ["Missing Data:"]
missing_data_formatted = [f"{row['Ticket']} | [{row['Employee']}]" for _, row in missing_tickets.iterrows()]
missing_data_text.extend(missing_data_formatted)

# Remove duplicates from status report
status_report = list(dict.fromkeys(status_report))

# Output file path
output_file_path = r'C:\Users\michaeljohn.roguel\Desktop\jenkins_mini_project\Status_Report.xlsx'

# Write the result dataframes to an Excel file with separate sheets
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    df_today.to_excel(writer, sheet_name='Today\'s Data', index=False)
    missing_tickets.to_excel(writer, sheet_name='Missing Tickets', index=False)
    
    # Create the "Status Report" sheet
    status_report_df = pd.DataFrame(status_report, columns=['Status Report'])
    status_report_df.to_excel(writer, sheet_name='Status Report', index=False)

    # Adjust formatting for the "Status Report" sheet
    workbook = writer.book
    worksheet = writer.sheets['Status Report']
    
    # Write "Status Report" in cell A1
    worksheet.write('A1', 'Status Report')
    
    # Define a format for left-aligned text
    left_align_format = workbook.add_format({'align': 'left'})
    
    # Define a format for light green background
    light_green_format = workbook.add_format({'bg_color': '#C6EFCE'})
    
    # Define a format for light red background
    light_red_format = workbook.add_format({'bg_color': '#FFC7CE'})
    
    # Apply the left-aligned format to all cells initially
    for row_num, value in enumerate(status_report):
        worksheet.write(row_num + 1, 0, value, left_align_format)
        
        # Apply conditional formatting to color cells light green based on status and only for today's new data
        if status_list[row_num] in ['On-going', 'On-hold', 'Not yet started'] and from_today_new[row_num]:
            worksheet.write(row_num + 1, 0, value, light_green_format)
    
    # Paste missing tickets
    startrow = len(status_report) + 3
    for i, text in enumerate(missing_data_text):
        if i == 0:  # The header "Missing Data:"
            worksheet.write(startrow + i, 0, text, left_align_format)
        else:  # The missing data rows
            worksheet.write(startrow + i, 0, text, light_red_format)

print(f"Output file created: {output_file_path}")
print("File creation completed")