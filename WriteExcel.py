import openpyxl
from pathlib  import Path
import os
from datetime import datetime

# Load the workbook and select the active sheet
file_path = Path('../Timesheet4Escalate/TimesheetTemplate.xlsx')
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Define the date you want to insert
date_value = datetime.now().strftime('%Y-%m-%d')  # Example: today's date

# Insert the date value in cell B15
#sheet['B15'] = date_value

# Save the workbook
workbook.save(file_path)

print(f"Date {date_value} has been inserted into cell B15.")

sheet['B15'] = None
print("Cell B15 has been cleared.")
