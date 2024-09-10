import openpyxl
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
from tkcalendar import Calendar


def insert_date_into_excel(date_value):
    
    file_path = Path('../Timesheet4Escalate/TimesheetTemplateDummy.xlsx')
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

   
    sheet['B15'] = date_value

    
    workbook.save(file_path)

   
    messagebox.showinfo("Success", f"Date {date_value} has been inserted into cell B15.")

# Function to handle date selection from the calendar
def get_selected_date():
    selected_date = calendar.get_date()  # Get the selected date from the calendar
    try:
        # Parse the selected date (from the format mm/dd/yyyy to yyyy-mm-dd)
        date_value = datetime.strptime(selected_date, '%m/%d/%Y').strftime('%Y-%m-%d')
        
        # Insert the date into the Excel file
        insert_date_into_excel(date_value)

    except ValueError:
        messagebox.showerror("Error", "Invalid date format. Please try again.")

# Create the main Tkinter window
root = tk.Tk()
root.title("Timesheet Date Selection")

# Create the calendar widget
calendar = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy")
calendar.pack(pady=20)

# Create a button to confirm the selected date
select_button = tk.Button(root, text="Select Date", command=get_selected_date)
select_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
