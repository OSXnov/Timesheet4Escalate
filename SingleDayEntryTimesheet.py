import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
from openpyxl import load_workbook
from tkinter import ttk

# Excel slot references for dates, descriptions, and hours
date_slots_week1 = ['B15', 'B18', 'B21', 'B24', 'B27', 'B30', 'B33']
date_slots_week2 = ['H15', 'H18', 'H21', 'H24', 'H27', 'H30', 'H33']

description_slots = ['C15', 'C18', 'C21', 'C24', 'C27', 'C30', 'C33',
                     'I15', 'I18', 'I21', 'I24', 'I27', 'I30', 'I33']

hours_slots = ['E15', 'E18', 'E21', 'E24', 'E27', 'E30', 'E33',
               'J15', 'J18', 'J21', 'J24', 'J27', 'J30', 'J33']

# Function to select a day and input description/hours
def select_day(wb, sheet, slot_index):
    def submit_entry():
        description = desc_entry.get("1.0", 'end-1c')
        hours = hours_entry.get()

        if no_work.get() == 1:  # If the user didn't work
            description = ""
            hours = 0

        sheet[description_slots[slot_index]] = description
        sheet[hours_slots[slot_index]] = int(hours)
        wb.save(filename)

        messagebox.showinfo("Success", "Data saved successfully!")
        input_window.destroy()  # Close after submission

    def go_back():
        input_window.destroy()  # Close the window and go back to day selection
        SelectWeekDay(wb, sheet)

    input_window = tk.Toplevel()
    input_window.title("Input Day Details")
    input_window.geometry("400x300")

    tk.Label(input_window, text="Description:").pack(pady=10)
    desc_entry = scrolledtext.ScrolledText(input_window, width=30, height=4)
    desc_entry.pack(pady=10)

    tk.Label(input_window, text="Hours Worked:").pack(pady=10)
    hours_entry = tk.Entry(input_window)
    hours_entry.pack(pady=10)

    no_work = tk.IntVar()
    no_work_check = tk.Checkbutton(input_window, text="Didn't Work", variable=no_work)
    no_work_check.pack(pady=10)

    submit_button = tk.Button(input_window, text="Submit", command=submit_entry)
    submit_button.pack(side="right", padx=10, pady=10)

    back_button = tk.Button(input_window, text="Back", command=go_back)
    back_button.pack(side="left", padx=10, pady=10)

# Function to display buttons for weekdays
def SelectWeekDay(wb, sheet):
    week_window = tk.Toplevel()
    week_window.title("Select a Day")
    week_window.geometry("500x400")

    # Fetch dates from the Excel sheet
    week1_dates = [sheet[cell].value for cell in date_slots_week1]
    week2_dates = [sheet[cell].value for cell in date_slots_week2]

    # Weekday names and corresponding date slot indices
    days = [
        ("Monday", 0), ("Tuesday", 1), ("Wednesday", 2),
        ("Thursday", 3), ("Friday", 4), ("Saturday", 5), ("Sunday", 6)
    ]

    # Create buttons for the two weeks
    for day, index in days:
        week1_date = week1_dates[index]
        week2_date = week2_dates[index]

        # First week (left column)
        tk.Button(week_window, text=f"{day}, {week1_date}", command=lambda i=index: select_day(wb, sheet, i)).grid(row=index, column=0, padx=20, pady=10)
        # Second week (right column)
        tk.Button(week_window, text=f"{day}, {week2_date}", command=lambda i=index+7: select_day(wb, sheet, i)).grid(row=index, column=1, padx=20, pady=10)

# File browser to select and open an Excel file
def filebrowser():
    global filename
    filename = filedialog.askopenfilename(
        initialdir="/", title="Select file",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )

    if filename:
        wb = load_workbook(filename)
        sheet = wb.active
        SelectWeekDay(wb, sheet)  # Call the function to select a weekday

# Initial window with Start Entry button
root = tk.Tk()
root.title("Single Day Entries Timesheet")
root.geometry("200x200")

def open_new_window():
    new_window = tk.Toplevel(root)
    new_window.title("File Options")
    new_window.geometry("500x200")

    create_button = tk.Button(new_window, text="Create File", font=("Sonoma", 12))
    create_button.pack(pady=20)

    find_button = tk.Button(new_window, text="Find File", font=("Sonoma", 12), command=filebrowser)
    find_button.pack(pady=20)

start_button = tk.Button(root, text="Start Entry", font=("Sonoma", 12), command=open_new_window)
start_button.pack(pady=50)

root.mainloop()
