import tkinter as tk
from tkinter import messagebox
from tkcalendar import Calendar
import shutil
import os
import openpyxl
from datetime import datetime

def insert_date_into_excel(date_value, file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        sheet['B15'] = date_value
        workbook.save(file_path)
        messagebox.showinfo("Success", f"Date {date_value} has been inserted into cell B15.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while inserting the date: {e}")

def open_calendar(file_path):
    calendar_window = tk.Toplevel(root)
    calendar_window.title("Select Date for TimeSheet Period")
    calendar_window.geometry("300x400")

    calendar = Calendar(calendar_window, selectmode='day', date_pattern='mm/dd/yyyy')
    calendar.pack(padx=20, pady=20)

    def select_date():
        selected_date = calendar.get_date()
        try:
            date_value = datetime.strptime(selected_date, '%m/%d/%Y').strftime('%Y-%m-%d')
            insert_date_into_excel(date_value, file_path)
            calendar_window.destroy()
        except ValueError:
            messagebox.showerror("Error", "Invalid date format selected.")

    submit_button = tk.Button(calendar_window, text="SUBMIT", command=select_date, font=("Arial", 10))
    submit_button.pack(pady=10)

def start_renaming():
    def open_name_window(new_file_path):
        name_window = tk.Toplevel(root)
        name_window.title("Set Name for Timesheet")
        name_window.geometry("300x300")

        def submit_name():
            user_name = name_entry.get()
            if user_name:
                try:
                    workbook = openpyxl.load_workbook(new_file_path)
                    sheet = workbook.active
                    sheet['C7'] = user_name
                    workbook.save(new_file_path)
                    messagebox.showinfo("Success", f"Name '{user_name}' has been set in cell C7.")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not write name to file: {e}")
            else:
                messagebox.showerror("Error", "Please enter a name.")

        set_name_label = tk.Label(name_window, text="Enter your name:", font=("Arial", 10))
        set_name_label.pack(pady=10)

        name_entry = tk.Entry(name_window, width=30)
        name_entry.pack(pady=10)

        submit_name_button = tk.Button(name_window, text="SUBMIT", command=submit_name, font=("Arial", 10))
        submit_name_button.pack(pady=10)

        set_period_button = tk.Button(name_window, text="Set TimeSheet Period", font=("Arial", 10),
                                      command=lambda: open_calendar(new_file_path))
        set_period_button.pack(pady=10)

        close_button = tk.Button(name_window, text="CLOSE", font=("Arial", 10), command=name_window.destroy)
        close_button.pack(pady=10)

        next_button = tk.Button(name_window, text="NEXT", font=("Arial", 10))  # Next button does nothing for now
        next_button.pack(pady=10)

    rename_window = tk.Toplevel(root)
    rename_window.title("Rename Timesheet File")
    rename_window.geometry("300x200")

    label = tk.Label(rename_window, text="Enter the new file name:", font=("Arial", 10))
    label.pack(pady=10)

    file_name_entry = tk.Entry(rename_window, width=30)
    file_name_entry.pack(pady=10)

    def submit_file_name():
        new_file_name = file_name_entry.get()
        original_file = 'TimesheetTemplate.xlsx'
        new_file_path = f'{new_file_name}.xlsx'

        if os.path.exists(original_file):
            try:
                shutil.copyfile(original_file, new_file_path)
                messagebox.showinfo("Success", f"File renamed to {new_file_name}.xlsx successfully.")
                rename_window.destroy()
                open_name_window(new_file_path)
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")
        else:
            messagebox.showerror("Error", "Original file not found!")

    submit_button = tk.Button(rename_window, text="SUBMIT", command=submit_file_name, font=("Arial", 10))
    submit_button.pack(pady=10)

    next_file_ren_button = tk.Button(rename_window, text="NEXT", font=("Arial", 10))  # Next button does nothing for now
    next_file_ren_button.pack(pady=10)

# Main window
root = tk.Tk()
root.title("Timesheet Automator")
root.geometry("400x300")

welcome_label = tk.Label(root, text="Welcome to the Timesheet Automator", font=("Arial", 12))
welcome_label.pack(pady=20)

start_button = tk.Button(root, text="START", font=("Arial", 12), command=start_renaming)
start_button.pack(side=tk.RIGHT, padx=20, pady=20)

close_button = tk.Button(root, text="CLOSE", font=("Arial", 12), command=root.quit)
close_button.pack(side=tk.LEFT, padx=20, pady=20)

# Run the Tkinter event loop
root.mainloop()
