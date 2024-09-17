# Single Day entry Code
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
import shutil
import os
import openpyxl
from datetime import datetime

def open_new_window():
    new_window = tk.Toplevel(root)
    new_window.title("File Options")
    new_window.geometry("500x200")

    # Button to Create File (functionality to be added later)
    create_button = tk.Button(new_window, text="Create File", font=("Sonoma", 12))
    create_button.pack(pady=20)

    # missing the loading of its data for the excel, needs the excel slots that timesheet has (its always the same)
    find_button = tk.Button(new_window, text="Find File", font=("Sonoma", 12), command=filebrowser)
    find_button.pack(pady=20)

def filebrowser():
    filename=filedialog.askopenfilename(initialdir ="/", title = "Select file", filetypes = (("Text files","*.xlsx*"),("all files","*.*")))
    label_file_explorer = Label(window, text = "File Explorer using Tkinter", width = 100, height = 4, fg = "blue")
    button_explore = Button(window, text = "Browse Files",command = browseFiles)
    button_exit = Button(window, text = "Exit", command = exit) 


    
    

root = tk.Tk()
root.title("Single Day Entries Timesheet")
root.geometry("200x200")

start_button = tk.Button(root, text="Start Entry", font=("Sonoma", 12), command=open_new_window)
start_button.pack(pady=50)

root.mainloop()
