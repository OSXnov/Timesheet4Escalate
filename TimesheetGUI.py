import tkinter as tk

# Create the main window
root = tk.Tk()
root.title("Timesheet Automator")

# Configure the grid 
for i in range(2):
    root.grid_columnconfigure(i, weight=1)
    root.grid_rowconfigure(i, weight=1)

# Create and place the welcome message in the middle (row 1, column 1 and 2)
welcome_label = tk.Label(root, text="Welcome to the Timesheet Automator", font=("Arial", 10))
welcome_label.grid(row=1, column=1, columnspan=2, pady=20, padx=20)

# Create the START button in the bottom-right corner (row 3, column 3)
start_button = tk.Button(root, text="START", font=("Arial", 12))
start_button.grid(row=3, column=3, sticky="e", padx=20, pady=20)

# Create the CLOSE button in the bottom-left corner (row 3, column 0)
close_button = tk.Button(root, text="CLOSE", font=("Arial", 12), command=root.quit)
close_button.grid(row=3, column=0, sticky="w", padx=20, pady=20)

root.mainloop()
