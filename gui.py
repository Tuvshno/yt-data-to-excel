import pandas as pd
from tkinter import *
from tkinter import filedialog

# Function to open the first Excel file
def open_first_file():
    global df1
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    df1 = pd.read_excel(file_path)
    check_display()

# Function to open the second Excel file
def open_second_file():
    global df2, second_file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    second_file_path = file_path
    df2 = pd.read_excel(file_path)
    check_display()

# Function to save the updated second Excel file
def save_file():
    df2.to_excel(second_file_path, index=False)

# Function to display the current row data
def display_row():
    row_data = df1.iloc[row_index].to_string()
    row_data_label.config(text=row_data)

# Function to check if both files are loaded and display the first row
def check_display():
    if df1 is not None and df2 is not None:
        display_row()

# Function to keep the row data
def keep_row():
    global df2
    df2 = pd.concat([df2, df1.iloc[[row_index]]], ignore_index=True)
    save_file()  # Automatically save the updated second spreadsheet
    next_row()

# Function to move to the next row
def next_row():
    global row_index
    row_index += 1
    display_row()

# Initialize the main window
root = Tk()
root.title("Excel Row Keeper")

# Define global variables
row_index = 0
df1 = None
df2 = None
second_file_path = ""

# Create and pack UI elements
open_first_file_button = Button(root, text="Open First File", command=open_first_file)
open_first_file_button.pack()

open_second_file_button = Button(root, text="Open Second File", command=open_second_file)
open_second_file_button.pack()

row_data_label = Label(root, text="")
row_data_label.pack()

keep_button = Button(root, text="Keep", command=keep_row)
keep_button.pack()

next_button = Button(root, text="Next", command=next_row)
next_button.pack()

# Run the main loop
root.mainloop()
