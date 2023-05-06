import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import ttk

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
    save_file()
    next_row()

# Function to move to the next row
def next_row():
    global row_index
    row_index += 1
    display_row()

# Initialize the main window
root = Tk()
root.title("Excel Row Keeper")
root.geometry("800x600")
root.configure(bg="white")

# Define global variables
row_index = 0
df1 = None
df2 = None
second_file_path = ""

# Create and pack UI elements
frame1 = Frame(root, bg="white", padx=20, pady=10)
frame1.pack(fill=X)

open_first_file_button = ttk.Button(frame1, text="Open First File", command=open_first_file)
open_first_file_button.grid(row=0, column=0, padx=5)

open_second_file_button = ttk.Button(frame1, text="Open Second File", command=open_second_file)
open_second_file_button.grid(row=0, column=1, padx=5)

frame1.grid_columnconfigure(0, weight=1)
frame1.grid_columnconfigure(1, weight=1)

frame2 = Frame(root, bg="white", padx=20, pady=10)
frame2.pack(fill=X)

row_data_label = Label(frame2, text="", wraplength=750, font=("Arial", 12), bg="white", justify=LEFT)
row_data_label.pack(pady=(40, 20))

frame3 = Frame(root, bg="white", padx=20, pady=10)
frame3.pack(fill=X)

keep_button = ttk.Button(frame3, text="Keep", command=keep_row)
keep_button.grid(row=0, column=0, padx=5)

next_button = ttk.Button(frame3, text="Next", command=next_row)
next_button.grid(row=0, column=1, padx=5)

frame3.grid_columnconfigure(0, weight=1)
frame3.grid_columnconfigure(1, weight=1)

# Run the main loop
root.mainloop()
