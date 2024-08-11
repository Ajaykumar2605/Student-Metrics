import tkinter as tk
from tkinter import *
from tkinter.ttk import Combobox
from tkinter import messagebox
import openpyxl
from openpyxl import Workbook
import requests
from io import BytesIO
import pathlib
from PIL import Image, ImageTk  # Ensure correct import

# Function to download and open an Excel file from a URL
def download_excel_file(url):
    response = requests.get(url)
    response.raise_for_status()  # Check for HTTP errors
    return BytesIO(response.content)

# URL for the Excel file
xlsx_file_url = "https://github.com/Ajaykumar2605/Student-Metrics/raw/main/source%20Files/Student_data.xlsx"

# Initialize the main window
root = tk.Tk()
root.title("Student Metrics")
root.geometry('700x500+300+200')
root.resizable(False, False)
root.configure(bg="#117864")

# Function to handle Excel file and check for its existence
def initialize_excel_file():
    try:
        excel_file_content = download_excel_file(xlsx_file_url)
        file = openpyxl.load_workbook(excel_file_content)
        file.save('Student_data.xlsx')  # Save locally to ensure it's available
    except Exception as e:
        print(f"Error downloading or saving the Excel file: {e}")

# Initialize Excel file if it does not exist
initialize_excel_file()

# Function to submit data
def submit():
    name = namevalue.get()
    contact = contactvalue.get()
    age = agevalue.get()
    gender = gender_combobox.get()
    address = addressentry.get(1.0, END)

    try:
        file = openpyxl.load_workbook('Student_data.xlsx')
        sheet = file.active
        sheet.append([name, contact, age, gender, address])
        file.save('Student_data.xlsx')
        messagebox.showinfo('Info', "Detail Added Successfully")
    except Exception as e:
        messagebox.showerror('Error', f"Failed to save data: {e}")

    namevalue.set('')
    contactvalue.set('')
    agevalue.set('')
    addressentry.delete(1.0, END)

# Function to clear form fields
def clear():
    namevalue.set('')
    contactvalue.set('')
    agevalue.set('')
    addressentry.delete(1.0, END)

# Function to load and set image from URL
def load_image_from_url(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  # Check for HTTP errors
        image = Image.open(BytesIO(response.content))
        return ImageTk.PhotoImage(image)
    except Exception as e:
        print(f"Error loading image from URL: {e}")
        return None

# URL for the icon image
icon_image_url = "https://github.com/Ajaykumar2605/Student-Metrics/raw/main/source%20Files/icon/Applogo.png"

# Load and set the icon image
try:
    icon_image = load_image_from_url(icon_image_url)
    if icon_image:
        root.iconphoto(False, icon_image)
except Exception as e:
    print(f"Error setting icon image: {e}")

# Create and place widgets
tk.Label(root, text="Please fill out this Entry form", font='arial', bg="#117864", fg='#fff').place(x=20, y=20)

tk.Label(root, text='Name:', font=23, bg="#117864", fg="#fff").place(x=50, y=100)
tk.Label(root, text='Mobile No:', font=23, bg="#117864", fg="#fff").place(x=50, y=150)
tk.Label(root, text='Age:', font=23, bg="#117864", fg="#fff").place(x=50, y=200)
tk.Label(root, text='Gender:', font=23, bg="#117864", fg="#fff").place(x=370, y=200)
tk.Label(root, text='Address:', font=23, bg="#117864", fg="#fff").place(x=50, y=250)

namevalue = StringVar()
contactvalue = StringVar()
agevalue = StringVar()

nameentry = Entry(root, textvariable=namevalue, width=45, bd=2, font=20)
contactentry = Entry(root, textvariable=contactvalue, width=45, bd=2, font=20)
ageentry = Entry(root, textvariable=agevalue, width=15, bd=2, font=20)

gender_combobox = Combobox(root, values=['Male', 'Female'], font="arial, 14", state='r', width=14)
gender_combobox.place(x=440, y=200)
gender_combobox.set('Male')

addressentry = Text(root, width=50, height=4, bd=4)

nameentry.place(x=200, y=100)
contactentry.place(x=200, y=150)
ageentry.place(x=200, y=200)
addressentry.place(x=200, y=250)

tk.Button(root, text="Submit", bg="#717d7e", fg="White", width=15, height=2, command=submit).place(x=200, y=350)
tk.Button(root, text="Clear", bg="#717d7e", fg="White", width=15, height=2, command=clear).place(x=340, y=350)
tk.Button(root, text="Exit", bg="#717d7e", fg="White", width=15, height=2, command=root.destroy).place(x=480, y=350)

# Copy right logo
footer_frame = tk.Frame(root, pady=10)
footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
copyright_notice = tk.Label(footer_frame, text="© 2024 Ajay Kumar. All rights reserved.", anchor='e')
copyright_notice.pack(side=tk.BOTTOM, padx=10)

# Start the Tkinter event loop
root.mainloop()
