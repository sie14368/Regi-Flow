from tkinter import *
from tkinter import ttk, messagebox
import openpyxl
import os
import re

# Main Window

root = Tk()
root.title("RegiFlow")
root.geometry("600x500")
root.resizable(False, False)


# Excel File

FILE_NAME = "data.xlsx"

# Form Frame

frame = Frame(root, padx=20, pady=20)
frame.pack()

# Candidate Information


Label(frame, text="Candidate Name").grid(row=0, column=0, sticky="w")
name_entry = Entry(frame, width=30)
name_entry.grid(row=0, column=1, pady=5)

Label(frame, text="Date of Birth").grid(row=1, column=0, sticky="w")
dob_entry = Entry(frame, width=30)
dob_entry.grid(row=1, column=1, pady=5)

Label(frame, text="Gender").grid(row=2, column=0, sticky="w")
gender_combo = ttk.Combobox(
    frame,
    values=["Male", "Female", "Other"],
    state="readonly",
    width=27
)
gender_combo.grid(row=2, column=1, pady=5)

Label(frame, text="Email").grid(row=3, column=0, sticky="w")
email_entry = Entry(frame, width=30)
email_entry.grid(row=3, column=1, pady=5)

Label(frame, text="Phone Number").grid(row=4, column=0, sticky="w")
phone_entry = Entry(frame, width=30)
phone_entry.grid(row=4, column=1, pady=5)

Label(frame, text="City").grid(row=5, column=0, sticky="w")
city_entry = Entry(frame, width=30)
city_entry.grid(row=5, column=1, pady=5)

Label(frame, text="State").grid(row=6, column=0, sticky="w")
state_entry = Entry(frame, width=30)
state_entry.grid(row=6, column=1, pady=5)

Label(frame, text="Preferred Course").grid(row=7, column=0, sticky="w")
course_combo = ttk.Combobox(
    frame,
    values=[
        "Computer Science",
        "Information Technology",
        "Electronics",
        "Electrical"
    ],
    state="readonly",
    width=27
)
course_combo.grid(row=7, column=1, pady=5)


# Create Excel File


def create_excel():
    if os.path.exists(FILE_NAME):
        return

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet.append([
        "Name",
        "DOB",
        "Gender",
        "Email",
        "Phone",
        "City",
        "State",
        "Course"
    ])

    workbook.save(FILE_NAME)
    workbook.close()

# Validate Form

def validate():

    # Check empty fields

    if name_entry.get().strip() == "":
        messagebox.showerror("Error", "Enter candidate name.")
        return False

    if dob_entry.get().strip() == "":
        messagebox.showerror("Error", "Enter date of birth.")
        return False

    if gender_combo.get() == "":
        messagebox.showerror("Error", "Select gender.")
        return False

    if email_entry.get().strip() == "":
        messagebox.showerror("Error", "Enter email.")
        return False

    if phone_entry.get().strip() == "":
        messagebox.showerror("Error", "Enter phone number.")
        return False

    if city_entry.get().strip() == "":
        messagebox.showerror("Error", "Enter city.")
        return False

    if state_entry.get().strip() == "":
        messagebox.showerror("Error", "Enter state.")
        return False

    if course_combo.get() == "":
        messagebox.showerror("Error", "Select course.")
        return False

    # Email Validation

    email = email_entry.get().strip()

    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'

    if not re.match(pattern, email):
        messagebox.showerror("Error", "Invalid email address.")
        return False


    # Phone Validation

    phone = phone_entry.get().strip()

    if not phone.isdigit() or len(phone) != 10:
        messagebox.showerror(
            "Error",
            "Phone number must contain 10 digits."
        )
        return False

    return True

# Clear Form

def clear_form():
    name_entry.delete(0, END)
    dob_entry.delete(0, END)

    gender_combo.set("")

    email_entry.delete(0, END)
    phone_entry.delete(0, END)

    city_entry.delete(0, END)
    state_entry.delete(0, END)

    course_combo.set("")

# Submit Form

def submit():

    if not validate():
        return

    create_excel()

    workbook = openpyxl.load_workbook(FILE_NAME)
    sheet = workbook.active

    sheet.append([
        name_entry.get().strip(),
        dob_entry.get().strip(),
        gender_combo.get(),
        email_entry.get().strip(),
        phone_entry.get().strip(),
        city_entry.get().strip(),
        state_entry.get().strip(),
        course_combo.get()
    ])

    workbook.save(FILE_NAME)
    workbook.close()

    messagebox.showinfo(
        "Success",
        "Candidate registered successfully!"
    )

    clear_form()

# Submit Button

submit_button = Button(
    frame,
    text="Submit",
    command=submit,
    width=20,
    bg="green",
    fg="white"
)

submit_button.grid(
    row=8,
    column=0,
    columnspan=2,
    pady=20
)

# Run Application

root.mainloop()
