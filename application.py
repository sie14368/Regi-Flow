# RegiFlow - Candidate Registration System
from tkinter import *
from tkinter import ttk, messagebox
import openpyxl
import os
import re
from datetime import datetime

# Application Window

root = Tk()
root.title("RegiFlow - Candidate Registration System")
root.geometry("850x760")
root.resizable(False, False)

frame = Frame(root)
frame.pack(padx=10, pady=10)

# Variables

terms_var = BooleanVar()
disability_var = StringVar(value="No")

# Personal Information

personal_info_frame = LabelFrame(
    frame,
    text="Personal Information",
    font=("Segoe UI", 12, "bold")
)
personal_info_frame.grid(row=0, column=0, padx=5, pady=5, sticky="news")

Label(personal_info_frame, text="Candidate Name").grid(row=0, column=0)
candidate_name_entry = Entry(personal_info_frame)
candidate_name_entry.grid(row=1, column=0)

Label(personal_info_frame, text="Date of Birth").grid(row=0, column=1)
dob_entry = Entry(personal_info_frame)
dob_entry.grid(row=1, column=1)

Label(personal_info_frame, text="Gender").grid(row=0, column=2)
gender_combobox = ttk.Combobox(
    personal_info_frame,
    values=["Male", "Female", "Other"],
    state="readonly"
)
gender_combobox.grid(row=1, column=2)

Label(personal_info_frame, text="Guardian's Name").grid(row=2, column=0)
guardian_name_entry = Entry(personal_info_frame)
guardian_name_entry.grid(row=3, column=0)

Label(personal_info_frame, text="Guardian's Occupation").grid(row=2, column=1)
guardian_occupation_entry = Entry(personal_info_frame)
guardian_occupation_entry.grid(row=3, column=1)

Label(personal_info_frame, text="Caste").grid(row=2, column=2)
caste_combobox = ttk.Combobox(
    personal_info_frame,
    values=["General", "OBC", "SC", "ST"],
    state="readonly"
)
caste_combobox.grid(row=3, column=2)

Label(personal_info_frame, text="Annual Income").grid(row=4, column=0)
annual_income_entry = Entry(personal_info_frame)
annual_income_entry.grid(row=5, column=0)

Label(personal_info_frame, text="Religion").grid(row=4, column=1)
religion_combobox = ttk.Combobox(
    personal_info_frame,
    values=["Hindu", "Muslim", "Christian", "Other"],
    state="readonly"
)
religion_combobox.grid(row=5, column=1)

Label(personal_info_frame, text="Nationality").grid(row=4, column=2)
nationality_combobox = ttk.Combobox(
    personal_info_frame,
    values=[
        "Indian",
        "Chinese",
        "Korean",
        "American",
        "Canadian",
        "British",
        "Australian",
        "German",
        "French",
        "Japanese",
        "Brazilian"
    ],
    state="readonly"
)
nationality_combobox.grid(row=5, column=2)

Label(personal_info_frame, text="Physical Disability").grid(row=6, column=0)

Radiobutton(
    personal_info_frame,
    text="Yes",
    variable=disability_var,
    value="Yes"
).grid(row=7, column=0)

Radiobutton(
    personal_info_frame,
    text="No",
    variable=disability_var,
    value="No"
).grid(row=7, column=1)

for widget in personal_info_frame.winfo_children():
    widget.grid_configure(padx=20, pady=5)

# Contact Details

contact_info_frame = LabelFrame(
    frame,
    text="Contact Details",
    font=("Segoe UI", 12, "bold")
)
contact_info_frame.grid(row=1, column=0, padx=5, pady=5, sticky="news")

Label(contact_info_frame, text="Email ID").grid(row=0, column=0)
email_entry = Entry(contact_info_frame)
email_entry.grid(row=1, column=0)

Label(contact_info_frame, text="Phone Number").grid(row=0, column=1)
phone_entry = Entry(contact_info_frame)
phone_entry.grid(row=1, column=1)

Label(contact_info_frame, text="Additional Phone Number").grid(row=0, column=2)
additional_phone_entry = Entry(contact_info_frame)
additional_phone_entry.grid(row=1, column=2)

Label(contact_info_frame, text="City").grid(row=2, column=0)
city_entry = Entry(contact_info_frame)
city_entry.grid(row=3, column=0)

Label(contact_info_frame, text="State").grid(row=2, column=1)
state_entry = Entry(contact_info_frame)
state_entry.grid(row=3, column=1)

Label(contact_info_frame, text="Pin Code").grid(row=2, column=2)
pin_code_entry = Entry(contact_info_frame)
pin_code_entry.grid(row=3, column=2)

for widget in contact_info_frame.winfo_children():
    widget.grid_configure(padx=20, pady=5)

# Educational Information

educational_info_frame = LabelFrame(
    frame,
    text="Educational Information",
    font=("Segoe UI", 12, "bold")
)
educational_info_frame.grid(row=2, column=0, padx=5, pady=5, sticky="news")

Label(educational_info_frame, text="Secondary Board").grid(row=0, column=0)
secondary_board_combobox = ttk.Combobox(
    educational_info_frame,
    values=["CBSE", "ICSE", "State Board"],
    state="readonly"
)
secondary_board_combobox.grid(row=1, column=0)

Label(educational_info_frame, text="Secondary Percentage").grid(row=0, column=1)
secondary_percentage_spinbox = Spinbox(
    educational_info_frame,
    from_=50,
    to=100
)
secondary_percentage_spinbox.grid(row=1, column=1)

Label(educational_info_frame, text="Higher Secondary Board").grid(row=2, column=0)
higher_secondary_board_combobox = ttk.Combobox(
    educational_info_frame,
    values=["CBSE", "ICSE", "State Board"],
    state="readonly"
)
higher_secondary_board_combobox.grid(row=3, column=0)

Label(educational_info_frame, text="Higher Secondary Percentage").grid(row=2, column=1)
higher_secondary_percentage_spinbox = Spinbox(
    educational_info_frame,
    from_=50,
    to=100
)
higher_secondary_percentage_spinbox.grid(row=3, column=1)

Label(educational_info_frame, text="Preferred Course").grid(row=0, column=2)
course_combobox = ttk.Combobox(
    educational_info_frame,
    values=[
        "Computer Science and Engineering",
        "Information Technology",
        "Electronics and Communication",
        "Electrical Engineering"
    ],
    state="readonly"
)
course_combobox.grid(row=1, column=2)

for widget in educational_info_frame.winfo_children():
    widget.grid_configure(padx=20, pady=5)

# Terms & Conditions

terms_frame = LabelFrame(
    frame,
    text="Terms & Conditions",
    font=("Segoe UI", 12, "bold")
)
terms_frame.grid(row=3, column=0, padx=5, pady=5, sticky="news")

Checkbutton(
    terms_frame,
    text="I agree to the Terms & Conditions.",
    variable=terms_var
).grid(row=0, column=0, padx=10, pady=5, sticky="w")
# Helper Functions

FILE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "data.xlsx"
)


def create_excel_file():
    """Create the Excel workbook if it does not already exist."""

    if os.path.exists(FILE_PATH):
        return

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Candidates"

    headings = [
        "Registration ID",
        "Timestamp",
        "Candidate Name",
        "DOB",
        "Gender",
        "Guardian's Name",
        "Guardian's Occupation",
        "Caste",
        "Annual Income",
        "Religion",
        "Nationality",
        "Disability Status",
        "Email ID",
        "Phone Number",
        "Additional Phone Number",
        "City",
        "State",
        "Pin Code",
        "Secondary Board",
        "Secondary Percentage",
        "Higher Secondary Board",
        "Higher Secondary Percentage",
        "Preferred Course"
    ]

    sheet.append(headings)

    workbook.save(FILE_PATH)


def generate_registration_id():
    """Generate Registration IDs like RGF001, RGF002..."""

    workbook = openpyxl.load_workbook(FILE_PATH)
    sheet = workbook.active

    row_count = sheet.max_row

    workbook.close()

    return f"RGF{row_count:03d}"


def is_duplicate_email(email):
    """Return True if email already exists."""

    workbook = openpyxl.load_workbook(FILE_PATH)
    sheet = workbook.active

    for row in sheet.iter_rows(min_row=2, values_only=True):

        stored_email = row[12]

        if stored_email is not None:

            if stored_email.lower() == email.lower():

                workbook.close()

                return True

    workbook.close()

    return False


def clear_form():
    """Reset the form after successful registration."""

    candidate_name_entry.delete(0, END)
    dob_entry.delete(0, END)

    gender_combobox.set("")
    caste_combobox.set("")
    religion_combobox.set("")
    nationality_combobox.set("")

    guardian_name_entry.delete(0, END)
    guardian_occupation_entry.delete(0, END)
    annual_income_entry.delete(0, END)

    disability_var.set("No")

    email_entry.delete(0, END)
    phone_entry.delete(0, END)
    additional_phone_entry.delete(0, END)

    city_entry.delete(0, END)
    state_entry.delete(0, END)
    pin_code_entry.delete(0, END)

    secondary_board_combobox.set("")
    higher_secondary_board_combobox.set("")
    course_combobox.set("")

    secondary_percentage_spinbox.delete(0, END)
    secondary_percentage_spinbox.insert(0, "50")

    higher_secondary_percentage_spinbox.delete(0, END)
    higher_secondary_percentage_spinbox.insert(0, "50")

    terms_var.set(False)


def validate_form():
    """Validate all user inputs."""

    if not terms_var.get():
        messagebox.showwarning(
            "Terms & Conditions",
            "Please accept the Terms & Conditions."
        )
        return False

    required_fields = [
        (candidate_name_entry.get(), "Candidate Name"),
        (dob_entry.get(), "Date of Birth"),
        (gender_combobox.get(), "Gender"),
        (email_entry.get(), "Email ID"),
        (phone_entry.get(), "Phone Number"),
        (city_entry.get(), "City"),
        (state_entry.get(), "State"),
        (pin_code_entry.get(), "Pin Code"),
        (course_combobox.get(), "Preferred Course")
    ]

    for value, field in required_fields:

        if value.strip() == "":
            messagebox.showerror(
                "Missing Information",
                f"{field} cannot be empty."
            )
            return False

    email = email_entry.get().strip()

    email_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'

    if not re.match(email_pattern, email):

        messagebox.showerror(
            "Invalid Email",
            "Please enter a valid email address."
        )

        return False

    phone = phone_entry.get().strip()

    if not phone.isdigit() or len(phone) != 10:

        messagebox.showerror(
            "Invalid Phone Number",
            "Phone number must contain exactly 10 digits."
        )

        return False

    pin = pin_code_entry.get().strip()

    if not pin.isdigit() or len(pin) != 6:

        messagebox.showerror(
            "Invalid PIN Code",
            "PIN Code must contain exactly 6 digits."
        )

        return False

    create_excel_file()

    if is_duplicate_email(email):

        messagebox.showerror(
            "Duplicate Registration",
            "This email is already registered."
        )

        return False

    return True

# Submit Form

def submit_form():

    if not validate_form():
        return

    registration_id = generate_registration_id()

    timestamp = datetime.now().strftime("%d-%m-%Y %I:%M:%S %p")

    workbook = openpyxl.load_workbook(FILE_PATH)
    sheet = workbook.active

    candidate_data = [

        registration_id,
        timestamp,

        candidate_name_entry.get().strip(),
        dob_entry.get().strip(),
        gender_combobox.get(),

        guardian_name_entry.get().strip(),
        guardian_occupation_entry.get().strip(),

        caste_combobox.get(),
        annual_income_entry.get().strip(),

        religion_combobox.get(),
        nationality_combobox.get(),

        disability_var.get(),

        email_entry.get().strip(),
        phone_entry.get().strip(),
        additional_phone_entry.get().strip(),

        city_entry.get().strip(),
        state_entry.get().strip(),
        pin_code_entry.get().strip(),

        secondary_board_combobox.get(),
        secondary_percentage_spinbox.get(),

        higher_secondary_board_combobox.get(),
        higher_secondary_percentage_spinbox.get(),

        course_combobox.get()
    ]

    sheet.append(candidate_data)

    workbook.save(FILE_PATH)
    workbook.close()

    messagebox.showinfo(
        "Registration Successful",
        f"Candidate registered successfully!\n\nRegistration ID: {registration_id}"
    )

    clear_form()


# Submit Button

submit_button = Button(
    frame,
    text="Submit Registration",
    command=submit_form,
    font=("Segoe UI", 11, "bold"),
    width=25,
    bg="#4CAF50",
    fg="white",
    cursor="hand2"
)

submit_button.grid(row=4, column=0, pady=15)


# ============================================================
# Run Application
# ============================================================

root.mainloop()
