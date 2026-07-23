# RegiFlow - Candidate Registration System

## About the Project

RegiFlow is a simple Candidate Registration System developed using Python, Tkinter, and OpenPyXL. It allows users to enter candidate details through a graphical interface and stores the information in an Excel file. The project was created to practice GUI development, file handling, and form validation in Python.

## Features

* User-friendly registration form
* Stores candidate details in an Excel file
* Basic input validation
* Checks for valid email format
* Validates 10-digit phone numbers
* Automatically creates the Excel file if it does not exist
* Clears the form after successful registration

## Technologies Used

* Python
* Tkinter
* OpenPyXL

## Project Structure

```
RegiFlow/
│
├── main.py
├── data.xlsx
└── README.md
```

## How to Run

1. Install Python on your system.
2. Install the required library:

```
pip install openpyxl
```

3. Run the program:

```
python main.py
```

The application window will open, and all submitted records will be stored in `data.xlsx`.

## Learning Outcomes

Through this project, I learned how to:

* Build graphical user interfaces using Tkinter.
* Handle user input and perform basic validation.
* Read from and write to Excel files using OpenPyXL.
* Organize Python code using functions for better readability.

## Future Improvements

Some features that can be added in the future include:

* Search and edit candidate records.
* Delete existing registrations.
* Store data in a database instead of Excel.
* Add login authentication for administrators.
