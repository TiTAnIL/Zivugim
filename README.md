# Zivugim
read data from an Excel file, organize content by program names and attribute and save the structured output to a new Excel file. The script uses the openpyxl library for Excel manipulation and tkinter for the GUI. Additionally, it colors cells based on certain criteria and logs key program information.
Zivugim - Excel Data Processor

Zivugim is a Python application that processes Excel files, categorizing and organizing data based on program attributes like HD, DUB, SC, and others. This tool simplifies data handling for large lists by organizing entries, generating an output file with structured data, and visually marking entries based on specific attributes.
Features

    Load Excel Files: Easily load .xlsx files for processing.
    Organize Program Data: Collects data by program attributes (HD, SC, DUB, etc.) and organizes them in a structured dictionary format.
    Output File Generation: Creates a new Excel file containing the organized data for each program, arranged by attributes.
    Row Coloring: Visually distinguishes entries by highlighting rows based on program type (e.g., MASTER and SLAVE).
    Completion Message: Shows a random completion message upon finishing the processing.

Requirements

    Python 3.x
    Required Libraries:
        openpyxl: For handling Excel file operations.
        tkinter: For the graphical interface (usually included with Python).
        datetime, os, re, shutil: Standard libraries for file and time management.

Installation

    Clone the repository:

    bash

git clone https://github.com/yourusername/zivugim-excel-processor.git
cd zivugim-excel-processor

Install the required package:

bash

pip install openpyxl

Run the application:

bash

    python zivugim.py

Usage

    Run Application: Launch the tool by running python zivugim.py.
    Load Excel File:
        Click Load Magic File in the GUI to select an Excel file (.xlsx format) for processing.
    Process Data:
        Click Run to start processing.
        The application will:
            Identify program attributes.
            Organize data into a new output Excel file.
            Log key information.
            Highlight rows based on attribute matches (e.g., MASTER, SLAVE).
    View Results:
        The output Excel file is saved in the same directory as the script with a filename that includes the current date.
        Rows are highlighted to help visually distinguish data based on attributes.

Example

    Load File: Click Load Magic File to select a .xlsx file with program data.
    Run: Press Run to organize data by attributes, generate a structured output file, and highlight specific rows.
    Output: Check the output directory for the new file and view organized entries with color-coded rows.

Completion Message

After processing, a random message from the pre-defined CoolExit list will be displayed as a fun way to confirm completion.
