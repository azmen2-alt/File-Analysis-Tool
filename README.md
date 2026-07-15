File Analysis and Organization Tools
Overview
These tools help you organize and track your files. The main program looks through your computer or network folders, reads different types of documents, automatically sorts them based on your rules, and creates a neat Excel summary report.

Core Components
1. File Analysis App
The main app is built to handle massive folders without crashing. It comes in two versions:


Slow but Accurate Version: Reads every single file from start to finish and checks a unique digital signature to guarantee it finds exact duplicates.


Fast but Less Accurate Version: Speeds through large folders. It runs faster by only reading the first two pages of PDFs, the first two slides of PowerPoints, the first 50 paragraphs of Word documents, the first Excel sheet, and just the first part of a file to check for duplicates.

Key Features:


High Speed: Scans multiple files at the same time and uses a fast search method to find keywords quickly.


Crash Protection: Saves your progress row-by-row into a temporary file (_temp.csv) as it works. If the app closes unexpectedly, your progress is saved, and it will combine everything for you the next time it finishes.


Skips Errors: If it finds a corrupted file, a locked document, or a file path that is too long for Windows, it simply logs it as an "Error" and moves to the next file instead of crashing.


Clean Excel Output: The final Excel report includes clickable links to your files and dropdown menus to keep your data clean and easy to review.

2. Excel Rulebook Updater
A helper tool that safely adds missing classification codes to your main rulebook (FBCRS_Master_Full.xlsx).


Smart Matching: It automatically finds the right column name to prevent duplicating data.


Safe Updating: It safely adds new information or new rows to the bottom without altering your existing notes or other columns.

Setup (For Python Users)
If you are running the code directly, you need Python 3.10 or newer. Install the required add-ons using your command prompt:


openpyxl: For reading and writing Excel files.


pypdf, python-docx, python-pptx: For reading PDFs, Word, and PowerPoint files.


pyahocorasick: Highly recommended to make the keyword search much faster.


pandas: Used for updating the rulebook.

Ready-to-Use Windows App (For Regular Users)
You can turn the code into a standard Windows app (.exe) so others can use it without having to install Python.

How to Build the App:

Install the building tool: python -m pip install pyinstaller.

Run the build command in your terminal (do not put this inside your Python file): python -m PyInstaller --onefile "C:\Path\To\Your\Script.py".

When it says "completed successfully," you will find your ready-to-use .exe app in a folder named dist (or on your Desktop).

Notes for the User
The app opens a black screen while it runs. This is normal and must stay open for the app to work.


Crucial Rule: The main rulebook file (FBCRS_Master_Full.xlsx) must stay in the exact same folder as the app. Because the files sit next to each other, you can easily change the rules or keywords in Excel at any time, and the app will instantly use those new rules the next time you run it.
