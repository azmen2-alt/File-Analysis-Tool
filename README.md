# File Analysis Engine

## Overview
The File Analysis Engine is a desktop application built to automate the classification and indexing of document records for the Business Intelligence Unit. By scanning local directories and reading file contents, this tool maps documents to official government retention codes based on a master rulebook. It recently helped reduce an indexing project of 41,000 records from three months to under three days.

## Key Features
* **Multi-Format Support:** Reads text from Word, Excel, PDF, PowerPoint, and TXT files.
* **Automated Classification:** Uses a custom Excel rulebook to score and assign retention codes.
* **Duplicate Detection:** Checks file hashes to find and flag exact duplicate files.
* **Filter System:** Automatically skips unsupported media formats, empty files, and temporary system files.
* **Excel Reporting:** Generates a complete summary report with confidence levels, retention expiry dates, and direct file links.
* **User Interface:** A simple dark-mode interface with a real-time progress bar and analytics dashboard.

## Required Files
To run this application, the following two files must be in the exact same folder:
1. **`Record_Inventory_Engine.exe`** (or the `.py` script) - The main application.
2. **`the library.xlsx`** - The master rulebook containing the tabs for Records, Keywords, Phrases, and Synonyms. *(Note: This file must remain closed while the app is running).*

## How to Use
1. Open the application.
2. Click **Browse Folder** and select the directory you want to scan.
3. Click **Start Analysis**. 
4. Wait for the progress bar to complete. Break large folders into smaller sub-folders for better performance.
5. When finished, an analytics dashboard will appear. A new Excel summary report will be saved in the same folder as the application.

## Troubleshooting
* **App closes instantly:** Ensure `the library.xlsx` is saved in the exact same folder as the application.
* **Rulebook Read Error:** Make sure `the library.xlsx` is not open in Microsoft Excel when you click "Start Analysis".
* **Permission Errors:** If you hit network security blocks, move the files you want to scan to your local computer (like your Desktop) and run the scan there.

## How the Code Works
1. **File Reading:** The app opens each file in the folder you select. It reads the text inside Word, Excel, PDF, PowerPoint, and text files.
2. **Keyword Scanning:** It checks the document text against your master rulebook (`the library.xlsx`). It searches for specific keywords, phrases, and synonyms.
3. **Scoring:** Every time it finds a matching word, it adds points to a specific government retention code.
4. **Classification:** The code with the highest score wins. The app assigns this code to the document.
5. **Reporting:** It saves all results in a final Excel report. This report shows the assigned code, the confidence score, and a direct link to open the file.

## Python Libraries Used
The application is built with Python and uses these main libraries to do the heavy lifting:
* **pandas:** Organizes the data and handles the final output.
* **openpyxl:** Reads your `the library.xlsx` rulebook.
* **pypdf:** Extracts text from PDF files.
* **python-docx:** Extracts text from Word documents.
* **python-pptx:** Extracts text from PowerPoint presentations.
* **pyahocorasick:** A fast search tool that looks for multiple keywords at the same time without slowing down your computer.
* **customtkinter:** Builds the visual user interface, including the dark-mode window, buttons, and progress bar.

## About the Rulebook and Codes
The engine cannot work without `the library.xlsx`. This rulebook connects plain text to official Business Intelligence Unit retention rules. 

The rulebook is broken down into specific tabs:
* **Records Tab:** Lists the main government codes, categories, and retention timeframes.
* **Keywords & Phrases Tabs:** Lists the exact words the application should search for. 

When the app reads a document, it uses these tabs to figure out what the document is about. It then links the file to the correct official code so the record is kept or destroyed on the correct date.
