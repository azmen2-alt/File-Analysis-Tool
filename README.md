# FBCRS Record Inventory & File Analysis Tools

## Overview

This repository contains a suite of Python utilities designed to function as a full-featured Information Governance and Records Management engine. The primary application scans local or network directories, reads various document types, automatically categorizes them based on record retention rules, and generates a formatted Excel summary report.

## Core Components

### 1. File Analysis Engine

The main application is designed to handle massive file directories with built-in crash resilience. It is available in two distinct processing modes:

* 
**Slow but Accurate Version:** Processes complete files and generates full hash fingerprints for absolute duplicate validation.


* 
**Fast but Less Accurate Version:** Optimized to scan large data drops quickly. It achieves this by reading only the first two pages of PDFs, the first two slides of PowerPoints, the first 50 paragraphs of Word documents, the first Excel sheet, and only the first 1 MB of a file for duplicate checking .



**Key Features:**

* 
**High Performance:** Utilizes `ThreadPoolExecutor` for concurrent I/O processing and the `pyahocorasick` automaton to reduce algorithmic classification complexity, significantly speeding up large record volume scans.


* 
**Crash Recovery (Checkpointing):** Features an incremental saving design that writes every processed row to a temporary CSV file (`_temp.csv`) in real-time. If the script crashes, data processed up to that exact moment is safely retained and automatically combined upon a successful restart.


* 
**Graceful Error Handling:** Includes `try...except` blocks around metadata retrieval to handle inaccessible files, corrupted paths, or Windows UNC network paths that exceed the 260-character limit. The app logs an "Error" row and proceeds rather than terminating the process.


* 
**Standardized Output:** Integrates Data Validation in the final Excel output, enforcing standardized dropdown menus for downstream auditing and applying clickable file hyperlinks.



### 2. Master Code Updater

A supplementary utility designed to safely update the live rulebook, `FBCRS_Master_Full.xlsx`, with missing Business Intelligence FBCRS codes (such as STR-PLA-002, STR-REP-001, etc.).

* 
**Dynamic Column Matching:** It searches the Excel file for the first column's actual name to map data correctly, preventing duplication.


* 
**Safe Updating:** Utilizes the `.update()` function to inject new BI descriptions for existing codes or append new rows securely at the bottom. This ensures that existing, unrelated columns (like "Notes" or "Owner") are left completely untouched.



---

## Installation & Setup

If you are running the tools directly via Python, ensure you have **Python 3.10 or higher** installed and added to your system PATH.

Install all necessary dependencies via your terminal or command prompt:

```bash
pip install pandas openpyxl pypdf python-docx python-pptx pyahocorasick

```

* 
`openpyxl`: Required for reading/writing Excel files (.xlsx).


* 
`pypdf`, `python-docx`, `python-pptx`: Used for content extraction across various document types .


* 
`pyahocorasick`: Highly recommended for the high-performance search algorithm.


* 
`pandas`: Required for merging and updating the FBCRS code data.



---

## Deployment (Standalone Executable)

To distribute the File Analysis Engine to end-users without requiring them to install Python or external libraries, the scripts can be bundled into a standalone Windows executable (`.exe`) .

**Build Instructions:**

1. Install PyInstaller:
```bash
python -m pip install pyinstaller

```



*(Note: Ensure you use `python -m` if Windows does not recognize the raw `pip` command)*.


2. Run the PyInstaller command directly in your terminal (do not place this command inside your `.py` file to avoid syntax errors). For example:


```bash
python -m PyInstaller --onefile "C:\Path\To\Your\Script.py"

```


3. Once the terminal reports "completed successfully," the ready-to-use `.exe` file will be generated (either inside a `dist` folder or on your Desktop, depending on your `--distpath` flag).



### Usage Notes for End-Users

* The `.exe` application will display a black terminal screen during operation; this is the app's engine and must remain open while running.


* 
**Crucial Rulebook Requirement:** The `FBCRS_Master_Full.xlsx` library file must always be kept in the exact same folder as the `.exe` application. Because it sits adjacent to the app, users can update keywords or retention rules in the Excel file at any time, and the engine will instantly apply the new rules on the next run.
