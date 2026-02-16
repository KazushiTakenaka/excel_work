---
name: Excel & Sheets Environment Setup
description: Automates the setup of a Python virtual environment and installs necessary libraries for Excel and Google Sheets operations.
---

# Excel & Google Sheets Environment Setup Skill

Use this skill when the user wants to set up a new project environment for Excel or Google Sheets automation.
This skill eliminates the need to manually copy `setup.bat` or `requirements.txt`.

## Prerequisites
- The user must be in the target directory where they want to create the environment.
- Python must be installed on the user's system.

## Workflow

1.  **Check Python Installation**
    - Run: `python --version`
    - If it fails, inform the user they need to install Python first.

2.  **Create Virtual Environment (venv)**
    - Check if a directory named `venv` exists.
    - If not, run: `python -m venv venv`
    - Wait for the command to complete.

3.  **Install Dependencies**
    - Install the required libraries into the virtual environment.
    - Run the following command (using the venv's pip directly):
      ```powershell
      .\venv\Scripts\pip install gspread oauth2client pandas openpyxl
      ```

4.  **Verification**
    - Run: `.\venv\Scripts\pip list`
    - Confirm that `gspread`, `pandas`, and `openpyxl` are listed.

5.  **Completion Message**
    - Notify the user that the environment is ready.
    - Remind them to place their `credentials.json` in this folder.
