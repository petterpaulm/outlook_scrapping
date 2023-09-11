# Extract and Compile Excel Attachments from Microsoft Outlook

A Python script to automatically extract Excel attachments from specific Microsoft Outlook emails and compile them into a single pandas DataFrame.

## 🚀 Features

- Extract Excel attachments from Outlook emails based on a keyword in the subject line
- Compile all extracted Excel data into a single Pandas DataFrame

## 💼 Prerequisites

- Python
- Microsoft Outlook installed and configured on your system
- Python packages: `pywin32`, `pandas`

## 💡 Usage

1. Install the necessary Python packages by running these commands in your terminal or command prompt:

   ```sh
   pip install pywin32
   pip install pandas
   ```
2. Edit the `subject_keyword` variable in `email_scrapping.py` to the keyword you wish to search in the email subjects.
3. Run `email_scrapping.py`

## 💼 Usage

1. **Configuration:** Open the `email_scrapping.py` script and set the `subject_keyword` variable to the keyword you wish to use for filtering emails based on their subject lines.
   
2. **Running the Script:** Run the `email_scrapping.py` script in a Python environment. You can do this from a terminal or an integrated development environment (IDE) that supports Python.
   
3. **Output:** The script will automatically save the Excel attachments to a folder (default: 'attachments'). It will then read these Excel files and compile the data into a single pandas DataFrame, which will be displayed in the console.

## 📁 Files Included

- `email_scrapping.py`: The primary Python script that contains the code to extract Excel attachments from Outlook emails and compile them into a single pandas DataFrame.
- `README.md`: The markdown file you are reading now, offering an overview of the project and instructions for usage.
