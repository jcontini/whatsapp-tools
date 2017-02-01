## First Install
1. Make sure Python is installed.
    - If you have OSX, Python is already has installed.
    - If you have windows, install Python from https://www.python.org/downloads/release/python-2713/
2. Download or clone this repository
3. Install required modules using pip (any directory)
    - `pip install openpyxl`

## Excelify
Input: A WhatsApp conversation (txt) from 'Email chat'
Output: An Excel file with messages split into rows with columns for Type, Datetime, Sender, and Message.

1. Send a WhatsApp conversation to yourself and get the .txt file.
2. Put the .txt in the same folder as the script
3. Run the script like this:
    - `python excelify.py WhatsApp\ Chat\ with\ Joe.txt`