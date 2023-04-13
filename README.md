# Archimedes

Archimedes is a management tool designed to automate the creation and processing of pattern work orders, pattern receipts, shop copies, and other SEFCOR engineering files. This tool helps engineers to create and manage these documents more efficiently while reducing the chances of errors.

## Features

- GUI for creating and managing Pattern Work Orders (PWOs), Pattern Receipts, and Shop Copies
- Automatically generate and send email notifications with attached PDF files
- Generate Word and PDF documents based on user input and store them in appropriate folders
- Update and manage Engineering Log Book (Excel spreadsheet)
- Integration with JIRA for issue tracking and management

## Installation

1. Clone the repository to your local machine.
2. Ensure you have Python 3.6 or higher installed.
3. Install the required Python packages by running the following command in your terminal or command prompt:

```bash
pip install -r requirements.txt

## Usage

1. Open your terminal or command prompt and navigate to the project directory.
2. Run the following command to start the application:
'''bash
python main.py
3. The application GUI will appear. Fill in the required fields and click the relevant buttons to create the desired document.

## Dependencies

This project relies on several external Python packages. The required packages are:
- comtypes
- openpyxl
- jira
- pdf2image
- img2pdf
- pillow
- pytesseract

It is also required to install tesseract-ocr on system PATH. (<- needs expansion)

## Configuration

A configuration file 'config.json' is used to store file paths, email addresses, auth tokens, and other settings. Make sure to update the file paths in this configuration file to match your local setup. A template is include in the repo: 'config.template.json'. NEVER COMMIT YOUR WORKING CONFIG FILE TO GIT. It is best to add your configuration file to .gitignore.

## Contributing

Feel free to submit issues and enhancement requests or open pull requests with improvements.
