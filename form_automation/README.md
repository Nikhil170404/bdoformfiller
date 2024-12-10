# Form Automation Tool

This tool helps you automate web form filling by first analyzing the form fields and then filling them with data from an Excel spreadsheet.

## Features

1. Form Analysis
   - Automatically detects form fields (inputs, selects, textareas)
   - Creates an Excel template with all form fields as columns
   - Supports required fields marking
   - Handles dropdowns with predefined options

2. Form Filling
   - Fills forms using data from Excel spreadsheet
   - One row = One form submission
   - Supports multiple form submissions
   - Handles different field types (text, select, checkbox, radio)

## Setup

1. Install Python 3.11 or later
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Step 1: Analyze Form and Create Template

1. Run the form analyzer:
   ```bash
   python form_analyzer.py
   ```
2. Enter the URL of the form when prompted
3. The script will:
   - Analyze all form fields
   - Create an Excel template file (`form_template.xlsx`)
   - List all detected fields

### Step 2: Fill the Excel Template

1. Open `form_template.xlsx`
2. Each column represents a form field
3. Each row represents one form submission
4. Fill in the data for as many users as needed
5. Save the file

### Step 3: Run the Form Filler

1. Run the form filler:
   ```bash
   python form_filler.py
   ```
2. Enter the form URL and Excel file path when prompted
3. The script will:
   - Process each row in the Excel file
   - Fill and submit the form for each entry
   - Create a log file with results

## Troubleshooting

- Check `form_automation.log` for detailed error messages
- Make sure you have Chrome browser installed
- Verify that the form URL is accessible
- Ensure all required fields in Excel are filled
