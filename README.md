# Employee Attendance Processing System

A web-based UI for processing employee attendance data and generating reports.

## Features

- Upload employee attendance CSV files
- Generate missing records report (employees who didn't check in/out)
- Generate summary report by department
- Download processed reports in Excel format

## Requirements

- Python 3.7 or higher
- Flask
- Pandas
- OpenPyXL

## Installation

1. Clone or download this repository
2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```
   python app.py
   ```
   
   Or for the demo version (without processing):
   ```
   python app_demo.py
   ```

2. Open your web browser and go to `http://localhost:5000`

3. Upload the required CSV files:
   - Employee file
   - First & Last report
   - Leave Report

4. Click "Process Files" to generate reports

5. Download the generated reports

## Note

If you encounter issues with dependencies, you can run the demo version (`app_demo.py`) which provides the UI without the actual processing functionality.

Alternatively, you can view the HTML prototype (`ui_prototype.html`) which demonstrates the UI in a static HTML file that can be opened directly in a browser.

## File Structure

- `app.py` - Main Flask application
- `templates/` - HTML templates
- `uploads/` - Temporary folder for uploaded files
- `outputs/` - Folder for generated reports
- `summary/` - Contains summary report template and processing script

## How It Works

The application processes three types of reports:

1. **Missing Records Report**: Identifies employees who didn't check in/out and were not on leave
2. **Summary Report**: Aggregates attendance data by department

## Customization

You can modify the department mappings in `app.py` to match your organization's department names.