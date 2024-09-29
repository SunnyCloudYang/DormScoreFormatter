# Dorm Score Formatter

## Description

DormScoreFormatter.py is a Python script designed to process and format dorm inspection scores. It takes multiple CSV files containing weekly dorm scores and combines them into a single, formatted Excel file and a PDF file. This tool is particularly useful for dorm administrators or student organizations managing dorm inspections.

## Features

- Combines multiple CSV files with weekly dorm scores
- Processes and formats the data
- Creates a well-structured Excel file with:
  - A title including the dorm number and week
  - Contact information for inquiries
  - Neatly organized scores and comments
  - Automatic pagination for easy printing
- Export PDF files for easy printing

## Requirements

To run this script, you need:

1. Python 3.6 or higher
2. The following Python libraries:
   - pandas
   - openpyxl
   - win32com.client (only required for exporting PDFs on Windows)

You can install these libraries using pip:

```
pip install pandas openpyxl pywin32
```

## How to Use

1. Place all your WeekScoreManage_*.csv files in a single folder.

2. Open a command prompt or terminal.

3. Navigate to the folder containing the DormScoreFormatter.py script.

4. Run the script with the following command to view the script's help information and optional arguments:

   ```
   python DormScoreFormatter.py --help
   ```

5. Run the script using the following command:

   ```
   python DormScoreFormatter.py --folder path/to/your/csv/files
   ```

   Replace `path/to/your/csv/files` with the actual path to the folder containing your CSV files.

6. The script will process the files and create an Excel file named after the dorm number and week (e.g., "紫荆公寓2号楼第1周.xlsx") in the same folder as the script.

## Output

The resulting Excel file will contain:

- A title with the dorm number and week
- Contact information for inquiries
- Formatted tables with room numbers, bed numbers, total scores, and improvement comments

If there are blank cells in the score table, the position of the blank cell will be printed in the console for easy checking and the cell will be filled with red in the Excel file.

## Troubleshooting

If you encounter any issues:

1. Ensure you have Python and all required libraries installed.
2. Check that your CSV files are in the correct format and named properly (starting with 'WeekScoreManage_').
3. Make sure you're providing the correct path to the folder containing your CSV files.

For any other issues, please [contact me](mailto:sunnycloudyang@outlook.com).

## Note for New Python Users

If you're new to Python, you might need to set up your Python environment first. Here are some steps to get started:

1. Download and install Python from [python.org](https://www.python.org/downloads/).
2. During installation, make sure to check the box that says "Add Python to PATH".
3. After installation, open a command prompt or terminal and type `python --version` to verify that Python is installed correctly.
4. Use the pip command mentioned earlier to install the required libraries.

Once you've set up Python and installed the required libraries, you should be able to run the script as described in the "How to Use" section.
