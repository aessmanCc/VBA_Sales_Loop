Project Title: VBA Data Extraction and Formatting Tool
Overview:
This project provides a VBA-based solution for extracting data from PDF reports converted into Excel format (.xlsx) and formatting it for further analysis. The process involves importing employee information, pulling data from multiple spreadsheets, and generating a formatted output for analysis.

Workflow:

1. Data Source: The data is sourced from PDF reports converted into Excel format using external software.

2. File Organization: Batch files are utilized to organize files in a central "Holding" folder.

3. VBA Macro Execution: The main VBA macro (xxxxTotalxxxxx.bas) orchestrates the data extraction and formatting process. It involves calling several subroutines for specific tasks:
empMaster: Generates mock employee data.

- Ost_New: Extracts data from the first Excel report.

- ImportCref, ImportRec, ImportDelPro, ImportCben: Imports subsequent Excel reports.

- ImportTemplate: Imports a template for data presentation.

- FormatData: Cleans up and formats the extracted data.

4. Final Output: The formatted data is saved as an Excel file with a timestamp in the specified directory.


Instructions:
1. Setup: Ensure the necessary files are organized in the specified directory structure.
2. Execution: Run the batch file to pull files into the central "Holding" folder. Then execute the VBA macro xxxxTotalxxxxx.bas.
3. Output: The formatted data will be saved as an Excel file in the specified location.

Usage:
This tool can be used for automating the extraction and formatting of data from PDF reports into a structured format for further analysis.
It can be customized and integrated into existing workflows for data processing and analysis tasks.
