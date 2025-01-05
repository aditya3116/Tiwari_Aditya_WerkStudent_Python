
## Introduction
This Python application extracts financial data from PDF invoices and organizes it into Excel and CSV formats, facilitating easy analysis and record-keeping.

## Project Structure
```
.
├── extract.py           # Main script for data extraction
├── executable           # Compiled version of the script for non-Python environments
└── invoices             # Directory containing sample PDF invoices
```

## Features

### PDF Parsing
The application utilizes the PyMuPDF library (`fitz`) to parse PDF documents. It employs regular expressions to detect and extract key information:
- **Amounts**: Identifies monetary values following specific keywords.
- **Dates**: Extracts date formations related to invoice issuance.

### Data Compilation

#### Creating an Excel File
The script writes the extracted data into an Excel file with two distinct sheets:
- **Sheet 1 (DataSheet)**: Contains a tabular representation with columns for FileName, Date, and Amount, listing detailed information extracted from each invoice.
- **Sheet 2 (Pivot Table)**: A pivot table summarizes the data from Sheet 1 by date, using the sum function to aggregate amounts. This setup helps in analyzing financial data over specific periods.

#### Creating a CSV File
Following the data compilation in Excel, the script converts the workbook into a CSV file:
- **Delimiter**: Configured to use a semicolon (`;`) as the delimiter to suit regional settings where commas serve as decimal points.

### Executable Creation
The Python script can be compiled into an executable using PyInstaller. This executable runs independently of a Python installation, making it practical for environments lacking Python support:
```bash
pyinstaller --onefile extract.py
```
The executable processes PDFs from the 'invoices' directory and automatically generates the Excel and CSV files.

## Usage

### Script Execution
To run the script from the command line:
```bash
python extract.py
# This will process all PDFs within the 'invoices' directory.
```

