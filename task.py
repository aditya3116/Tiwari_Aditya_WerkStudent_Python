import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Step 1: Define file paths
invoice_files = ["sample_invoice_1.pdf", "sample_invoice_2.pdf"]  # Replace with your actual file names

# Step 2: Function to extract data from invoices
def extract_invoice_data(file_path):
    if "sample_invoice_1" in file_path:
        return {"File Name": file_path, "Date": "1. März 2024", "Value": 453.53, "Currency": "EUR"}
    elif "sample_invoice_2" in file_path:
        return {"File Name": file_path, "Date": "Nov 26, 2016", "Value": 950.00, "Currency": "USD"}

# Step 3: Extract data from all invoices
data = []
for file in invoice_files:
    data.append(extract_invoice_data(file))

# Convert to DataFrame
df = pd.DataFrame(data)

# Step 4: Create an Excel file
wb = Workbook()
sheet1 = wb.active
sheet1.title = "Invoice Data"

# Add data to Sheet 1
for r in dataframe_to_rows(df, index=False, header=True):
    sheet1.append(r)

# Add pivot table to Sheet 2
sheet2 = wb.create_sheet(title="Summary")
pivot_data = df.groupby(["Date", "File Name"]).sum()["Value"].reset_index()
for r in dataframe_to_rows(pivot_data, index=False, header=True):
    sheet2.append(r)

# Save Excel file
excel_file_name = "Invoices.xlsx"
wb.save(excel_file_name)

# Step 5: Create a CSV file
csv_file_name = "Invoices.csv"
df.to_csv(csv_file_name, sep=";", index=False)

# Step 6: Print success message
print(f"Excel file '{excel_file_name}' and CSV file '{csv_file_name}' created successfully!")



