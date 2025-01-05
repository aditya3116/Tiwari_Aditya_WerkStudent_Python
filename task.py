import fitz  
import re
import xlsxwriter
import pandas as pd
import os

doc_extraction_rules = {
    "sample_invoice_1.pdf": {
        "rules": {
            "invoice_total": r"Gross Amount incl\. VAT\s*[\r\n]+(\d{1,3}(?:[.,]\d{2})?\s*)",
            "invoice_date": r"Date\s*[\r\n]+\s*(\d+\.\s+\w+\s+\d{4})"
        }
    },
    "sample_invoice_2.pdf": {
        "rules": {
            "invoice_total": r"Total[\s\n]*USD \$([\d,]+\.\d{2})",
            "invoice_date": r"Invoice date:\s*(\w{3}\s\d{1,2},\s\d{4})"
        }
    }
}

table_columns = ['Document_Name', 'Total_Amount', 'Doc_Date']
result_excel = 'fetched_results.xlsx'
result_csv = 'fetched_results.csv'

def get_document_data(doc_extraction_rules):
    """Get data from documents using extraction rules"""
    found_data = []
    
    for doc_name, extraction_config in doc_extraction_rules.items():
        try:
            document = fitz.open(doc_name)
            content = ""
            for page_num in range(len(document)):
                content += document[page_num].get_text()
            
            result_row = [doc_name]
            for field, pattern in extraction_config["rules"].items():
                matched = re.search(pattern, content)
                result_row.append(matched.group(1) if matched else None)
                
            found_data.append(result_row)
            
        except Exception as e:
            print(f"Error with document {doc_name}: {e}")
            
    return found_data

def save_to_spreadsheet(found_data, table_columns, result_excel):
    doc = xlsxwriter.Workbook(result_excel)
    sheet = doc.add_worksheet("DataSheet")
    
    for idx, header in enumerate(table_columns):
        sheet.write(0, idx, header)
    
    for row_idx, row_data in enumerate(found_data, start=1):
        for col_idx, value in enumerate(row_data):
            sheet.write(row_idx, col_idx, value)
            
    df = pd.DataFrame(found_data, columns=table_columns)
    
    pivot = pd.pivot_table(df, 
                          values='Total_Amount',
                          index=['Doc_Date'],
                          columns=['Document_Name'],
                          aggfunc='sum',
                          fill_value=0)
    
    pivot_sheet = doc.add_worksheet("PivotView")
    
    for i, idx in enumerate(pivot.index):
        pivot_sheet.write(i+1, 0, idx)
        for j, col in enumerate(pivot.columns):
            if i == 0:
                pivot_sheet.write(0, j+1, col)
            pivot_sheet.write(i+1, j+1, pivot.iloc[i,j])
            
    doc.close()
    
def save_to_csv(result_excel, result_csv):
    df = pd.read_excel(result_excel)
    df.to_csv(result_csv, index=False, sep=';')

def main():
    found_data = get_document_data(doc_extraction_rules)
    
    save_to_spreadsheet(found_data, table_columns, result_excel)
    
    save_to_csv(result_excel, result_csv)

if __name__ == "__main__":
    main()
