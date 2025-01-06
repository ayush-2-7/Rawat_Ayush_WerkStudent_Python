import os 
import re
import fitz
import pandas as pd
from os import listdir
from os.path import isfile, join
from openpyxl import Workbook

def get_pdf_files(path):
    # Get all PDF files in the specified directory
    return [f for f in os.listdir(path) if isfile(join(path, f)) and f.lower().endswith('.pdf')]


def extract_values_from_sample1(pdf_path):
    # Open the PDF
    pdf_document = fitz.open(pdf_path)
    extracted_text = ""

    # Iterate through all pages
    for page in pdf_document:
        blocks = page.get_text("blocks")
        extracted_text += "\n".join(block[4] for block in blocks)
    
    pdf_document.close()

    # Extract gross amount and date
    gross_amount_match = re.search(r"Gross Amount incl\. VAT\s+([\d.,]+\s*€)", extracted_text)
    date_match = re.search(r"Date\s+([\d.]+\s+\w+\s+\d{4})", extracted_text)
    return gross_amount_match.group(1), date_match.group(1)

def extract_values_from_sample2(pdf_path):
    # Open the PDF
    pdf_document = fitz.open(pdf_path)
    extracted_text = ""

    # Iterate through all pages
    for page in pdf_document:
        blocks = page.get_text("blocks")
        extracted_text += "\n".join(block[4] for block in blocks)
    
    pdf_document.close()

    # Extract Total and Invoice date
    total_match = re.search(r"Total\s+(USD\s*\$[\d,]+\.\d{2})", extracted_text)
    date_match2 = re.search(r"Invoice date:\s([A-Za-z]{3}\s\d{1,2},\s\d{4})", extracted_text)
    return total_match.group(1), date_match2.group(1)

def create_excel_file(data, output_file):
    # Create a DataFrame from the extracted data
    df = pd.DataFrame(data, columns=['File Name', 'Date', 'Value'])

    # Convert 'Value' column to numeric (if possible)
    df['Numeric Value'] = df['Value'].str.extract(r'([\d.,]+)').replace({',': ''}, regex=True).astype(float)

    # Create a pivot table
    pivot_table = df.pivot_table(index='Date', columns='File Name', values='Numeric Value', aggfunc='sum', fill_value=0)

    # Write to Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Extracted Data', index=False)
        pivot_table.to_excel(writer, sheet_name='Pivot Table')

def create_csv_file(data, output_file):
    # Create a DataFrame from the extracted data
    df = pd.DataFrame(data, columns=['File Name', 'Date', 'Value'])

    # Write to CSV file with semicolon as the separator
    df.to_csv(output_file, sep=';', index=False)
    

def main():
    # Get current directory path
    path = os.getcwd()
    
    # Get list of PDF files
    files = get_pdf_files(path)

    # Data storage
    extracted_data = []

    # Process sample 1
    sample1_name = "sample_invoice_1.pdf"
    if sample1_name in files:
        pdf_path = os.path.join(path, sample1_name)
        gross_amount, date = extract_values_from_sample1(pdf_path)
        if gross_amount and date:
            extracted_data.append([sample1_name, date, gross_amount])

    # Process sample 2
    sample2_name = "sample_invoice_2.pdf"
    if sample2_name in files:
        pdf_path = os.path.join(path, sample2_name)
        total_amount, date2 = extract_values_from_sample2(pdf_path)
        if total_amount and date2:
            extracted_data.append([sample2_name, date2, total_amount])

    # Create Excel file
    if extracted_data:
        output_file_excel = os.path.join(path, 'Task.xlsx')
        create_excel_file(extracted_data, output_file_excel)
        print(f"Excel file created: {output_file_excel}")

        # Create CSV file
        output_file_csv = os.path.join(path, 'Task.csv')
        create_csv_file(extracted_data, output_file_csv)

if __name__ == "__main__":
    main()