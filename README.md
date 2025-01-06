# Uniper Task

This Python script automatically extracts financial data from PDF invoices and generates summary reports in Excel and CSV formats. The program is designed to handle two specific invoice formats but can be extended for additional templates.

## Overview

The script performs these key functions:
1. Scans a directory for PDF invoice files
2. Extracts dates and amount values from recognized invoice formats
3. Creates an Excel file with raw data and a pivot table summary
4. Generates a CSV export with semicolon separators

## Supported Invoice Types

Currently processes two invoice formats:
- sample_invoice_1: Extracts "Gross Amount incl. VAT" and date
- sample_invoice_2: Extracts "Total" and invoice date

## Output Files

The script generates two summary files in the same directory:
- `Task.xlsx`: Contains raw data and a pivot table
- `Task.csv`: Contains raw data in semicolon-separated format

## Dependencies

Required Python packages:
- PyMuPDF 
- fitz
- pandas
- openpyxl


## Problem Faced

### Executable(.exe) Creation
The attempt to create an executable (.exe) file failed despite the script running successfully in Python. The package dependencies were not properly bundled in the executable, preventing it from running as intended. This remains an unresolved issue requiring further investigation into packaging configuration.
