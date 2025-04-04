# PDF Table Extractor

This project provides a Python-based solution to detect and extract tables from system-generated PDF documents without relying on external libraries like Tabula or Camelot, and without converting the PDF to images. The extracted tables are saved as separate worksheets in an Excel (.xlsx) file, preserving their original structure.

## Overview

The goal of this hackathon was to create a tool that can intelligently identify and extract tabular data directly from the text content of PDFs. This approach avoids the limitations and dependencies associated with image-based table recognition or specialized PDF parsing libraries.

## Key Features

* **Pure Python Implementation:** Uses standard Python libraries (`PyPDF2`, `pandas`, `openpyxl`, `re`, `os`, `numpy`).
* **No External Dependencies (Beyond Standard Libraries):** Does not require Tabula, Camelot, or image processing tools.
* **Table Detection Algorithm:** Employs a custom heuristic algorithm to identify table boundaries based on text alignment and spacing patterns.
* **Handles Bordered and Borderless Tables:** Designed to recognize tables regardless of the presence of visible borders.
* **Structure Preservation:** Maintains the row and column structure of the extracted tables in the Excel output.
* **Multi-Page PDF Support:** Can process PDFs with tables spanning multiple pages.
* **Organized Output:** Each detected table is saved in a separate worksheet within the generated Excel file, named with the page number and table index.
* **Basic Excel Formatting:** Includes automatic adjustment of column widths for better readability.
* **Error Handling:** Includes basic error handling for file reading and table processing.

## Solution Approach

1.  **Text Extraction:** Utilizes `PyPDF2` to extract text content from each page of the PDF, along with approximate positional information derived from line breaks and spacing.
2.  **Table Boundary Detection:** A custom algorithm analyzes the extracted text line by line. It identifies potential table rows by looking for lines with multiple words or phrases separated by consistent spacing (two or more spaces). Consecutive lines with a similar number of such "parts" are grouped as a potential table.
3.  **Table Refinement:** The detected table regions are further processed using `pandas` DataFrames. Empty rows and columns are removed, and basic data cleaning is performed.
4.  **Excel Export:** The extracted and refined tables (as `pandas` DataFrames) are written to an Excel file using `openpyxl`. Each table is placed in a new worksheet, with the first row of the detected table used as the header.

## How to Use

1.  **Prerequisites:**
    * Python 3.6 or higher is required.
    * Install the necessary Python libraries:
        ```bash
        pip install PyPDF2 pandas openpyxl numpy
        ```
2.  **Download the Script:**
    * Download the `pdf_table_extractor.py` file from this repository.
3.  **Place PDF Files:**
    * Place the PDF files you want to process in the same directory as the `pdf_table_extractor.py` script, or in the input folder specified in the script.
4.  **Run the Script:**
    * Open a command prompt or terminal.
    * Navigate to the directory where you saved the script.
    * Run the script using the command:
        ```bash
        python pdf_table_extractor.py
        ```
5.  **Find the Output:**
    * A new folder named `output_tables` will be created in the same directory as the script (or the input folder).
    * For each processed PDF, an Excel file named `[PDF_filename]_tables.xlsx` will be generated in the `output_tables` folder. This Excel file will contain one or more worksheets, each representing a detected table from the corresponding PDF.


## Code Structure

* `PDFTableExtractor` Class:
    * `__init__(self, pdf_path)`: Initializes the extractor with the path to the PDF file.
    * `extract_text_with_positions(self)`: Extracts text content from the PDF with basic positional information.
    * `detect_table_boundaries(self, pages_data)`: Identifies potential table regions based on line structure.
    * `refine_tables(self, raw_tables)`: Cleans and validates the detected tables using pandas.
    * `process_pdf(self)`: Orchestrates the entire PDF processing pipeline.
    * `to_excel(self, output_path)`: Exports the extracted tables to an Excel file.
* `process_pdf_files(input_folder, output_folder)`: A helper function to process all PDF files in a given input folder.
* `if __name__ == "__main__":`: The main execution block that sets the input and output folders and processes the PDF files.

* 
