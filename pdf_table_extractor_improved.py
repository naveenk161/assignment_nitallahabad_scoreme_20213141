import os
import re
import pandas as pd
import numpy as np
from PyPDF2 import PdfReader
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

class PDFTableExtractor:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.tables = []
        
    def clean_text(self, text):
        """Remove problematic characters"""
        # Replace non-ASCII and special characters
        text = re.sub(r'[^\x00-\x7F]+', ' ', text)
        # Replace other problematic characters
        text = text.replace('\x00', ' ').replace('\xff', ' ')
        return text.strip()

    def extract_text_with_positions(self):
        """Extract text with approximate positioning information"""
        try:
            reader = PdfReader(self.pdf_path)
            pages_data = []
            
            for page_num, page in enumerate(reader.pages):
                text = page.extract_text()
                if not text:
                    continue
                    
                # Clean the text first
                text = self.clean_text(text)
                
                # Split text into lines and estimate positions
                lines = text.split('\n')
                structured_lines = []
                
                for line in lines:
                    # Clean each line
                    line = self.clean_text(line)
                    # Estimate column positions by splitting on 2+ spaces
                    parts = re.split(r'\s{2,}', line)
                    if parts:
                        structured_lines.append({
                            'text': line,
                            'parts': parts,
                            'part_count': len(parts),
                            'page': page_num + 1
                        })
                
                pages_data.append(structured_lines)
            
            return pages_data
        except Exception as e:
            print(f"Error reading PDF: {str(e)}")
            return []

    def detect_table_boundaries(self, pages_data):
        """Identify potential table regions"""
        tables = []
        current_table = None
        
        for page_data in pages_data:
            for line in page_data:
                # If we find a line with multiple parts, it might be a table row
                if line['part_count'] > 1:
                    if current_table is None:
                        # Start new table with cleaned header
                        cleaned_header = [self.clean_text(part) for part in line['parts']]
                        current_table = {
                            'page': line['page'],
                            'header': cleaned_header,
                            'rows': []
                        }
                    else:
                        # Add to current table if column count matches
                        if line['part_count'] == len(current_table['header']):
                            cleaned_row = [self.clean_text(part) for part in line['parts']]
                            current_table['rows'].append(cleaned_row)
                else:
                    # End of table detected
                    if current_table and len(current_table['rows']) >= 1:
                        tables.append(current_table)
                    current_table = None
                    
            # Add the last table if it exists
            if current_table and len(current_table['rows']) >= 1:
                tables.append(current_table)
                current_table = None
                
        return tables
    
    def refine_tables(self, raw_tables):
        """Clean and validate detected tables"""
        refined_tables = []
        
        for table in raw_tables:
            try:
                # Convert to DataFrame for easier processing
                df = pd.DataFrame(table['rows'], columns=table['header'])
                
                # Basic cleaning
                df.replace('', np.nan, inplace=True)
                df.dropna(how='all', inplace=True)
                df.dropna(axis=1, how='all', inplace=True)
                
                # Additional cleaning
                df = df.applymap(lambda x: self.clean_text(str(x)) if isinstance(x, str) else x)
                
                if not df.empty and len(df.columns) > 1:
                    refined_tables.append({
                        'page': table['page'],
                        'data': df
                    })
            except Exception as e:
                print(f"Error processing table: {str(e)}")
                continue
                
        return refined_tables
    
    def process_pdf(self):
        """Main processing pipeline"""
        try:
            # Step 1: Extract text with positional information
            pages_data = self.extract_text_with_positions()
            
            if not pages_data:
                print("No text data extracted from PDF")
                return []
            
            # Step 2: Detect table boundaries
            raw_tables = self.detect_table_boundaries(pages_data)
            
            # Step 3: Refine and validate tables
            self.tables = self.refine_tables(raw_tables)
            
            return self.tables
        except Exception as e:
            print(f"Error processing PDF: {str(e)}")
            return []
    
    def clean_sheet_name(self, name):
        """Clean sheet name to be Excel-compatible"""
        # Remove invalid characters
        invalid_chars = r'[]:*?/\'
        for char in invalid_chars:
            name = name.replace(char, '')
        # Trim to 31 characters
        return name[:31]
    
    def to_excel(self, output_path):
        """Export extracted tables to Excel with improved formatting"""
        if not self.tables:
            print("No tables found to export")
            return False
            
        try:
            wb = Workbook()
            wb.remove(wb.active)  # Remove default sheet
            
            for i, table in enumerate(self.tables):
                try:
                    # Create clean sheet name
                    sheet_name = f"Page_{table['page']}_Table_{i+1}"
                    sheet_name = self.clean_sheet_name(sheet_name)
                    
                    # Ensure sheet name is not empty
                    if not sheet_name.strip():
                        sheet_name = f"Table_{i+1}"
                    
                    ws = wb.create_sheet(title=sheet_name)
                    
                    # Write data to sheet
                    for r_idx, row in enumerate(dataframe_to_rows(table['data'], index=False, header=True)):
                        # Clean each cell value
                        cleaned_row = []
                        for cell in row:
                            if isinstance(cell, str):
                                cleaned_cell = self.clean_text(cell)
                            else:
                                cleaned_cell = cell
                            cleaned_row.append(cleaned_cell)
                        ws.append(cleaned_row)
                        
                    # Add some basic formatting
                    for column in ws.columns:
                        max_length = 0
                        column_cells = [cell for cell in column]
                        for cell in column_cells:
                            try:
                                cell_value = str(cell.value) if cell.value is not None else ""
                                if len(cell_value) > max_length:
                                    max_length = len(cell_value)
                            except:
                                pass
                        adjusted_width = (max_length + 2) * 1.2
                        ws.column_dimensions[column_cells[0].column_letter].width = adjusted_width
                except Exception as e:
                    print(f"Error exporting table {i+1}: {str(e)}")
                    continue
                    
            wb.save(output_path)
            print(f"Successfully exported tables to {output_path}")
            return True
        except Exception as e:
            print(f"Error saving Excel file: {str(e)}")
            return False

def process_pdf_files(input_folder, output_folder):
    """Process all PDF files in the input folder"""
    if not os.path.exists(input_folder):
        print(f"Input folder not found: {input_folder}")
        return
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    processed_files = 0
    
    for filename in os.listdir(input_folder):
        if filename.lower().endswith('.pdf'):
            input_path = os.path.join(input_folder, filename)
            output_filename = f"{os.path.splitext(filename)[0]}_tables.xlsx"
            output_path = os.path.join(output_folder, output_filename)
            
            print(f"\nProcessing {filename}...")
            extractor = PDFTableExtractor(input_path)
            tables = extractor.process_pdf()
            
            if tables:
                print(f"Found {len(tables)} tables in {filename}")
                if extractor.to_excel(output_path):
                    processed_files += 1
            else:
                print(f"No tables found in {filename}")
    
    print(f"\nProcessing complete. {processed_files} PDF files were processed successfully.")

if __name__ == "__main__":
    # Set your folder paths here
    input_folder = r"C:\Users\HP\Desktop\scoreme"
    output_folder = os.path.join(input_folder, "output_tables")
    
    # Process all PDF files in the input folder
    process_pdf_files(input_folder, output_folder)
    
    print("\nScript execution finished. Press Enter to exit...")
    input()