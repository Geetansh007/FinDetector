import PyPDF2
import re
import pdfplumber
import pandas as pd
import os
from xlsxwriter import Workbook


class PDFExtractor:
    def __init__(self, file_path, output_dir):
        self.file_path = file_path
        self.output_dir = output_dir
        self.reader = PyPDF2.PdfReader(file_path)
        self.names = ['[110000] Balance sheet','[210000] Statement of profit and loss','[401100] Notes - Subclassification and notes on liabilities and assets',
                      '[500100] Notes - Subclassification and notes on income and expenses','[610800] Notes - Related party']

    def shorten_sheet_name(self, name):
        return ''.join(e for e in name if e.isalnum())[:31]

    def extract_all_tables(self):
        os.makedirs(self.output_dir, exist_ok=True)
        
        for name in self.names:
            try:
                table_data = self.find_headings(name)
                if table_data is not None:
                    file_name = f"{self.shorten_sheet_name(name)}.xlsx"
                    excel_path = os.path.join(self.output_dir, file_name)
                    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                        table_data.to_excel(writer, sheet_name='Sheet1', index=False)
                    print(f"Extracted data for {name} saved to {file_name}")
                else:
                    print(f"No matching heading found for {name}")
            except Exception as e:
                print(f"Error processing {name}: {e}")

    def find_headings(self, name):
        try:
            with open(self.file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                num_pages = len(reader.pages)
                
                heading_pattern = re.compile(r'\[\d{6}\]\s.*')

                start_page_number = -1
                end_page_number = 0

                for page_num in range(num_pages):
                    page_obj = reader.pages[page_num]
                    text = page_obj.extract_text()
                    
                    if text:
                        if name in text:
                            start_page_number = page_num
                        elif heading_pattern.search(text) and start_page_number != -1:
                            end_page_number = page_num
                            break
                
                if end_page_number == 0:
                    end_page_number = start_page_number + 1
                
                if start_page_number != -1: 
                    print(start_page_number,end_page_number,name)
                    return self.extract_table(start_page_number, end_page_number, name)
                else:
                    return None
        except Exception as e:
            print(f"Error in find_headings for {name}: {e}")
            return None

    def extract_table(self, start, end, name):
        try:
            all_data = pd.DataFrame()
            with pdfplumber.open(self.file_path) as pdf:
                for page_num in range(start, end + 1):
                    if page_num < len(pdf.pages):
                        page = pdf.pages[page_num]
                        tables = page.extract_tables()

                        for table in tables:
                            table_data = pd.DataFrame(table)
                            if all_data.empty:
                                all_data = table_data
                            else:
                                if table_data.shape[1] == all_data.shape[1]:
                                    all_data = pd.concat([all_data, table_data.iloc[1:]], ignore_index=True)
                                else:
                                    all_data = pd.concat([all_data, table_data], ignore_index=True)
            return all_data
        except Exception as e:
            print(f"Error in extract_table for {name}: {e}")
            return pd.DataFrame()

class Save:
    def __init__(self,file_path):
        self.file_path = file_path
        self.reader = PyPDF2.PdfReader(file_path)

    def extract_company_name(self):
        first_page = self.reader.pages[0]
        text = first_page.extract_text()
        lines = text.split('\n')
        for line in lines:
            if line.strip():
                company_name = line.strip()
                return company_name
        return None

    def extract_monetary_unit(self):
        pattern = re.compile(r"Unless otherwise specified, all monetary values are in ([\w\s]+) of INR")
        unit_mapping = {
            'Millions': 1000000,
            'Lakhs': 100000,
            'Thousands': 1000,
            'Crores': 10000000,
            'Billions': 1000000000
        }

        for page_num in range(6):
            if page_num >= len(self.reader.pages):
                break
            page = self.reader.pages[page_num]
            text = page.extract_text()
            match = pattern.search(text)
            if match:
                unit = match.group(1).strip()
                return unit_mapping.get(unit, 1)
        
        return 1








if __name__ == "__main__":
    extractor = PDFExtractor(file_path="", output_dir="")
    extractor.extract_all_tables()