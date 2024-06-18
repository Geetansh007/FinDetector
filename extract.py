import fitz  # PyMuPDF
import pandas as pd
import re
import camelot
import json

def find_heading_pages(pdf_path, heading):
    # Open the PDF file
    document = fitz.open(pdf_path)
    start_page = None
    end_page = None
    heading_pattern = re.compile(r"\[\d+\]")

    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text = page.get_text("text")

        if heading in text and start_page is None:
            start_page = page_num
        elif heading_pattern.search(text) and start_page is not None:
            end_page = page_num
            break

    if start_page is None:
        raise ValueError(f"Heading '{heading}' not found in the document.")

    return start_page, end_page if end_page else len(document)

def extract_tables_within_page_range(pdf_path, start_page, end_page):
    # Use Camelot to extract tables from the specified page range
    tables = camelot.read_pdf(pdf_path, pages=f'{start_page}-{end_page+1}', flavor='stream')

    # Combine all extracted tables into a single DataFrame
    combined_df = pd.concat([table.df for table in tables], ignore_index=True)

    # Convert the DataFrame into a dictionary with the first column as the key
    table_dict = {}
    header = combined_df.iloc[0, 1:].tolist()  # Exclude the first column as it will be used as the key
    for _, row in combined_df.iloc[1:].iterrows():
        key = row[0]
        values = row[1:].tolist()
        table_dict[key] = values

    return table_dict, header

def save_to_csv(table_dict, header, csv_path):
    # Convert the dictionary back to a DataFrame for CSV saving
    data = {'Key': list(table_dict.keys())}
    for i, col_name in enumerate(header):
        data[col_name] = [values[i] for values in table_dict.values()]

    df = pd.DataFrame(data)
    df.to_csv(csv_path, index=False)

# Specify the path to your PDF file and the heading
pdf_path = "/Users/geetanshjoshi/Desktop/Findata/pdf/XBRL financial statements duly authenticated as per section 134 (including Board.pdf"
heading = "[110000] Balance sheet"

# Find the pages containing the table
start_page, end_page = find_heading_pages(pdf_path, heading)

# Extract the tables within the page range
table_dict, header = extract_tables_within_page_range(pdf_path, start_page, end_page)

# Specify the CSV file path
csv_path = "Balance_sheet.csv"

# Save the DataFrame to CSV
save_to_csv(table_dict, header, csv_path)

# Print JSON representation of the extracted table for verification
print(json.dumps(table_dict, indent=4))

print(f"CSV file saved to: {csv_path}")
