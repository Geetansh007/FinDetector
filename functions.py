import os
import shutil
from werkzeug.utils import secure_filename
from extract import PDFExtractor, Save
from extract_excel import fill_values, create_excel_template, update_values
import glob
import zipfile
import pandas as pd

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'pdf'

def clear_directories(upload_folder, output_base_folder, excel_folder):
    """
    Remove specified directories to ensure a fresh start for each upload process.
    """
    for folder in [upload_folder, output_base_folder, excel_folder]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"Cleared existing directory: {folder}")

def save_uploaded_files(request, upload_folder='uploads', output_base_folder='output', excel_folder='Excel_folder'):
    try:
        if request.method != 'POST':
            raise ValueError("Invalid request method")

        os.makedirs(upload_folder, exist_ok=True)

        if 'files[]' not in request.files:
            raise ValueError("No files part in the request")

        files = request.files.getlist('files[]')

        if not files:
            raise ValueError("No files selected for uploading")

        for file in files:
            if file.filename == '':
                raise ValueError("No selected file")
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file.save(os.path.join(upload_folder, filename))
            else:
                raise ValueError("Invalid file format. Only PDF files are allowed.")

        return upload_folder

    except Exception as e:
        raise e

def process_uploaded_pdfs(upload_folder, output_base_folder):
    os.makedirs(output_base_folder, exist_ok=True)
    results = []
    for filename in os.listdir(upload_folder):
        if filename.endswith('.pdf'):
            file_path = os.path.join(upload_folder, filename)
            output_dir = os.path.join(output_base_folder, os.path.splitext(filename)[0])
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
                extractor = PDFExtractor(file_path, output_dir)
                extractor.extract_all_tables()

    for filename in os.listdir(upload_folder):
        if filename.endswith('.pdf'):
            file_path = os.path.join(upload_folder, filename)
            company_name, monetary_unit_value = extract_and_save(file_path)
            results.append((filename, company_name, monetary_unit_value))
            print("\n", results)

    shutil.rmtree(upload_folder)

    return output_base_folder, results

def load_pdf_excel(output_base_path, excel_folder, result):
    try:
        directories = [d for d in os.listdir(output_base_path) if os.path.isdir(os.path.join(output_base_path, d))]

        os.makedirs(excel_folder, exist_ok=True)

        for folder in directories:
            folder_path = os.path.join(output_base_path, folder)
            print(f"Loading files from folder: {folder_path}")

            files = os.listdir(folder_path)
            excel_path = create_excel_template(folder_path, excel_folder)
            print("\nMaking excel\n")
            for file in files:
                if file.lower().endswith('.xlsx'):
                    file_path = os.path.join(folder_path, file)
                    fill_values(file_path, excel_path)
            update_values(excel_path, result,folder_path)

    except Exception as e:
        print(f"An error occurred: {e}")

def extract_and_save(file_path):
    extractor = Save(file_path)
    company_name = extractor.extract_company_name()
    monetary_unit_value = extractor.extract_monetary_unit()
    return company_name, monetary_unit_value

def download_folder():
    folder_name = "Excel_folder"
    zip_file_name = "excel_output.zip"

    with zipfile.ZipFile(zip_file_name, 'w') as zipf:
        excel_files = glob.glob(os.path.join(folder_name, "*.xlsx"))

        for file in excel_files:
            zipf.write(file, os.path.basename(file))

    return zip_file_name

def table_display(folder_path):
    if not os.path.exists(folder_path):
        raise ValueError(f"The folder {folder_path} does not exist.")
    dataframes = {}
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            df = pd.read_excel(file_path)
            name = os.path.splitext(filename)[0]
            dataframes[name] = df
            
    return dataframes