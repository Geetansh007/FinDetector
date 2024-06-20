import os
from werkzeug.utils import secure_filename
from extract import PDFExtractor

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'pdf'


def save_uploaded_files(request, upload_folder='uploads'):
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
    
    for filename in os.listdir(upload_folder):
        if filename.endswith('.pdf'):
            file_path = os.path.join(upload_folder, filename)
            output_dir = os.path.join(output_base_folder, os.path.splitext(filename)[0])
            os.makedirs(output_dir, exist_ok=True)
            extractor = PDFExtractor(file_path, output_dir)
            extractor.extract_all_tables()
