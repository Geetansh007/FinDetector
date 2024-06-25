from flask import Flask, request
from functions import save_uploaded_files, process_uploaded_pdfs,load_pdf_excel,download_folder

app = Flask(__name__)


@app.route('/uploads', methods=['POST', 'GET'])
def upload():
    try:
        if request.method == "POST":
            path = save_uploaded_files(request)
            new_path,result = process_uploaded_pdfs(path, 'output_path')
            load_pdf_excel('output_path','Excel_folder',result)
            return "Files processed successfully", 200
    except Exception as e:
        return str(e), 500


@app.route("/download",methods = ['GET'])
def download():
    try:
        if request.method == "GET":
            return download_folder()
    except Exception as e:
        return str(e),500

if __name__ == "__main__":
    app.run(debug=True, port=5000)
