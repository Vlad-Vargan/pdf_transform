import os
from datetime import datetime
from werkzeug.utils import secure_filename

import flask
from flask import Flask, render_template, request, send_from_directory
import pdftotext

from parser import PDFParser
from parser import ExcelConverter


app = Flask(__name__, template_folder="templates")
# app.config['MAX_CONTENT_LENGTH'] = 300 * 1024 * 1024 # will secure the request body size
app.config['UPLOAD_EXTENSIONS'] = ['.pdf']
app.config['UPLOAD_PATH'] = 'uploads'


@app.errorhandler(413)
def too_large(e):
    return "File is too large", 413

@app.route('/')
def index():
    # files = os.listdir(app.config['UPLOAD_PATH'])
    return render_template("index.html")  #, files=files)

@app.route('/', methods=['POST'])
def upload_files():
    uploaded_file = request.files['file']

    now = datetime.now()
    date_time_filename = now.strftime("%m_%d_%Y-%H%M%S-") + uploaded_file.filename
    filename = secure_filename(date_time_filename)

    if filename != '':
        file_ext = os.path.splitext(filename)[1]
        if file_ext not in app.config['UPLOAD_EXTENSIONS']:
            return "Not a PDF extension", 400
        upload_path = os.path.join(app.config['UPLOAD_PATH'], filename)
        uploaded_file.save(upload_path)

        try:
            data = PDFParser.parse(upload_path)
        except pdftotext.Error as e:
            return "Failed to open pdf file", 400
        os.remove(upload_path)
        if not data:
            return send_from_directory("xlsx", "empty.xlsx")

        xlsx_file_path = ExcelConverter.convert(data, filename)

        response = send_from_directory("uploads", xlsx_file_path)
        # print("file sent, deleting...")
        os.remove(os.path.join("uploads", xlsx_file_path))
        # print("done.")
        return response
