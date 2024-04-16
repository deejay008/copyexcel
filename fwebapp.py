from flask import Flask, render_template, request, redirect, url_for
from werkzeug.utils import secure_filename
import pandas as pd
from openpyxl import Workbook
import os

app = Flask(__name__)

# Set the upload folder
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Function to check if the file extension is allowed
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an empty file without a filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            return redirect(url_for('copy_to_new_excel', filename=filename))

@app.route('/copy/<filename>')
def copy_to_new_excel(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        base_filename = "copied_data.xlsx"
        i = 1
        new_file_path = os.path.join(app.config['UPLOAD_FOLDER'], base_filename)
        while os.path.exists(new_file_path):
            new_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"copied_data_{i}.xlsx")
            i += 1
        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        return f"Data successfully copied to: {new_file_path}"
    else:
        return "File not found."

if __name__ == '__main__':
    app.run(debug=True)