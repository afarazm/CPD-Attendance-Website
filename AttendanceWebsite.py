from flask import Flask, request, redirect, url_for, send_from_directory, render_template
import pandas as pd
import os
import webview
import sys

# Determine paths for the frozen app
if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    static_folder = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
else:
    app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['LOGO_PATH'] = os.path.join(app.static_folder, 'images', 'EG.png')

os.makedirs(os.path.join(app.root_path, app.config['UPLOAD_FOLDER']), exist_ok=True)
os.makedirs(os.path.join(app.root_path, app.config['OUTPUT_FOLDER']), exist_ok=True)

@app.route('/')
def upload_files():
    return render_template('Design.html', logo_path=app.config['LOGO_PATH'])

@app.route('/upload', methods=['POST'])
def process_files():
    if 'directory_file' not in request.files or 'cpd_file' not in request.files:
        return 'No file part', 400

    directory_file = request.files['directory_file']
    cpd_file = request.files['cpd_file']

    if directory_file.filename == '' or cpd_file.filename == '':
        return 'No selected file', 400

    directory_path = os.path.join(app.config['UPLOAD_FOLDER'], directory_file.filename)
    cpd_path = os.path.join(app.config['UPLOAD_FOLDER'], cpd_file.filename)

    directory_file.save(directory_path)
    cpd_file.save(cpd_path)

    # Process files
    new_excel = pd.read_excel(directory_path)
    new_form = pd.read_excel(cpd_path)

    merged_excel = pd.merge(new_form, new_excel[['Work Email', 'Name', 'Title', 'Specialty', 'Mobile No.', 'QID', 'Medical License']], left_on='Email Address', right_on='Work Email', how='left')
    merged_excel.drop(columns=['Work Email', 'Item Type', 'Path', 'Submitted By'], inplace=True)
    columns = ['Name'] + [col for col in merged_excel if col != 'Name']
    merged_excel = merged_excel[columns]

    final_file_path = os.path.join(app.config['OUTPUT_FOLDER'], 'CPD_Attendance_Data_for_MOPH.xlsx')
    merged_excel.to_excel(final_file_path, index=False)

    return redirect(url_for('uploaded_file', filename='CPD_Attendance_Data_for_MOPH.xlsx'))

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename)

def run_flask():
    app.run(debug=True)

def run_webview():
    window = webview.create_window('AttendanceWebsite', app)
    webview.start()

if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        run_webview()
    else:
        run_flask()
