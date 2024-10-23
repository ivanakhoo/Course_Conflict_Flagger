import os
from flask import Flask, request, send_file, redirect, url_for, render_template
from werkzeug.utils import secure_filename
import python.flagger
import python.modify

app = Flask(__name__)

# Folder to store uploaded and processed files
UPLOAD_FOLDER = 'uploads/'
PROCESSED_FOLDER = 'processed/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if both files are present
        if 'course_schedule' not in request.files or 'course_conflict' not in request.files:
            return 'No file part', 400

        course_schedule = request.files['course_schedule']
        course_conflict = request.files['course_conflict']

        # If user does not select file, browser may submit an empty part without filename
        if course_schedule.filename == '' or course_conflict.filename == '':
            return 'No selected file', 400

        if course_schedule and course_conflict:
            # Save the uploaded files
            schedule_filename = secure_filename(course_schedule.filename)
            conflict_filename = secure_filename(course_conflict.filename)

            schedule_path = os.path.join(app.config['UPLOAD_FOLDER'], schedule_filename)
            conflict_path = os.path.join(app.config['UPLOAD_FOLDER'], conflict_filename)

            course_schedule.save(schedule_path)
            course_conflict.save(conflict_path)

            # Define the output file path for the processed file
            output_filename = f"highlighted_{schedule_filename}"
            output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
            modified_filename = f"mod_{schedule_filename}"
            modified_path = os.path.join(app.config['PROCESSED_FOLDER'], modified_filename) 
            
            # Modify the excel file
            python.modify.shift_and_delete_rows(schedule_path)

            # Call the process function to modify the file
            python.flagger.process(schedule_path, conflict_path, output_path)

            # Redirect to download the processed file
            return redirect(url_for('download_file', filename=output_filename))

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['PROCESSED_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
