from flask import Flask, render_template, request, send_from_directory
import pandas as pd
import os
import re
from openpyxl import load_workbook
from additional_service import additional_service

# Initialize the Flask application
app = Flask(__name__)

app.register_blueprint(additional_service)

# Configure the upload and processed directories
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['PROCESSED_FOLDER'] = 'processed'

# Ensure the folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)

# Route to serve the index.html page (file upload form)
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle file upload and processing
@app.route('/upload', methods=['POST'])
def upload_file():
    # Check if a file was uploaded
    if 'file' not in request.files:
        return "No file uploaded!", 400

    file = request.files['file']
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    
    # Save the uploaded file
    file.save(file_path)

    # Process the file
    output_file = process_excel(file_path)

    # Return a download link for the processed file
    return f'''
    <h1>File Processed Successfully!</h1>
    <a href="/download/{output_file}">Click here to download the processed file</a>
    '''

# Function to process the uploaded Excel file
def process_excel(file_path):
    # Read the uploaded Excel file
    df = pd.read_excel(file_path)

    # Assuming 'Faculty Name' and 'Designation' are columns in the input file
    if 'Faculty Name' not in df.columns or 'Designation' not in df.columns:
        return "Missing required columns: 'Faculty Name' or 'Designation'.", 400

    # Function to parse subject preference for name, count, and semester
    def parse_subject(pref):
        match = re.search(r'^(.*?)(?:\((\d+)\))?$', str(pref))
        if match:
            subject = match.group(1).strip()
            count = int(match.group(2)) if match.group(2) else 0
            semester = int(subject[4]) if len(subject) > 4 and subject[4].isdigit() else None
            return subject, count, semester
        return pref, 0, None

    # Process each preference column
    for col in df.columns:
        if "Preference" in col:
            # Create new columns for subject, count, and semester
            df[f"{col} Name"], df[f"{col} Count"], df[f"{col} Semester"] = zip(*df[col].map(parse_subject))

    # Assign subjects based on preferences and semester
    def assign_subjects(row):
        assigned_subject_1 = "No Assignment Available"
        assigned_subject_2 = "No Assignment Available"

        # Iterate through preferences for Subject 1 (4th semester)
        for pref_num in range(1, 5):  # Assuming up to 4 preferences
            subject_col = f"Subject Preference {pref_num} Name"
            count_col = f"Subject Preference {pref_num} Count"
            semester_col = f"Subject Preference {pref_num} Semester"

            if subject_col in row and count_col in row and semester_col in row:
                if row[semester_col] == 4 and row[count_col] <= 1:
                    assigned_subject_1 = row[subject_col]
                    break

        # Iterate through preferences for Subject 2 (6th semester)
        for pref_num in range(1, 5):
            subject_col = f"Subject Preference {pref_num} Name"
            count_col = f"Subject Preference {pref_num} Count"
            semester_col = f"Subject Preference {pref_num} Semester"

            if subject_col in row and count_col in row and semester_col in row:
                if row[semester_col] == 6 and row[count_col] <= 1:
                    assigned_subject_2 = row[subject_col]
                    break

        return assigned_subject_1, assigned_subject_2

    df['4th Semester'], df['6th Semester'] = zip(*df.apply(assign_subjects, axis=1))

    # Retain only the necessary columns
    df = df[['Faculty Name', 'Designation', '4th Semester', '6th Semester']]

    # Save the processed file
    output_file = 'processed_output.xlsx'
    output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_file)

    # Save DataFrame to Excel using openpyxl to allow column width adjustment
    df.to_excel(output_path, index=False)

    # Open the saved file to adjust the column widths
    wb = load_workbook(output_path)
    ws = wb.active

    # Adjust column widths based on the maximum length of the content in each column
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name (e.g., "A", "B", "C", ...)
        for cell in col:
            try:
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = max_length + 2  # Add some padding to the column width
        ws.column_dimensions[column].width = adjusted_width

    # Save the adjusted file
    wb.save(output_path)

    return os.path.basename(output_path)

# Route to serve the processed file for download
@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['PROCESSED_FOLDER'], filename)

# Run the Flask app
if __name__ == '__main__':
    app.run(debug=True)
