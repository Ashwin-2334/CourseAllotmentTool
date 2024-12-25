from flask import Flask, render_template, request, send_from_directory
import pandas as pd
import os

# Initialize the Flask application
app = Flask(__name__)

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

    # Example: Add a new column 'Allocated Slot' with dummy values (you can customize this part)
    # Ensure the list is long enough by repeating or truncating it
    slots = ['Slot A', 'Slot B', 'Slot C'] * (len(df) // 3) + ['Slot A', 'Slot B', 'Slot C'][:len(df) % 3]

# Add the 'Allocated Slot' column
    df['Allocated Slot'] = slots

    #df['Allocated Slot'] = ['Slot A', 'Slot B', 'Slot C'][:len(df)]

    # Save the processed file
    output_file = 'processed_' + os.path.basename(file_path)
    output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_file)
    df.to_excel(output_path, index=False)

    return output_file

# Route to serve the processed file for download
@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['PROCESSED_FOLDER'], filename)

# Run the Flask app
if __name__ == '__main__':
    app.run(debug=True)
