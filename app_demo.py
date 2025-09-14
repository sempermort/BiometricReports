from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'

# Ensure upload and output folders exist
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# Route for the home page
@app.route('/')
def index():
    return render_template('index.html')

# Route for handling file uploads and processing
@app.route('/process', methods=['POST'])
def process_files():
    try:
        # Get uploaded files
        employee_file = request.files.get('employee_file')
        first_last_file = request.files.get('first_last_file')
        leave_report_file = request.files.get('leave_report_file')
        
        # Check if files were uploaded
        if not all([employee_file, first_last_file, leave_report_file]):
            flash('Please upload all required files')
            return redirect(url_for('index'))
        
        # In a real implementation, we would process the files here
        # For this demo, we'll just simulate the process
        
        # Return the processed files
        return render_template('results.html')
    except Exception as e:
        flash(f'Error processing files: {str(e)}')
        return redirect(url_for('index'))

@app.route('/download/<path:filename>')
def download_file(filename):
    # For demo purposes, we'll just redirect to the home page
    flash('In a real implementation, this would download the processed file')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)