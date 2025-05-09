from flask import Flask, render_template, request, send_file, redirect, url_for, flash, jsonify
import os
import uuid
import tempfile
from werkzeug.utils import secure_filename
import threading
import time
from pathlib import Path

# Import the enhanced OCR modules
from enhanced_ocr import (
    image_to_searchable_pdf,
    process_image_to_docx,
    pdf_to_searchable_pdf,
    pdf_to_docx
)

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.getcwd(), 'outputs')
app.config['ALLOWED_EXTENSIONS'] = {'jpg', 'jpeg', 'png', 'tiff', 'tif', 'pdf'}
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload size

# Make sure the folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Dictionary to track job status
job_status = {}

def allowed_file(filename):
    """Check if the file extension is allowed"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    """Render the main page"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing"""
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        # Generate a unique ID for this job
        job_id = str(uuid.uuid4())
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{job_id}_{filename}")
        file.save(file_path)
        
        # Get processing options from form
        output_format = request.form.get('output_format', 'pdf')
        language = request.form.get('language', 'eng')
        dpi = int(request.form.get('dpi', 300))
        deskew = 'deskew' in request.form
        clean = 'clean' in request.form
        table_detection = 'table_detection' in request.form
        table_style = request.form.get('table_style', 'grid')
        
        # Set output filename
        output_filename = f"{job_id}_{Path(filename).stem}.{output_format}"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        # Initialize job status
        job_status[job_id] = {
            'status': 'processing',
            'progress': 0,
            'message': 'Starting processing...',
            'input_file': filename,
            'output_file': output_filename,
            'output_path': output_path
        }
        
        # Start processing in a background thread
        thread = threading.Thread(
            target=process_file,
            args=(job_id, file_path, output_path, output_format, language, dpi, deskew, clean, table_detection, table_style)
        )
        thread.daemon = True
        thread.start()
        
        return jsonify({'job_id': job_id, 'message': 'Processing started'})
    
    flash('File type not allowed')
    return redirect(request.url)

def process_file(job_id, file_path, output_path, output_format, language, dpi, deskew, clean, table_detection, table_style):
    """Process the file in a background thread"""
    try:
        job_status[job_id]['message'] = 'Processing file...'
        job_status[job_id]['progress'] = 10
        
        input_type = Path(file_path).suffix.lower()
        
        # Process based on input file type and desired output
        if input_type in ['.jpg', '.jpeg', '.png', '.tiff', '.tif']:
            job_status[job_id]['message'] = 'Processing image...'
            job_status[job_id]['progress'] = 30
            
            if output_format == 'pdf':
                image_to_searchable_pdf(
                    file_path, 
                    output_path, 
                    lang=language, 
                    dpi=dpi, 
                    deskew=deskew, 
                    clean=clean,
                    table_detection=table_detection
                )
            else:  # docx
                process_image_to_docx(
                    file_path,
                    output_path,
                    lang=language,
                    dpi=dpi,
                    deskew=deskew,
                    clean=clean,
                    table_detection=table_detection,
                    table_style=table_style
                )
        
        elif input_type == '.pdf':
            job_status[job_id]['message'] = 'Processing PDF...'
            job_status[job_id]['progress'] = 30
            
            if output_format == 'pdf':
                pdf_to_searchable_pdf(
                    file_path,
                    output_path,
                    lang=language,
                    deskew=deskew,
                    clean=clean,
                    table_detection=table_detection
                )
            else:  # docx
                # First make sure the PDF is searchable
                with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                    temp_pdf_path = temp_pdf.name
                
                tables = pdf_to_searchable_pdf(
                    file_path,
                    temp_pdf_path,
                    lang=language,
                    deskew=deskew,
                    clean=clean,
                    table_detection=table_detection
                )
                
                job_status[job_id]['message'] = 'Converting to DOCX...'
                job_status[job_id]['progress'] = 70
                
                # Then convert to DOCX
                pdf_to_docx(
                    temp_pdf_path, 
                    output_path, 
                    lang=language, 
                    with_tables=tables, 
                    table_style=table_style
                )
                
                # Clean up temp file
                if os.path.exists(temp_pdf_path):
                    os.unlink(temp_pdf_path)
        
        job_status[job_id]['status'] = 'completed'
        job_status[job_id]['progress'] = 100
        job_status[job_id]['message'] = 'Processing completed!'
        
    except Exception as e:
        job_status[job_id]['status'] = 'error'
        job_status[job_id]['message'] = f'Error: {str(e)}'
        print(f"Error processing file: {e}")
    
    finally:
        # Clean up the input file
        if os.path.exists(file_path):
            os.unlink(file_path)

@app.route('/status/<job_id>')
def check_status(job_id):
    """Check the status of a job"""
    if job_id in job_status:
        return jsonify(job_status[job_id])
    return jsonify({'status': 'not_found'})

@app.route('/download/<job_id>')
def download_file(job_id):
    """Download the processed file"""
    if job_id in job_status and job_status[job_id]['status'] == 'completed':
        output_path = job_status[job_id]['output_path']
        if os.path.exists(output_path):
            return send_file(output_path, as_attachment=True, download_name=job_status[job_id]['output_file'])
    return "File not found", 404

@app.route('/cleanup')
def cleanup():
    """Clean up old files and jobs"""
    # Remove job statuses older than 1 hour
    current_time = time.time()
    for job_id in list(job_status.keys()):
        if 'created_at' in job_status[job_id] and current_time - job_status[job_id]['created_at'] > 3600:
            # Try to remove output file if it exists
            if 'output_path' in job_status[job_id]:
                if os.path.exists(job_status[job_id]['output_path']):
                    os.unlink(job_status[job_id]['output_path'])
            del job_status[job_id]
    
    # Remove old files from output directory
    for filename in os.listdir(app.config['OUTPUT_FOLDER']):
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.isfile(file_path):
            file_age = current_time - os.path.getmtime(file_path)
            if file_age > 3600:  # 1 hour
                os.unlink(file_path)
    
    return "Cleanup completed", 200

if __name__ == '__main__':
    app.run(debug=True)