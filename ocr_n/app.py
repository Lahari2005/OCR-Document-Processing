import os
from flask import Flask, render_template, request, jsonify
import fitz
import pytesseract
from PIL import Image
import io
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# Tesseract path
pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\902273\AppData\Local\Tesseract-OCR\tesseract.exe"
)

# Create uploads folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(pdf_path):
    """Extract text from PDF using OCR"""
    if not os.path.exists(pdf_path):
        return None, "Error: PDF file not found."
    
    try:
        pdf = fitz.open(pdf_path)
        all_text = []
        
        for page_number, page in enumerate(pdf, start=1):
            # Render PDF page as image
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            image = Image.open(io.BytesIO(pix.tobytes("png")))
            
            # OCR
            text = pytesseract.image_to_string(image)
            all_text.append({
                'page': page_number,
                'text': text
            })
        
        pdf.close()
        return all_text, None
    except Exception as e:
        return None, str(e)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Only PDF files are allowed'}), 400
    
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Extract text
        result, error = extract_text_from_pdf(filepath)
        
        # Clean up
        os.remove(filepath)
        
        if error:
            return jsonify({'error': error}), 500
        
        return jsonify({'result': result}), 200
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='localhost', port=5000)
