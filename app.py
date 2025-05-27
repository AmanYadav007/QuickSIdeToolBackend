from flask import Flask, request, send_file, jsonify # Import jsonify
from PyPDF2 import PdfReader, PdfWriter
from flask_cors import CORS
import io
import logging

# Initialize Flask app
app = Flask(__name__)
CORS(app)  # Enable CORS for cross-origin requests

# Configure logging
logging.basicConfig(level=logging.DEBUG)

# Root route for health check or info
@app.route('/')
def home():
    return "PDF Manipulator Backend is running!"

# Unlock PDF endpoint
@app.route('/unlock-pdf', methods=['POST'])
def unlock_pdf():
    # Check for file and password in request
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request."}), 400
    if 'password' not in request.form:
        return jsonify({"error": "Password not provided."}), 400

    file = request.files['file']
    password = request.form.get('password')

    # Validate file presence and type
    if file.filename == '':
        return jsonify({"error": "No selected file."}), 400
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Invalid file type. Only PDF files are accepted."}), 400

    try:
        reader = PdfReader(file.stream) # Use file.stream for direct reading

        if not reader.is_encrypted:
            return jsonify({"error": "This PDF is not encrypted."}), 400 # More specific message
        
        # Attempt to decrypt
        try:
            reader.decrypt(password)
        except Exception: # PyPDF2 raises generic exception for incorrect password
            return jsonify({"error": "Incorrect password for this PDF."}), 400

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        output = io.BytesIO()
        writer.write(output)
        output.seek(0)

        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"unlocked_{file.filename}" # Dynamic download name
        )
    except Exception as e:
        logging.error(f"Error unlocking PDF: {e}")
        return jsonify({"error": f"Failed to unlock PDF: {str(e)}. It might be corrupted."}), 500

# Lock PDF endpoint
@app.route('/lock-pdf', methods=['POST'])
def lock_pdf():
    # Check for file and password in request
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request."}), 400
    if 'password' not in request.form:
        return jsonify({"error": "Password not provided."}), 400

    file = request.files['file']
    password = request.form.get('password')

    # Validate file presence and type
    if file.filename == '':
        return jsonify({"error": "No selected file."}), 400
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Invalid file type. Only PDF files are accepted."}), 400
        
    try:
        reader = PdfReader(file.stream)
        
        # Optional: Check if PDF is already encrypted to give specific message
        # This requires attempting to decrypt first, which might not be desired for simple locking.
        # For now, we'll just re-encrypt if already locked.

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        # Encrypt the PDF with the provided password
        writer.encrypt(password)

        output = io.BytesIO()
        writer.write(output)
        output.seek(0)

        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"locked_{file.filename}" # Dynamic download name
        )
    except Exception as e:
        logging.error(f"Error locking PDF: {e}")
        # Add more specific error messages if needed, e.g., if PDF is corrupted
        return jsonify({"error": f"Failed to lock PDF: {str(e)}"}), 500


# NEW PDF LINK REMOVER ENDPOINT
@app.route('/remove-pdf-links', methods=['POST'])
def remove_pdf_links():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request."}), 400

    file = request.files['file']
    
    if file.filename == '':
        return jsonify({"error": "No selected file."}), 400
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Invalid file type. Only PDF files are accepted."}), 400

    try:
        reader = PdfReader(file.stream)
        
        if reader.is_encrypted:
            # You might want to handle decryption here if the user provides a password
            # or simply return an error that links cannot be removed from encrypted PDFs without password.
            return jsonify({"error": "Failed to remove links: PDF is encrypted. Unlock it first."}), 400

        writer = PdfWriter()

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            
            # This is the core logic for removing link annotations
            # It iterates through existing annotations on the page
            # and only adds back annotations that are *not* of the '/Link' subtype.
            if '/Annots' in page:
                new_annots = []
                for annot in page['/Annots']:
                    annot_obj = annot.get_object() # Get the actual annotation dictionary
                    if annot_obj.get('/Subtype') != '/Link':
                        new_annots.append(annot) # Keep non-link annotations
                
                if new_annots:
                    page['/Annots'] = new_annots
                else:
                    # If no non-link annotations remain, remove the /Annots entry entirely
                    # to clean up the PDF structure for this page.
                    del page['/Annots'] 
            
            writer.add_page(page) # Add the potentially modified page to the writer

        output_pdf = io.BytesIO()
        writer.write(output_pdf)
        output_pdf.seek(0) # Go to the start of the BytesIO object

        # Send the modified PDF back
        return send_file(
            output_pdf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"links_removed_{file.filename}"
        )

    except Exception as e:
        logging.error(f"Error processing PDF for link removal: {e}")
        # Add more specific error messages if needed for other PyPDF2 errors
        return jsonify({"error": f"Failed to remove links from PDF: {str(e)}. It might be corrupted or complex."}), 500


# Main entry point
if __name__ == '__main__':
    # Run Flask app on port 4000
    app.run(debug=True, host='0.0.0.0', port=4000)