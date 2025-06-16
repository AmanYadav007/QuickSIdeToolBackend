from flask import Flask, request, send_file, jsonify
from PyPDF2 import PdfReader, PdfWriter, errors # Import errors for specific exception handling
from flask_cors import CORS
import io
import logging
import fitz # Import PyMuPDF for compression

# Initialize Flask app
app = Flask(__name__) # Changed from __no_name__ to __name__ for standard practice
CORS(app)  # Enable CORS for cross-origin requests

# Configure logging
logging.basicConfig(level=logging.INFO) # Changed to INFO, DEBUG can be very verbose

# Root route for health check or info
@app.route('/')
def home():
    return "PDF Manipulator Backend is running!"

# Unlock PDF endpoint
@app.route('/unlock-pdf', methods=['POST'])
def unlock_pdf():
    # Check for file and password in request
    if 'file' not in request.files:
        logging.error("Unlock PDF: No file part in the request.")
        return jsonify({"error": "No file part in the request."}), 400
    if 'password' not in request.form:
        logging.error("Unlock PDF: Password not provided.")
        return jsonify({"error": "Password not provided."}), 400

    file = request.files['file']
    password = request.form.get('password')

    # Validate file presence and type
    if file.filename == '':
        logging.error("Unlock PDF: No selected file.")
        return jsonify({"error": "No selected file."}), 400
    if not file.filename.lower().endswith('.pdf'):
        logging.error(f"Unlock PDF: Invalid file type uploaded: {file.filename}")
        return jsonify({"error": "Invalid file type. Only PDF files are accepted."}), 400

    try:
        reader = PdfReader(file.stream)

        if not reader.is_encrypted:
            logging.info(f"Unlock PDF: File '{file.filename}' is not encrypted.")
            return jsonify({"error": "This PDF is not encrypted."}), 400

        try:
            # PyPDF2's decrypt method returns 0 for incorrect password, 1 for correct password
            # or raises an exception for other issues like corrupted file.
            if reader.decrypt(password) != 1:
                logging.warning(f"Unlock PDF: Incorrect password for '{file.filename}'.")
                return jsonify({"error": "Incorrect password for this PDF."}), 400
        except errors.FileTruncatedError as e:
            logging.error(f"Unlock PDF: Corrupted PDF file '{file.filename}': {e}")
            return jsonify({"error": "Failed to unlock PDF: Corrupted file or invalid structure."}), 400
        except Exception as e: # Catch other potential decryption issues
            logging.error(f"Unlock PDF: Unexpected error during decryption for '{file.filename}': {e}")
            return jsonify({"error": f"Failed to unlock PDF: An unexpected error occurred during decryption: {str(e)}"}), 500

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        output = io.BytesIO()
        writer.write(output)
        output.seek(0)

        logging.info(f"Unlock PDF: Successfully unlocked and sent '{file.filename}'.")
        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"unlocked_{file.filename}"
        )
    except errors.PdfReadError as e:
        logging.error(f"Error reading PDF file '{file.filename}': {e}")
        return jsonify({"error": f"Failed to read PDF: {str(e)}. It might be corrupted or malformed."}), 400
    except Exception as e:
        logging.error(f"Error unlocking PDF '{file.filename}': {e}", exc_info=True) # exc_info to log traceback
        return jsonify({"error": f"Failed to unlock PDF: An unexpected server error occurred: {str(e)}"}), 500

# Lock PDF endpoint
@app.route('/lock-pdf', methods=['POST'])
def lock_pdf():
    # Check for file and password in request
    if 'file' not in request.files:
        logging.error("Lock PDF: No file part in the request.")
        return jsonify({"error": "No file part in the request."}), 400
    if 'password' not in request.form:
        logging.error("Lock PDF: Password not provided.")
        return jsonify({"error": "Password not provided."}), 400

    file = request.files['file']
    password = request.form.get('password')

    # Validate file presence and type
    if file.filename == '':
        logging.error("Lock PDF: No selected file.")
        return jsonify({"error": "No selected file."}), 400
    if not file.filename.lower().endswith('.pdf'):
        logging.error(f"Lock PDF: Invalid file type uploaded: {file.filename}")
        return jsonify({"error": "Invalid file type. Only PDF files are accepted."}), 400

    try:
        reader = PdfReader(file.stream)

        # Optional: Check if PDF is already encrypted with the same password
        # This part is a bit tricky with PyPDF2 as it doesn't easily tell you
        # *which* password it's encrypted with without attempting to decrypt.
        # For simplicity, if it's already encrypted, we'll just re-encrypt it,
        # which effectively changes the password if different, or keeps it the same.
        # If you truly want to prevent re-locking with the same password,
        # you'd need to try decrypting with the given password first,
        # which might not be desirable for a "lock" operation.
        # So, the current approach is generally fine: just encrypt it.

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        # Encrypt the PDF with the provided password
        writer.encrypt(password)

        output = io.BytesIO()
        writer.write(output)
        output.seek(0)

        logging.info(f"Lock PDF: Successfully locked and sent '{file.filename}'.")
        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"locked_{file.filename}"
        )
    except errors.PdfReadError as e:
        logging.error(f"Error reading PDF file '{file.filename}' for locking: {e}")
        return jsonify({"error": f"Failed to read PDF: {str(e)}. It might be corrupted or malformed."}), 400
    except Exception as e:
        logging.error(f"Error locking PDF '{file.filename}': {e}", exc_info=True)
        return jsonify({"error": f"Failed to lock PDF: An unexpected server error occurred: {str(e)}"}), 500


# PDF LINK REMOVER ENDPOINT
@app.route('/remove-pdf-links', methods=['POST'])
def remove_pdf_links():
    if 'file' not in request.files:
        logging.error("Remove Links: No file part in the request.")
        return jsonify({"error": "No file part in the request."}), 400

    file = request.files['file']

    if file.filename == '':
        logging.error("Remove Links: No selected file.")
        return jsonify({"error": "No selected file."}), 400
    if not file.filename.lower().endswith('.pdf'):
        logging.error(f"Remove Links: Invalid file type uploaded: {file.filename}")
        return jsonify({"error": "Invalid file type. Only PDF files are accepted."}), 400

    try:
        reader = PdfReader(file.stream)

        if reader.is_encrypted:
            logging.warning(f"Remove Links: Attempt to remove links from encrypted PDF '{file.filename}'.")
            return jsonify({"error": "Failed to remove links: PDF is encrypted. Unlock it first."}), 400

        writer = PdfWriter()

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]

            if '/Annots' in page:
                new_annots = []
                for annot in page['/Annots']:
                    annot_obj = annot.get_object()
                    if annot_obj.get('/Subtype') != '/Link':
                        new_annots.append(annot)

                if new_annots:
                    page['/Annots'] = new_annots
                else:
                    del page['/Annots']

            writer.add_page(page)

        output_pdf = io.BytesIO()
        writer.write(output_pdf)
        output_pdf.seek(0)

        logging.info(f"Remove Links: Successfully removed links from and sent '{file.filename}'.")
        return send_file(
            output_pdf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"links_removed_{file.filename}"
        )

    except errors.PdfReadError as e:
        logging.error(f"Error reading PDF file '{file.filename}' for link removal: {e}")
        return jsonify({"error": f"Failed to read PDF for link removal: {str(e)}. It might be corrupted or malformed."}), 400
    except Exception as e:
        logging.error(f"Error processing PDF for link removal '{file.filename}': {e}", exc_info=True)
        return jsonify({"error": f"Failed to remove links from PDF: An unexpected server error occurred: {str(e)}. It might be corrupted or complex."}), 500


# PDF COMPRESSION ENDPOINT (using PyMuPDF)
@app.route('/compress_pdf', methods=['POST'])
def compress_pdf():
    if 'pdf_file' not in request.files:
        logging.error("Compress PDF: No PDF file provided.")
        return jsonify({"error": "No PDF file provided"}), 400

    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        logging.error("Compress PDF: No selected file.")
        return jsonify({"error": "No selected file"}), 400
    if not pdf_file.filename.lower().endswith('.pdf'):
        logging.error(f"Compress PDF: Invalid file type uploaded: {pdf_file.filename}")
        return jsonify({"error": "Only PDF files are accepted"}), 400

    try:
        input_pdf_bytes = pdf_file.read()
        doc = fitz.open(stream=input_pdf_bytes, filetype="pdf")

        output_buffer = io.BytesIO()
        doc.save(
            output_buffer,
            garbage=4,          # Remove unused objects
            deflate=True,       # Apply Flate compression to streams (text, vector graphics)
            clean=True,         # Perform additional cleanup
            pretty=False,       # Don't pretty-print objects (saves minor space)
            # You can uncomment and adjust these for image-specific compression:
            # img=fitz.PDF_NAME_DCTDECODE, # Force images to JPEG. This can significantly reduce size
            # img_quality=75,              # JPEG quality for recompressed images (0-100). Adjust as needed.
        )
        doc.close()
        output_buffer.seek(0)

        logging.info(f"Compress PDF: Successfully compressed and sent '{pdf_file.filename}'.")
        return send_file(
            output_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"compressed_{pdf_file.filename}" # Dynamic download name
        )

    except fitz.FileDataError as e:
        logging.error(f"Compress PDF: PyMuPDF error reading file '{pdf_file.filename}': {e}")
        return jsonify({"error": f"Failed to compress PDF: Corrupted file or invalid PDF structure."}), 400
    except Exception as e:
        logging.error(f"Error compressing PDF '{pdf_file.filename}': {e}", exc_info=True)
        return jsonify({"error": f"Failed to compress PDF: An unexpected server error occurred: {str(e)}"}), 500


# Main entry point
if __name__ == '__main__':
    # Run Flask app on port 4000
    # For production, consider using a production-ready WSGI server like Gunicorn or uWSGI
    # e.g., gunicorn -w 4 -b 0.0.0.0:4000 app:app
    app.run(debug=True, host='0.0.0.0', port=4000)