from flask import Flask, request, send_file, jsonify
from PyPDF2 import PdfReader, PdfWriter, errors
from flask_cors import CORS
import io
import logging

# Initialize Flask app
app = Flask(__name__)
CORS(app)  # Enable CORS for cross-origin requests

# Configure logging
logging.basicConfig(level=logging.INFO) # Keep INFO for production, DEBUG for dev

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
        # Using file.stream as it's generally recommended for uploaded files
        reader = PdfReader(file.stream)

        if not reader.is_encrypted:
            logging.info(f"Unlock PDF: File '{file.filename}' is not encrypted.")
            return jsonify({"error": "This PDF is not encrypted."}), 400
        
        # Explicitly check the decrypt result
        # PyPDF2's decrypt method returns 0 for incorrect password, 1 for correct password
        decrypted = reader.decrypt(password)
        
        if decrypted == 0: # Incorrect password
            logging.warning(f"Unlock PDF: Incorrect password for '{file.filename}'.")
            return jsonify({"error": "Incorrect password for this PDF."}), 400
        elif decrypted == 1: # Successfully decrypted
            logging.info(f"Unlock PDF: Successfully decrypted '{file.filename}'. Proceeding to write.")
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
        else:
            # Catch any other unexpected return value (like the '2' you saw)
            logging.error(f"Unlock PDF: PyPDF2.decrypt returned unexpected value: {decrypted} for '{file.filename}'.")
            return jsonify({"error": "Failed to unlock PDF due to an unexpected decryption issue. The PDF might be corrupted or encrypted in an unsupported way."}), 500

    except errors.PdfReadError as e:
        logging.error(f"Error reading PDF file '{file.filename}' for unlock: {e}")
        return jsonify({"error": f"Failed to read PDF: {str(e)}. It might be corrupted or malformed."}), 400
    except Exception as e:
        logging.error(f"General error unlocking PDF '{file.filename}': {e}", exc_info=True)
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
        # Using file.stream
        reader = PdfReader(file.stream)
        
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

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


# PDF LINK REMOVER ENDPOINT (unchanged, but included for completeness)
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


# Main entry point
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=4000)