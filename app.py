from flask import Flask, request, send_file, jsonify
import pikepdf # Use pikepdf for PDF operations
from flask_cors import CORS
import io
import logging

# Initialize Flask app
app = Flask(__name__)
CORS(app)  # Enable CORS for cross-origin requests

# Configure logging
logging.basicConfig(level=logging.INFO) # Set to INFO for production, DEBUG for development

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
        pdf = None
        file.stream.seek(0) # Ensure stream is at the beginning

        # Attempt to open the PDF. pikepdf.Pdf.open handles decryption directly.
        # It will raise an error if the password is incorrect or PDF is malformed.
        try:
            pdf = pikepdf.Pdf.open(file.stream, password=password)
            if not pdf.is_encrypted:
                # If it opens without error *and* reports not encrypted, it means it was never locked
                logging.info(f"Unlock PDF: File '{file.filename}' is not encrypted. Cannot unlock.")
                return jsonify({"error": "This PDF is not encrypted. Cannot unlock."}), 400

        except pikepdf.PasswordError:
            logging.warning(f"Unlock PDF: Incorrect password for '{file.filename}'.")
            return jsonify({"error": "Incorrect password for this PDF."}), 400
        except pikepdf.PdfError as e: # Catches various PDF-related errors from pikepdf
            logging.error(f"Unlock PDF: pikepdf error during open/decrypt for '{file.filename}': {e}")
            return jsonify({"error": f"Failed to unlock PDF: Invalid PDF file or corrupted encryption: {str(e)}"}), 400
        except Exception as e:
            # Catch any other unexpected exceptions during the open attempt
            logging.error(f"Unlock PDF: Unexpected error during pikepdf open/decrypt for '{file.filename}': {e}", exc_info=True)
            return jsonify({"error": f"Failed to unlock PDF: An unexpected error occurred: {str(e)}"}), 500

        # If we reach here, the PDF was successfully opened and implicitly decrypted by pikepdf.open
        output = io.BytesIO()
        pdf.save(output) # Saves the decrypted PDF without encryption
        output.seek(0)

        logging.info(f"Unlock PDF: Successfully unlocked and sent '{file.filename}'.")
        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"unlocked_{file.filename}"
        )

    except Exception as e:
        # General catch-all for any errors not caught by more specific pikepdf errors
        logging.error(f"General error in unlock_pdf for '{file.filename}': {e}", exc_info=True)
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
        file.stream.seek(0) # Ensure stream is at the beginning
        pdf = pikepdf.Pdf.open(file.stream)

        output = io.BytesIO()
        
        # Define modern AES-256 encryption settings
        encryption = pikepdf.Encryption(
            user=password,
            owner=password, # Owner password often same as user for simplicity in tools
            R=6 # Revision 6 for AES-256 encryption
        )
        
        pdf.save(output, encryption=encryption)
        output.seek(0)

        logging.info(f"Lock PDF: Successfully locked and sent '{file.filename}'.")
        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"locked_{file.filename}"
        )
    except pikepdf.PdfError as e:
        logging.error(f"Lock PDF: pikepdf error during lock for '{file.filename}': {e}")
        return jsonify({"error": f"Failed to lock PDF: Invalid PDF structure or internal error: {str(e)}"}), 400
    except Exception as e:
        logging.error(f"Error locking PDF '{file.filename}': {e}", exc_info=True)
        return jsonify({"error": f"Failed to lock PDF: An unexpected server error occurred: {str(e)}"}), 500


# PDF LINK REMOVER ENDPOINT (updated for pikepdf)
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
        file.stream.seek(0) # Ensure stream is at the beginning
        pdf = pikepdf.Pdf.open(file.stream)

        if pdf.is_encrypted:
            logging.warning(f"Remove Links: Attempt to remove links from encrypted PDF '{file.filename}'.")
            return jsonify({"error": "Failed to remove links: PDF is encrypted. Unlock it first."}), 400

        for page in pdf.pages:
            # Check if '/Annots' exists and is an Array
            if '/Annots' in page and isinstance(page.Annots, pikepdf.Array):
                new_annots = pikepdf.Array()
                for annot in page.Annots:
                    # Keep annotations that are NOT links
                    # A link typically has /Subtype /Link or an Action (A) of type /URI or /GoTo
                    is_link = False
                    if annot.get('/Subtype') == '/Link':
                        is_link = True
                    elif annot.get('/A') and annot.A.get('/S') in ('/URI', '/GoTo'):
                        is_link = True
                    
                    if not is_link:
                        new_annots.append(annot)
                
                # Replace the /Annots array or delete it if empty
                if len(new_annots) > 0:
                    page.Annots = new_annots
                else:
                    del page.Annots # Remove the key if no annotations remain

        output_pdf = io.BytesIO()
        pdf.save(output_pdf) # Save the modified PDF
        output_pdf.seek(0)

        logging.info(f"Remove Links: Successfully removed links from and sent '{file.filename}'.")
        return send_file(
            output_pdf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"links_removed_{file.filename}"
        )

    except pikepdf.PdfError as e:
        logging.error(f"Error reading PDF file '{file.filename}' for link removal: {e}")
        return jsonify({"error": f"Failed to read PDF for link removal: {str(e)}. It might be corrupted or malformed."}), 400
    except Exception as e:
        logging.error(f"Error processing PDF for link removal '{file.filename}': {e}", exc_info=True)
        return jsonify({"error": f"Failed to remove links from PDF: An unexpected server error occurred: {str(e)}. It might be corrupted or complex."}), 500


# Main entry point
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=4000)