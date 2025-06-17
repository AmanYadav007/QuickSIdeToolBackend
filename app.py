from flask import Flask, request, send_file, jsonify
# from PyPDF2 import PdfReader, PdfWriter, errors # REMOVED PyPDF2
import pikepdf # ADDED pikepdf
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
        # pikepdf.open can take a file-like object or a path
        # It raises various exceptions on errors
        try:
            # Try to open without a password first to check if it's encrypted
            pdf = pikepdf.Pdf.open(file.stream)
            # If successful, it means it was not encrypted, or it was opened without needing a password.
            # In unlock context, if it was encrypted and we opened it without password, it means incorrect password.
            # However, if it opened without error and we are in unlock mode, it means it wasn't encrypted.
            if not pdf.is_encrypted:
                logging.info(f"Unlock PDF: File '{file.filename}' is not encrypted. Cannot unlock.")
                return jsonify({"error": "This PDF is not encrypted. Cannot unlock."}), 400
            # If it reached here, it means it *was* encrypted but opened without a password,
            # indicating a potential logic error or that PyPDF2 encrypted it weakly.
            # However, pikepdf.Pdf.open should raise an error if it's password-protected.
            # Let's re-open with the password.
            file.stream.seek(0) # Reset stream position for re-opening
            pdf = pikepdf.Pdf.open(file.stream, password=password)


        except pikepdf.PasswordError:
            logging.warning(f"Unlock PDF: Incorrect password for '{file.filename}'.")
            return jsonify({"error": "Incorrect password for this PDF."}), 400
        except pikepdf.PdfError as e: # Catch other pikepdf specific errors (e.g., malformed PDF)
            logging.error(f"Unlock PDF: pikepdf error during open/decrypt for '{file.filename}': {e}")
            return jsonify({"error": f"Failed to unlock PDF: Invalid PDF file or corrupted encryption: {str(e)}"}), 400
        except Exception as e:
            logging.error(f"Unlock PDF: Unexpected error during pikepdf open/decrypt for '{file.filename}': {e}", exc_info=True)
            return jsonify({"error": f"Failed to unlock PDF: An unexpected error occurred: {str(e)}"}), 500

        # If we reach here, the PDF was successfully opened and implicitly decrypted by pikepdf.open
        # We now just need to save it without encryption.
        output = io.BytesIO()
        pdf.save(output) # Saves without encryption if opened successfully decrypted
        output.seek(0)

        logging.info(f"Unlock PDF: Successfully unlocked and sent '{file.filename}'.")
        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"unlocked_{file.filename}"
        )

    except Exception as e:
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
        # Open the PDF using pikepdf
        # If it's already encrypted and you don't provide a password,
        # pikepdf might still open it but with limited access.
        # When you save with encryption, it will apply the new encryption.
        pdf = pikepdf.Pdf.open(file.stream)

        output = io.BytesIO()
        
        # Define encryption settings (Owner password is same as user password for simplicity)
        # R=6 is for AES-256 encryption, which is modern and strong.
        encryption = pikepdf.Encryption(
            user=password,
            owner=password,
            R=6 # Revision 6 for AES-256
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


# PDF LINK REMOVER ENDPOINT (modified to use pikepdf)
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
        pdf = pikepdf.Pdf.open(file.stream)

        if pdf.is_encrypted:
            logging.warning(f"Remove Links: Attempt to remove links from encrypted PDF '{file.filename}'.")
            return jsonify({"error": "Failed to remove links: PDF is encrypted. Unlock it first."}), 400

        # Iterate through pages and remove link annotations
        # pikepdf handles this differently than PyPDF2
        for page in pdf.pages:
            # Get the /Annots array. If it doesn't exist, create an empty one.
            if '/Annots' in page:
                new_annots = pikepdf.Array()
                for annot in page.Annots:
                    # Check if the annotation is a /Link (URI or GoTo action)
                    # This check is more direct and robust with pikepdf objects
                    if annot.A and annot.A.Type == '/Action' and (annot.A.S == '/URI' or annot.A.S == '/GoTo'):
                        logging.info(f"Removed a link annotation from page.")
                    elif annot.Subtype == '/Link': # Another check for direct Link annotation
                        logging.info(f"Removed a link annotation (Subtype /Link) from page.")
                    else:
                        new_annots.append(annot)
                
                # Replace the /Annots array if there are remaining annotations
                if len(new_annots) > 0:
                    page.Annots = new_annots
                else:
                    del page.Annots # Remove /Annots key if no annotations remain

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