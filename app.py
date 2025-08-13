from flask import Flask, request, send_file, jsonify
import pikepdf # Use pikepdf for PDF operations
from flask_cors import CORS
import io
import logging
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import fitz  # PyMuPDF for better text extraction

# Initialize Flask app
app = Flask(__name__)
CORS(app)  # Enable CORS for cross-origin requests

# Configure logging
logging.basicConfig(level=logging.INFO) # Set to INFO for production, DEBUG for development



# Root route and health endpoint
@app.route('/')
def home():
    return "PDF Manipulator Backend is running!"

@app.route('/health')
def health():
    return jsonify({"status": "ok"})

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
            # Attempt to open using the provided password. If it's wrong, PasswordError is thrown.
            pdf = pikepdf.Pdf.open(file.stream, password=password)
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
    # Always prefer fast lock for speed (AES-128). We keep strong available via env if needed.
    strength = os.getenv('DEFAULT_LOCK_STRENGTH', 'fast')  # 'fast' (AES-128) or 'strong' (AES-256)

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
        
        # Choose encryption strength
        # R mapping: 4 => AES-128, 6 => AES-256 (modern)
        revision = 4 if str(strength).lower() == 'fast' else 6
        encryption = pikepdf.Encryption(
            user=password,
            owner=password,  # Owner password same as user for simplicity
            R=revision
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


# PDF TO DOCX CONVERSION ENDPOINT
@app.route('/pdf-to-docx', methods=['POST'])
def pdf_to_docx():
    if 'file' not in request.files:
        logging.error("PDF to DOCX: No file part in the request.")
        return jsonify({"error": "No file part in the request."}), 400

    file = request.files['file']

    if file.filename == '':
        logging.error("PDF to DOCX: No selected file.")
        return jsonify({"error": "No selected file."}), 400
    if not file.filename.lower().endswith('.pdf'):
        logging.error(f"PDF to DOCX: Invalid file type uploaded: {file.filename}")
        return jsonify({"error": "Invalid file type. Only PDF files are accepted."}), 400

    try:
        file.stream.seek(0)
        
        # Open PDF with PyMuPDF for better text extraction
        pdf_document = fitz.open(stream=file.stream, filetype="pdf")
        
        # Create a new Word document
        doc = Document()
        
        # Set document margins
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)
        
        # Process each page
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # Extract text blocks with positioning information
            text_blocks = page.get_text("dict")
            
            # Add page break if not first page
            if page_num > 0:
                doc.add_page_break()
            
            # Process text blocks
            if "blocks" in text_blocks:
                for block in text_blocks["blocks"]:
                    if "lines" in block:
                        for line in block["lines"]:
                            if "spans" in line:
                                # Create a paragraph for each line
                                paragraph = doc.add_paragraph()
                                
                                for span in line["spans"]:
                                    if "text" in span and span["text"].strip():
                                        # Get font information
                                        font_size = span.get("size", 12)
                                        font_name = span.get("font", "Arial")
                                        is_bold = "bold" in font_name.lower() or font_size > 14
                                        
                                        # Add text run with formatting
                                        run = paragraph.add_run(span["text"])
                                        run.font.name = font_name
                                        run.font.size = Pt(font_size)
                                        run.bold = is_bold
                                
                                # Add spacing after paragraph
                                paragraph.space_after = Pt(6)
        
        # Close the PDF document
        pdf_document.close()
        
        # Save the Word document to a bytes buffer
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        # Generate output filename
        output_filename = file.filename.replace('.pdf', '.docx')
        if not output_filename.endswith('.docx'):
            output_filename += '.docx'
        
        logging.info(f"PDF to DOCX: Successfully converted '{file.filename}' to DOCX.")
        return send_file(
            docx_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=output_filename
        )

    except fitz.FileDataError as e:
        logging.error(f"PDF to DOCX: Invalid or corrupted PDF file '{file.filename}': {e}")
        return jsonify({"error": f"Invalid PDF file: {str(e)}"}), 400
    except Exception as e:
        logging.error(f"PDF to DOCX: Error converting '{file.filename}': {e}", exc_info=True)
        return jsonify({"error": f"Failed to convert PDF to DOCX: {str(e)}"}), 500

# Main entry point
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=4000)