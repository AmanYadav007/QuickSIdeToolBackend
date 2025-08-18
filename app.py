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
from PIL import Image, ImageOps, ImageEnhance  # Add Pillow imports for image processing

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
        # PyMuPDF expects bytes for 'stream', not a file-like object
        pdf_bytes = file.read()
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
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

# IMAGE COMPRESSION ENDPOINT
@app.route('/compress-image', methods=['POST'])
def compress_image():
    """
    Advanced server-side image compression with multiple options:
    - Quality control (1-100)
    - Format conversion (JPEG, PNG, WebP)
    - Resize options
    - Metadata preservation
    - Advanced compression algorithms
    """
    if 'file' not in request.files:
        logging.error("Image compression: No file part in the request.")
        return jsonify({"error": "No file part in the request."}), 400

    file = request.files['file']
    if file.filename == '':
        logging.error("Image compression: No selected file.")
        return jsonify({"error": "No selected file."}), 400

    # Get compression parameters
    quality = int(request.form.get('quality', 85))
    output_format = request.form.get('format', 'JPEG').upper()
    resize_width = request.form.get('resize_width')
    resize_height = request.form.get('resize_height')
    preserve_metadata = request.form.get('preserve_metadata', 'false').lower() == 'true'
    optimize = request.form.get('optimize', 'true').lower() == 'true'
    
    # Validate quality range
    if not 1 <= quality <= 100:
        return jsonify({"error": "Quality must be between 1 and 100"}), 400

    try:
        file.stream.seek(0)
        
        # Open image with Pillow
        with Image.open(file.stream) as img:
            # Convert to RGB if saving as JPEG
            if output_format == 'JPEG' and img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Handle resize if specified
            if resize_width or resize_height:
                current_width, current_height = img.size
                
                if resize_width and resize_height:
                    # Both dimensions specified - resize to exact size
                    new_size = (int(resize_width), int(resize_height))
                elif resize_width:
                    # Only width specified - maintain aspect ratio
                    ratio = int(resize_width) / current_width
                    new_size = (int(resize_width), int(current_height * ratio))
                else:
                    # Only height specified - maintain aspect ratio
                    ratio = int(resize_height) / current_height
                    new_size = (int(current_width * ratio), int(resize_height))
                
                # Use high-quality resampling
                img = img.resize(new_size, Image.Resampling.LANCZOS)
            
            # Prepare output buffer
            output_buffer = io.BytesIO()
            
            # Save with appropriate format and options
            if output_format == 'JPEG':
                # JPEG specific options
                save_kwargs = {
                    'format': 'JPEG',
                    'quality': quality,
                    'optimize': optimize,
                    'progressive': True  # Progressive JPEG for better compression
                }
                
                # Preserve EXIF data if requested
                if preserve_metadata and 'exif' in img.info:
                    save_kwargs['exif'] = img.info['exif']
                
                img.save(output_buffer, **save_kwargs)
                
            elif output_format == 'PNG':
                # PNG specific options
                save_kwargs = {
                    'format': 'PNG',
                    'optimize': optimize
                }
                
                # PNG doesn't use quality parameter, but we can optimize
                if optimize:
                    # Convert to P mode if image is grayscale for better compression
                    if img.mode == 'L':
                        img = img.convert('P', palette=Image.ADAPTIVE, colors=256)
                
                img.save(output_buffer, **save_kwargs)
                
            elif output_format == 'WEBP':
                # WebP specific options
                save_kwargs = {
                    'format': 'WEBP',
                    'quality': quality,
                    'method': 6,  # Compression method (0-6, higher = better compression but slower)
                    'lossless': False
                }
                
                img.save(output_buffer, **save_kwargs)
                
            else:
                return jsonify({"error": f"Unsupported output format: {output_format}"}), 400
            
            output_buffer.seek(0)
            
            # Generate output filename
            base_name = os.path.splitext(file.filename)[0]
            extension = output_format.lower()
            if output_format == 'JPEG':
                extension = 'jpg'
            output_filename = f"{base_name}_compressed.{extension}"
            
            logging.info(f"Image compression: Successfully compressed '{file.filename}' to {output_format} with quality {quality}")
            
            return send_file(
                output_buffer,
                mimetype=f'image/{output_format.lower()}',
                as_attachment=True,
                download_name=output_filename
            )

    except Exception as e:
        logging.error(f"Image compression: Error processing '{file.filename}': {e}", exc_info=True)
        return jsonify({"error": f"Failed to compress image: {str(e)}"}), 500

# BATCH IMAGE COMPRESSION ENDPOINT
@app.route('/compress-images-batch', methods=['POST'])
def compress_images_batch():
    """
    Batch compress multiple images with the same settings
    Returns a ZIP file containing all compressed images
    """
    if 'files' not in request.files:
        return jsonify({"error": "No files provided"}), 400
    
    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({"error": "No valid files selected"}), 400
    
    # Get compression parameters
    quality = int(request.form.get('quality', 85))
    output_format = request.form.get('format', 'JPEG').upper()
    optimize = request.form.get('optimize', 'true').lower() == 'true'
    
    try:
        import zipfile
        
        # Create ZIP buffer
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for file in files:
                if file.filename == '':
                    continue
                    
                try:
                    file.stream.seek(0)
                    
                    with Image.open(file.stream) as img:
                        # Convert to RGB if saving as JPEG
                        if output_format == 'JPEG' and img.mode != 'RGB':
                            img = img.convert('RGB')
                        
                        # Prepare output buffer for this image
                        img_buffer = io.BytesIO()
                        
                        # Save with appropriate format
                        if output_format == 'JPEG':
                            img.save(img_buffer, format='JPEG', quality=quality, optimize=optimize, progressive=True)
                        elif output_format == 'PNG':
                            img.save(img_buffer, format='PNG', optimize=optimize)
                        elif output_format == 'WEBP':
                            img.save(img_buffer, format='WEBP', quality=quality, method=6, lossless=False)
                        
                        img_buffer.seek(0)
                        
                        # Generate filename for this image
                        base_name = os.path.splitext(file.filename)[0]
                        extension = output_format.lower()
                        if output_format == 'JPEG':
                            extension = 'jpg'
                        output_filename = f"{base_name}_compressed.{extension}"
                        
                        # Add to ZIP
                        zip_file.writestr(output_filename, img_buffer.getvalue())
                        
                except Exception as e:
                    logging.warning(f"Failed to compress {file.filename}: {e}")
                    continue
        
        zip_buffer.seek(0)
        
        logging.info(f"Batch compression: Successfully compressed {len(files)} images to {output_format}")
        
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='compressed_images.zip'
        )
        
    except Exception as e:
        logging.error(f"Batch compression: Error processing files: {e}", exc_info=True)
        return jsonify({"error": f"Failed to process batch compression: {str(e)}"}), 500

# FILE CONVERSION ENDPOINTS
@app.route('/convert/pdf-to-word', methods=['POST'])
def convert_pdf_to_word():
    """
    Convert PDF to Word document (.docx)
    This is an alias for the existing pdf-to-docx endpoint
    """
    return pdf_to_docx()

@app.route('/convert/pdf-to-excel', methods=['POST'])
def convert_pdf_to_excel():
    """
    Convert PDF to Excel spreadsheet (.xlsx)
    Extracts tables and text from PDF and creates Excel file
    """
    if 'file' not in request.files:
        logging.error("PDF to Excel: No file part in the request.")
        return jsonify({"error": "No file part in the request."}), 400

    file = request.files['file']
    if file.filename == '':
        logging.error("PDF to Excel: No selected file.")
        return jsonify({"error": "No selected file."}), 400
    if not file.filename.lower().endswith('.pdf'):
        logging.error(f"PDF to Excel: Invalid file type uploaded: {file.filename}")
        return jsonify({"error": "Invalid file type. Only PDF files are accepted."}), 400

    try:
        file.stream.seek(0)
        
        # Open PDF with PyMuPDF for text and table extraction
        pdf_bytes = file.read()
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Create Excel file using openpyxl
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, Border, Side
        except ImportError:
            # Fallback to CSV if openpyxl is not available
            import csv
            csv_buffer = io.BytesIO()
            csv_writer = csv.writer(csv_buffer)
            
            # Extract text from each page
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                text = page.get_text()
                if text.strip():
                    csv_writer.writerow([f"Page {page_num + 1}"])
                    for line in text.split('\n'):
                        if line.strip():
                            csv_writer.writerow([line.strip()])
                    csv_writer.writerow([])  # Empty row between pages
            
            csv_buffer.seek(0)
            pdf_document.close()
            
            # Generate output filename
            output_filename = file.filename.replace('.pdf', '.csv')
            if not output_filename.endswith('.csv'):
                output_filename += '.csv'
            
            logging.info(f"PDF to Excel: Successfully converted '{file.filename}' to CSV (fallback).")
            return send_file(
                csv_buffer,
                mimetype='text/csv',
                as_attachment=True,
                download_name=output_filename
            )
        
        # Use openpyxl for proper Excel creation
        wb = Workbook()
        ws = wb.active
        ws.title = "PDF Content"
        
        # Set up styles
        header_font = Font(bold=True, size=14)
        page_font = Font(bold=True, size=12, color="366092")
        content_font = Font(size=11)
        
        # Extract content from each page
        row = 1
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # Add page header
            ws.cell(row=row, column=1, value=f"Page {page_num + 1}").font = page_font
            row += 1
            
            # Extract text
            text = page.get_text()
            if text.strip():
                # Split text into lines and add to Excel
                lines = text.split('\n')
                for line in lines:
                    if line.strip():
                        ws.cell(row=row, column=1, value=line.strip()).font = content_font
                        row += 1
            
            # Try to extract tables
            try:
                tables = page.get_tables()
                for table_idx, table in enumerate(tables):
                    if table:
                        # Add table header
                        ws.cell(row=row, column=1, value=f"Table {table_idx + 1}").font = header_font
                        row += 1
                        
                        # Add table data
                        for table_row in table:
                            for col_idx, cell_value in enumerate(table_row):
                                if cell_value and str(cell_value).strip():
                                    ws.cell(row=row, column=col_idx + 1, value=str(cell_value).strip()).font = content_font
                            row += 1
                        row += 1  # Space after table
            except Exception as e:
                logging.warning(f"Could not extract tables from page {page_num + 1}: {e}")
            
            row += 1  # Space between pages
        
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Close the PDF document
        pdf_document.close()
        
        # Save the Excel document to a bytes buffer
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        # Generate output filename
        output_filename = file.filename.replace('.pdf', '.xlsx')
        if not output_filename.endswith('.xlsx'):
            output_filename += '.xlsx'
        
        logging.info(f"PDF to Excel: Successfully converted '{file.filename}' to Excel.")
        return send_file(
            excel_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_filename
        )

    except fitz.FileDataError as e:
        logging.error(f"PDF to Excel: Invalid or corrupted PDF file '{file.filename}': {e}")
        return jsonify({"error": f"Invalid PDF file: {str(e)}"}), 400
    except Exception as e:
        logging.error(f"PDF to Excel: Error converting '{file.filename}': {e}", exc_info=True)
        return jsonify({"error": f"Failed to convert PDF to Excel: {str(e)}"}), 500

# PDF COMPRESSION ENDPOINT
@app.route('/compress-pdf', methods=['POST'])
def compress_pdf():
    """
    Compress PDF files using PyMuPDF with multiple compression levels
    Supports different compression strategies for various use cases
    """
    if 'file' not in request.files:
        logging.error("PDF compression: No file part in the request.")
        return jsonify({"error": "No file part in the request."}), 400

    file = request.files['file']
    if file.filename == '':
        logging.error("PDF compression: No selected file.")
        return jsonify({"error": "No selected file."}), 400
    if not file.filename.lower().endswith('.pdf'):
        logging.error(f"PDF compression: Invalid file type uploaded: {file.filename}")
        return jsonify({"error": "Invalid file type. Only PDF files are accepted."}), 400

    # Get compression parameters
    compression_level = request.form.get('compression_level', 'medium')
    
    try:
        file.stream.seek(0)
        
        # Read PDF bytes for PyMuPDF
        pdf_bytes = file.read()
        
        # Open PDF with PyMuPDF
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Prepare output buffer
        output_buffer = io.BytesIO()
        
        # Apply compression based on level
        if compression_level == 'low':
            # Light compression - maintain quality, minimal size reduction
            pdf_document.save(
                output_buffer,
                garbage=1,      # Remove unused objects
                deflate=True,   # Compress streams
                clean=True,     # Clean content streams
                linear=True     # Optimize for web
            )
        elif compression_level == 'high':
            # Aggressive compression - maximum size reduction
            # Use the most aggressive settings that actually reduce size
            pdf_document.save(
                output_buffer,
                garbage=4,      # Remove all unused objects
                deflate=True,   # Compress streams
                clean=True,     # Clean content streams
                linear=True,    # Optimize for web
                pretty=False,   # Remove formatting
                ascii=False     # Use binary instead of ASCII
            )
            
        else:  # medium (default)
            # Balanced compression - good quality and size
            pdf_document.save(
                output_buffer,
                garbage=3,      # Remove most unused objects
                deflate=True,   # Compress streams
                clean=True,     # Clean content streams
                linear=True     # Optimize for web
            )
        
        # Calculate initial compression ratio
        original_size = len(pdf_bytes)
        initial_compressed_size = len(output_buffer.getvalue())
        initial_ratio = ((original_size - initial_compressed_size) / original_size) * 100
        
        logging.info(f"Initial compression: {initial_ratio:.1f}% reduction")
        
        # If compression didn't work well, try more aggressive approach
        if initial_ratio < 0:  # File got bigger
            logging.info(f"File size increased, trying aggressive compression for '{file.filename}'")
            
            # Try with maximum compression settings
            aggressive_buffer = io.BytesIO()
            pdf_document.save(
                aggressive_buffer,
                garbage=4,      # Remove all unused objects
                deflate=True,   # Compress streams
                clean=True,     # Clean content streams
                linear=True,    # Optimize for web
                pretty=False,   # Remove formatting
                ascii=False     # Use binary instead of ASCII
            )
            
            aggressive_size = len(aggressive_buffer.getvalue())
            aggressive_ratio = ((original_size - aggressive_size) / original_size) * 100
            
            # Use the better result
            if aggressive_ratio > initial_ratio:
                output_buffer = aggressive_buffer
                compressed_size = aggressive_size
                compression_ratio = aggressive_ratio
                logging.info(f"Aggressive method better: {aggressive_ratio:.1f}% reduction")
            else:
                compressed_size = initial_compressed_size
                compression_ratio = initial_ratio
        else:
            compressed_size = initial_compressed_size
            compression_ratio = initial_ratio
        
        # Close the PDF document
        pdf_document.close()
        
        output_buffer.seek(0)
        
        # Calculate compression ratio
        original_size = len(pdf_bytes)
        compressed_size = len(output_buffer.getvalue())
        compression_ratio = ((original_size - compressed_size) / original_size) * 100
        
        # Check if the PDF was already well-optimized
        if compression_ratio < 5:  # Less than 5% reduction
            if compression_ratio < 0:
                logging.info(f"PDF '{file.filename}' appears to be already well-optimized or contains complex content that resists compression")
            else:
                logging.info(f"PDF '{file.filename}' achieved minimal compression - may already be optimized")
        
        # Special case for small PDFs
        if original_size < 100000:  # Less than 100KB
            logging.info(f"PDF '{file.filename}' is already small ({original_size/1024:.1f}KB) - compression may not provide significant benefits")
        
        logging.info(f"PDF compression: '{file.filename}' - Original: {original_size/1024:.1f}KB, Compressed: {compressed_size/1024:.1f}KB, Reduction: {compression_ratio:.1f}%")
        
        # If compression didn't work well, try alternative method
        if compression_ratio < 5:  # Less than 5% reduction
            logging.info(f"Low compression achieved, trying alternative method for '{file.filename}'")
            # Try with more aggressive settings
            pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
            alt_buffer = io.BytesIO()
            pdf_document.save(alt_buffer, garbage=4, deflate=True, clean=True, linear=True, pretty=False, ascii=False)
            pdf_document.close()
            alt_buffer.seek(0)
            
            alt_size = len(alt_buffer.getvalue())
            alt_ratio = ((original_size - alt_size) / original_size) * 100
            
            if alt_ratio > compression_ratio:
                output_buffer = alt_buffer
                compressed_size = alt_size
                compression_ratio = alt_ratio
                logging.info(f"Alternative method better: {alt_ratio:.1f}% reduction")
        
        logging.info(f"PDF compression: Successfully compressed '{file.filename}' with {compression_level} compression. Final reduction: {compression_ratio:.1f}%")
        
        # Generate output filename
        base_name = os.path.splitext(file.filename)[0]
        output_filename = f"compressed_{base_name}.pdf"
        
        return send_file(
            output_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=output_filename
        )

    except Exception as e:
        logging.error(f"PDF compression: Error processing '{file.filename}': {e}", exc_info=True)
        return jsonify({"error": f"Failed to compress PDF: {str(e)}"}), 500

# ADVANCED PDF COMPRESSION ENDPOINT
@app.route('/compress-pdf-advanced', methods=['POST'])
def compress_pdf_advanced():
    """
    Advanced PDF compression using multiple free libraries and smart algorithms
    Targets 50%+ compression for most PDFs
    """
    if 'file' not in request.files:
        logging.error("Advanced PDF compression: No file part in the request.")
        return jsonify({"error": "No file part in the request."}), 400

    file = request.files['file']
    if file.filename == '':
        logging.error("Advanced PDF compression: No selected file.")
        return jsonify({"error": "No selected file."}), 400
    if not file.filename.lower().endswith('.pdf'):
        logging.error(f"Advanced PDF compression: Invalid file type uploaded: {file.filename}")
        return jsonify({"error": "Invalid file type. Only PDF files are accepted."}), 400

    # Get compression parameters
    compression_level = request.form.get('compression_level', 'medium')
    
    try:
        file.stream.seek(0)
        pdf_bytes = file.read()
        original_size = len(pdf_bytes)
        
        logging.info(f"Advanced compression starting for '{file.filename}' - Original: {original_size/1024:.1f}KB")
        
        # Stage 1: Basic PyMuPDF compression
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        stage1_buffer = io.BytesIO()
        
        # Use aggressive settings for better compression
        pdf_document.save(
            stage1_buffer,
            garbage=4,      # Remove all unused objects
            deflate=True,   # Compress streams
            clean=True,     # Clean content streams
            linear=True,    # Optimize for web
            pretty=False,   # Remove formatting
            ascii=False     # Use binary instead of ASCII
        )
        
        stage1_size = len(stage1_buffer.getvalue())
        stage1_ratio = ((original_size - stage1_size) / original_size) * 100
        logging.info(f"Stage 1 (PyMuPDF): {stage1_ratio:.1f}% reduction")
        
        # Stage 2: Image compression (if images exist)
        stage2_buffer = io.BytesIO()
        try:
            # Check if PDF has images
            has_images = False
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                if page.get_images():
                    has_images = True
                    break
            
            if has_images and compression_level in ['medium', 'high']:
                logging.info("Stage 2: Compressing images within PDF")
                
                # Create a new PDF with compressed images
                new_pdf = fitz.open()
                
                for page_num in range(len(pdf_document)):
                    page = pdf_document[page_num]
                    new_page = new_pdf.new_page(width=page.rect.width, height=page.rect.height)
                    
                    # Copy page content
                    new_page.show_pdf_page(page.rect, pdf_document, page_num)
                    
                    # Compress images on this page
                    image_list = page.get_images()
                    for img_index, img in enumerate(image_list):
                        try:
                            xref = img[0]
                            img_info = pdf_document.extract_image(xref)
                            
                            if img_info and "image" in img_info:
                                # Get image size safely
                                img_size = img_info.get("size", 0)
                                if img_size > 50000:  # > 50KB
                                    # Compress image using Pillow
                                    from PIL import Image
                                    img_data = img_info["image"]
                                    img_pil = Image.open(io.BytesIO(img_data))
                                    
                                    # Determine compression quality based on level
                                    if compression_level == 'high':
                                        quality = 50  # More aggressive compression
                                    else:
                                        quality = 70  # Balanced compression
                                    
                                    # Convert to JPEG for better compression
                                    if img_pil.mode in ['RGBA', 'LA']:
                                        img_pil = img_pil.convert('RGB')
                                    
                                    # Resize large images for better compression
                                    if img_pil.width > 1200 or img_pil.height > 1200:
                                        img_pil.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
                                    
                                    # Compress image
                                    compressed_img_buffer = io.BytesIO()
                                    img_pil.save(compressed_img_buffer, 'JPEG', quality=quality, optimize=True, progressive=True)
                                    compressed_img_buffer.seek(0)
                                    
                                    # Replace image in PDF
                                    new_page.insert_image(page.rect, stream=compressed_img_buffer.getvalue())
                                    
                        except Exception as e:
                            logging.warning(f"Could not compress image {img_index} on page {page_num}: {e}")
                            continue
                
                # Save compressed PDF
                new_pdf.save(stage2_buffer, garbage=4, deflate=True, clean=True, linear=True)
                new_pdf.close()
                
                stage2_size = len(stage2_buffer.getvalue())
                stage2_ratio = ((original_size - stage2_size) / original_size) * 100
                logging.info(f"Stage 2 (Image compression): {stage2_ratio:.1f}% reduction")
                
                # Use stage 2 if it's better
                if stage2_ratio > stage1_ratio:
                    output_buffer = stage2_buffer
                    final_ratio = stage2_ratio
                    logging.info(f"Stage 2 selected: {stage2_ratio:.1f}% reduction")
                else:
                    output_buffer = stage1_buffer
                    final_ratio = stage1_ratio
                    logging.info(f"Stage 1 selected: {stage1_ratio:.1f}% reduction")
            else:
                output_buffer = stage1_buffer
                final_ratio = stage1_ratio
                logging.info(f"No images found, using Stage 1: {stage1_ratio:.1f}% reduction")
                
        except Exception as e:
            logging.warning(f"Stage 2 (image compression) failed: {e}")
            output_buffer = stage1_buffer
            final_ratio = stage1_ratio
        
        # Stage 3: Advanced optimization for high compression
        if compression_level == 'high' and final_ratio < 30:  # If we haven't achieved good compression
            logging.info("Stage 3: Advanced optimization techniques")
            
            try:
                # Try pikepdf for advanced compression
                import pikepdf
                
                # Convert to pikepdf format
                output_buffer.seek(0)
                pdf_pike = pikepdf.open(output_buffer)
                
                # Advanced compression settings (using correct parameters)
                pdf_pike.save(
                    output_buffer,
                    preserve_pdfa=False,   # Allow PDF/A restrictions to be removed
                    recompress_flate=True, # Recompress existing streams
                    deterministic_id=False  # Allow ID changes for better compression
                )
                
                final_size = len(output_buffer.getvalue())
                final_ratio = ((original_size - final_size) / original_size) * 100
                logging.info(f"Stage 3 (pikepdf): {final_ratio:.1f}% reduction")
                
            except ImportError:
                logging.info("pikepdf not available, skipping Stage 3")
            except Exception as e:
                logging.warning(f"Stage 3 (pikepdf) failed: {e}")
        
        # Stage 4: Content analysis and aggressive optimization
        if compression_level == 'high' and final_ratio < 20:  # Still not good enough
            logging.info("Stage 4: Content analysis and aggressive optimization")
            
            try:
                # Analyze PDF content and apply aggressive techniques
                pdf_document = fitz.open(stream=output_buffer.getvalue(), filetype="pdf")
                
                # Create new PDF with aggressive settings
                aggressive_pdf = fitz.open()
                
                for page_num in range(len(pdf_document)):
                    page = pdf_document[page_num]
                    
                    # Get page content
                    text_content = page.get_text()
                    image_list = page.get_images()
                    
                    # Create new page
                    new_page = aggressive_pdf.new_page(width=page.rect.width, height=page.rect.height)
                    
                    # If page has mostly text, optimize for text
                    if len(text_content) > 100 and len(image_list) < 3:
                        # Text-heavy page - use aggressive text optimization
                        new_page.insert_text((50, 50), text_content, fontsize=10)
                    else:
                        # Image-heavy page - copy with aggressive compression
                        new_page.show_pdf_page(page.rect, pdf_document, page_num)
                
                # Save with maximum compression
                aggressive_buffer = io.BytesIO()
                aggressive_pdf.save(
                    aggressive_buffer,
                    garbage=4,
                    deflate=True,
                    clean=True,
                    linear=True,
                    pretty=False,
                    ascii=False
                )
                
                aggressive_pdf.close()
                pdf_document.close()
                
                aggressive_size = len(aggressive_buffer.getvalue())
                aggressive_ratio = ((original_size - aggressive_size) / original_size) * 100
                
                if aggressive_ratio > final_ratio:
                    output_buffer = aggressive_buffer
                    final_ratio = aggressive_ratio
                    logging.info(f"Stage 4 (content analysis): {aggressive_ratio:.1f}% reduction")
                
            except Exception as e:
                logging.warning(f"Stage 4 (content analysis) failed: {e}")
        
        # Stage 5: Final optimization - metadata removal and font optimization
        if compression_level == 'high':
            logging.info("Stage 5: Final optimization - metadata and font optimization")
            
            try:
                # Try to remove metadata and optimize fonts
                final_pdf = fitz.open(stream=output_buffer.getvalue(), filetype="pdf")
                
                # Remove metadata if possible
                if hasattr(final_pdf, 'metadata'):
                    try:
                        # Keep only essential metadata
                        essential_metadata = {
                            'title': final_pdf.metadata.get('title', ''),
                            'author': final_pdf.metadata.get('author', ''),
                            'subject': final_pdf.metadata.get('subject', '')
                        }
                        final_pdf.set_metadata(essential_metadata)
                    except:
                        pass
                
                # Final save with maximum compression
                final_buffer = io.BytesIO()
                final_pdf.save(
                    final_buffer,
                    garbage=4,
                    deflate=True,
                    clean=True,
                    linear=True,
                    pretty=False,
                    ascii=False
                )
                
                final_pdf.close()
                
                final_size = len(final_buffer.getvalue())
                final_ratio = ((original_size - final_size) / original_size) * 100
                
                # Use the final optimized version
                output_buffer = final_buffer
                logging.info(f"Stage 5 (final optimization): {final_ratio:.1f}% reduction")
                
            except Exception as e:
                logging.warning(f"Stage 5 (final optimization) failed: {e}")
        
        # Close the PDF document
        pdf_document.close()
        
        # Final size calculation
        output_buffer.seek(0)
        final_size = len(output_buffer.getvalue())
        final_ratio = ((original_size - final_size) / original_size) * 100
        
        # Generate output filename
        base_name = os.path.splitext(file.filename)[0]
        output_filename = f"compressed_{base_name}.pdf"
        
        logging.info(f"Advanced PDF compression: '{file.filename}' - Original: {original_size/1024:.1f}KB, Final: {final_size/1024:.1f}KB, Total Reduction: {final_ratio:.1f}%")
        
        return send_file(
            output_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=output_filename
        )

    except Exception as e:
        logging.error(f"Advanced PDF compression: Error processing '{file.filename}': {e}", exc_info=True)
        return jsonify({"error": f"Failed to compress PDF: {str(e)}"}), 500

# Main entry point
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=4000)