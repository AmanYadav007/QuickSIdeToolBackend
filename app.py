from flask import Flask, request, send_file
from PyPDF2 import PdfReader, PdfWriter
from flask_cors import CORS
import io
import logging

# Initialize Flask app
app = Flask(__name__)
CORS(app)  # Enable CORS for cross-origin requests

# Configure logging
logging.basicConfig(level=logging.DEBUG)

# Error handler for global exceptions
@app.errorhandler(Exception)
def handle_exception(e):
    logging.error(f"An error occurred: {str(e)}")
    return "Internal Server Error", 500

# Unlock PDF endpoint
@app.route('/unlock-pdf', methods=['POST'])
def unlock_pdf():
    if 'file' not in request.files or 'password' not in request.form:
        return "File or password missing", 400

    file = request.files['file']
    password = request.form.get('password')

    try:
        reader = PdfReader(file)
        if reader.is_encrypted:
            if not reader.decrypt(password):
                return "Incorrect password", 400

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
            download_name='unlocked.pdf'
        )
    except Exception as e:
        logging.error(f"Error unlocking PDF: {e}")
        return str(e), 400

# Lock PDF endpoint
@app.route('/lock-pdf', methods=['POST'])
def lock_pdf():
    if 'file' not in request.files or 'password' not in request.form:
        return "File or password missing", 400

    file = request.files['file']
    password = request.form.get('password')

    if not password:
        return "Password is required", 400

    try:
        reader = PdfReader(file)
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
            download_name='locked.pdf'
        )
    except Exception as e:
        logging.error(f"Error locking PDF: {e}")
        return str(e), 400

# Main entry point
if __name__ == '__main__':
    # Run Flask app on port 4000
    app.run(debug=True, host='0.0.0.0', port=4000)
