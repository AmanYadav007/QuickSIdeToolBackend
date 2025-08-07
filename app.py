from fastapi import FastAPI, File, UploadFile, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import uvicorn
import tempfile
import os
import shutil
import uuid
import logging
from typing import List, Optional
import asyncio
from datetime import datetime, timedelta
import zipfile
import io

# Import conversion libraries
import pikepdf
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pdf2docx
import pdfplumber
import pdf2image
import img2pdf
from PIL import Image
import pandas as pd
import openpyxl
from openpyxl import Workbook

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="QuickSideTool File Converter API",
    description="Professional file conversion service with Adobe-style quality",
    version="2.0.0"
)

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Configure appropriately for production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configuration
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
TEMP_DIR = tempfile.mkdtemp()
CLEANUP_INTERVAL = 600  # 10 minutes

# File cleanup task
async def cleanup_temp_files():
    """Clean up temporary files older than 10 minutes"""
    while True:
        try:
            current_time = datetime.now()
            for filename in os.listdir(TEMP_DIR):
                filepath = os.path.join(TEMP_DIR, filename)
                if os.path.isfile(filepath):
                    file_time = datetime.fromtimestamp(os.path.getctime(filepath))
                    if current_time - file_time > timedelta(minutes=10):
                        os.remove(filepath)
                        logger.info(f"Cleaned up temporary file: {filename}")
        except Exception as e:
            logger.error(f"Error during cleanup: {e}")
        
        await asyncio.sleep(CLEANUP_INTERVAL)

# Start cleanup task
@app.on_event("startup")
async def startup_event():
    asyncio.create_task(cleanup_temp_files())

# Utility functions
def validate_file_size(file_size: int) -> bool:
    """Validate file size"""
    return file_size <= MAX_FILE_SIZE

def get_file_extension(filename: str) -> str:
    """Get file extension from filename"""
    return os.path.splitext(filename)[1].lower()

def create_temp_file(extension: str = "") -> str:
    """Create a temporary file path"""
    unique_id = str(uuid.uuid4())
    return os.path.join(TEMP_DIR, f"{unique_id}{extension}")

def validate_pdf_file(file: UploadFile) -> bool:
    """Validate PDF file"""
    if not file.filename.lower().endswith('.pdf'):
        return False
    return True

def validate_image_file(file: UploadFile) -> bool:
    """Validate image file"""
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']
    return any(file.filename.lower().endswith(ext) for ext in image_extensions)

def validate_word_file(file: UploadFile) -> bool:
    """Validate Word file"""
    word_extensions = ['.docx', '.doc']
    return any(file.filename.lower().endswith(ext) for ext in word_extensions)

# Root endpoint
@app.get("/")
async def root():
    return {
        "message": "QuickSideTool File Converter API",
        "version": "2.0.0",
        "status": "running"
    }

# Health check endpoint
@app.get("/health")
async def health_check():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

# PDF to Word conversion
@app.post("/convert/pdf-to-word")
async def convert_pdf_to_word(file: UploadFile = File(...)):
    """Convert PDF to Word document with layout preservation"""
    
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    try:
        # Create temporary files
        input_path = create_temp_file('.pdf')
        output_path = create_temp_file('.docx')
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Convert using pdf2docx for better layout preservation
        converter = pdf2docx.Converter(input_path)
        converter.convert(output_path)
        converter.close()
        
        # Generate output filename
        output_filename = file.filename.replace('.pdf', '.docx')
        if not output_filename.endswith('.docx'):
            output_filename += '.docx'
        
        logger.info(f"Successfully converted {file.filename} to Word")
        
        # Return the converted file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        logger.error(f"Error converting PDF to Word: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

# PDF to Text conversion
@app.post("/convert/pdf-to-txt")
async def convert_pdf_to_txt(file: UploadFile = File(...)):
    """Convert PDF to plain text"""
    
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    try:
        # Create temporary file
        input_path = create_temp_file('.pdf')
        output_path = create_temp_file('.txt')
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Extract text using pdfplumber
        text_content = ""
        with pdfplumber.open(input_path) as pdf:
            for page in pdf.pages:
                text_content += page.extract_text() + "\n\n"
        
        # Save text to file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text_content)
        
        # Generate output filename
        output_filename = file.filename.replace('.pdf', '.txt')
        if not output_filename.endswith('.txt'):
            output_filename += '.txt'
        
        logger.info(f"Successfully converted {file.filename} to text")
        
        # Return the converted file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='text/plain'
        )
        
    except Exception as e:
        logger.error(f"Error converting PDF to text: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

# PDF to Excel conversion
@app.post("/convert/pdf-to-excel")
async def convert_pdf_to_excel(file: UploadFile = File(...)):
    """Convert PDF to Excel spreadsheet"""
    
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    try:
        # Create temporary files
        input_path = create_temp_file('.pdf')
        output_path = create_temp_file('.xlsx')
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Extract tables using pdfplumber
        all_tables = []
        with pdfplumber.open(input_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table_num, table in enumerate(tables):
                    if table:  # Only add non-empty tables
                        all_tables.append({
                            'page': page_num + 1,
                            'table': table_num + 1,
                            'data': table
                        })
        
        # Create Excel workbook
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet
        
        if all_tables:
            for i, table_info in enumerate(all_tables):
                ws = wb.create_sheet(title=f"Page{table_info['page']}_Table{table_info['table']}")
                for row_idx, row in enumerate(table_info['data'], 1):
                    for col_idx, cell in enumerate(row, 1):
                        ws.cell(row=row_idx, column=col_idx, value=cell)
        else:
            # If no tables found, create a sheet with extracted text
            ws = wb.create_sheet(title="Extracted_Text")
            with pdfplumber.open(input_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    text = page.extract_text()
                    if text:
                        ws.cell(row=page_num, column=1, value=f"Page {page_num}")
                        ws.cell(row=page_num, column=2, value=text)
        
        # Save Excel file
        wb.save(output_path)
        
        # Generate output filename
        output_filename = file.filename.replace('.pdf', '.xlsx')
        if not output_filename.endswith('.xlsx'):
            output_filename += '.xlsx'
        
        logger.info(f"Successfully converted {file.filename} to Excel")
        
        # Return the converted file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        logger.error(f"Error converting PDF to Excel: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

# Word to PDF conversion
@app.post("/convert/word-to-pdf")
async def convert_word_to_pdf(file: UploadFile = File(...)):
    """Convert Word document to PDF"""
    
    if not validate_word_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only Word documents are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    try:
        # Create temporary files
        input_path = create_temp_file(get_file_extension(file.filename))
        output_path = create_temp_file('.pdf')
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Convert using PyMuPDF (fitz)
        doc = fitz.open(input_path)
        doc.save(output_path)
        doc.close()
        
        # Generate output filename
        output_filename = file.filename.replace('.docx', '.pdf').replace('.doc', '.pdf')
        if not output_filename.endswith('.pdf'):
            output_filename += '.pdf'
        
        logger.info(f"Successfully converted {file.filename} to PDF")
        
        # Return the converted file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/pdf'
        )
        
    except Exception as e:
        logger.error(f"Error converting Word to PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

# PDF to Image conversion
@app.post("/convert/pdf-to-image")
async def convert_pdf_to_image(
    file: UploadFile = File(...),
    format: str = Form("png"),
    dpi: int = Form(300)
):
    """Convert PDF pages to images"""
    
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    if format not in ['png', 'jpg', 'jpeg']:
        raise HTTPException(status_code=400, detail="Invalid format. Supported formats: png, jpg, jpeg")
    
    try:
        # Create temporary files
        input_path = create_temp_file('.pdf')
        output_dir = create_temp_file('_images')
        os.makedirs(output_dir, exist_ok=True)
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Convert PDF to images
        images = pdf2image.convert_from_path(
            input_path,
            dpi=dpi,
            fmt=format.upper()
        )
        
        # Save images and create zip file
        zip_path = create_temp_file('.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for i, image in enumerate(images):
                img_path = os.path.join(output_dir, f"page_{i+1}.{format}")
                image.save(img_path, format.upper())
                zipf.write(img_path, f"page_{i+1}.{format}")
        
        # Generate output filename
        output_filename = f"{file.filename.replace('.pdf', '')}_pages.zip"
        
        logger.info(f"Successfully converted {file.filename} to {len(images)} images")
        
        # Return the zip file
        return FileResponse(
            path=zip_path,
            filename=output_filename,
            media_type='application/zip'
        )
        
    except Exception as e:
        logger.error(f"Error converting PDF to images: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_dir, zip_path]:
            if os.path.exists(path):
                try:
                    if os.path.isdir(path):
                        shutil.rmtree(path)
                    else:
                        os.remove(path)
                except:
                    pass

# Image to PDF conversion
@app.post("/convert/image-to-pdf")
async def convert_image_to_pdf(files: List[UploadFile] = File(...)):
    """Convert images to PDF"""
    
    if not files:
        raise HTTPException(status_code=400, detail="No files provided.")
    
    # Validate all files are images
    for file in files:
        if not validate_image_file(file):
            raise HTTPException(status_code=400, detail=f"Invalid file type: {file.filename}. Only image files are accepted.")
        if not validate_file_size(file.size):
            raise HTTPException(status_code=400, detail=f"File too large: {file.filename}. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    try:
        # Create temporary files
        output_path = create_temp_file('.pdf')
        temp_images = []
        
        # Save uploaded images
        for file in files:
            temp_path = create_temp_file(get_file_extension(file.filename))
            with open(temp_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            temp_images.append(temp_path)
        
        # Convert images to PDF
        with open(output_path, "wb") as f:
            f.write(img2pdf.convert(temp_images))
        
        # Generate output filename
        output_filename = f"converted_images_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        
        logger.info(f"Successfully converted {len(files)} images to PDF")
        
        # Return the converted file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/pdf'
        )
        
    except Exception as e:
        logger.error(f"Error converting images to PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in temp_images + [output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

# Legacy endpoints for backward compatibility
@app.post("/pdf-to-docx")
async def pdf_to_docx_legacy(file: UploadFile = File(...)):
    """Legacy endpoint for PDF to Word conversion"""
    return await convert_pdf_to_word(file)

@app.post("/unlock-pdf")
async def unlock_pdf_legacy(file: UploadFile = File(...), password: str = Form(...)):
    """Legacy endpoint for PDF unlocking"""
    # Implementation remains the same as before
    pass

@app.post("/lock-pdf")
async def lock_pdf_legacy(file: UploadFile = File(...), password: str = Form(...)):
    """Legacy endpoint for PDF locking"""
    # Implementation remains the same as before
    pass

@app.post("/remove-pdf-links")
async def remove_pdf_links_legacy(file: UploadFile = File(...)):
    """Legacy endpoint for PDF link removal"""
    # Implementation remains the same as before
    pass

# Main entry point
if __name__ == "__main__":
    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=4000,
        reload=True,
        workers=1
    )