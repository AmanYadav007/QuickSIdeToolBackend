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
from PIL import Image

# Import Adobe service
from adobe_service import adobe_service

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="QuickSideTool PDF Security API",
    description="Professional PDF security and manipulation service with Adobe integration",
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
    allowed_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']
    return get_file_extension(file.filename) in allowed_extensions

def validate_word_file(file: UploadFile) -> bool:
    """Validate Word file"""
    allowed_extensions = ['.docx', '.doc']
    return get_file_extension(file.filename) in allowed_extensions

# Root endpoint
@app.get("/")
async def root():
    return {
        "message": "QuickSideTool PDF Security API",
        "version": "2.0.0",
        "endpoints": {
            "unlock_pdf": "/unlock-pdf",
            "lock_pdf": "/lock-pdf", 
            "remove_pdf_links": "/remove-pdf-links",
            "adobe_pdf_to_word": "/adobe/convert/pdf-to-word",
            "adobe_pdf_to_excel": "/adobe/convert/pdf-to-excel",
            "adobe_compress_pdf": "/adobe/compress-pdf",
            "adobe_optimize_pdf": "/adobe/optimize-pdf",
            "adobe_extract_text": "/adobe/extract-text",
            "health": "/health"
        }
    }

# Health check endpoint
@app.get("/health")
async def health_check():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

# PDF Unlock endpoint
@app.post("/unlock-pdf")
async def unlock_pdf_legacy(file: UploadFile = File(...), password: str = Form(...)):
    """Unlock PDF with password"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    try:
        # Create temporary file
        input_path = create_temp_file('.pdf')
        output_path = create_temp_file('.pdf')
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Unlock PDF using pikepdf
        with pikepdf.open(input_path, password=password) as pdf:
            pdf.save(output_path)
        
        # Generate output filename
        output_filename = f"unlocked_{file.filename}"
        
        logger.info(f"Successfully unlocked {file.filename}")
        
        # Return the unlocked file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/pdf'
        )
        
    except pikepdf.PasswordError:
        raise HTTPException(status_code=400, detail="Incorrect password for this PDF.")
    except pikepdf.PdfError as e:
        if "not encrypted" in str(e).lower():
            raise HTTPException(status_code=400, detail="This PDF is not encrypted.")
        else:
            raise HTTPException(status_code=400, detail=f"PDF error: {str(e)}")
    except Exception as e:
        logger.error(f"Error unlocking PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to unlock PDF: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

# PDF Lock endpoint
@app.post("/lock-pdf")
async def lock_pdf_legacy(file: UploadFile = File(...), password: str = Form(...)):
    """Lock PDF with password"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    try:
        # Create temporary file
        input_path = create_temp_file('.pdf')
        output_path = create_temp_file('.pdf')
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Check if PDF is already encrypted
        try:
            with pikepdf.open(input_path) as pdf:
                if pdf.is_encrypted:
                    raise HTTPException(status_code=400, detail="PDF is already encrypted.")
        except pikepdf.PasswordError:
            raise HTTPException(status_code=400, detail="PDF is already encrypted.")
        
        # Lock PDF using pikepdf
        with pikepdf.open(input_path) as pdf:
            pdf.save(output_path, encrypt=True, password=password)
        
        # Generate output filename
        output_filename = f"locked_{file.filename}"
        
        logger.info(f"Successfully locked {file.filename}")
        
        # Return the locked file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/pdf'
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error locking PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to lock PDF: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

# PDF Link Removal endpoint
@app.post("/remove-pdf-links")
async def remove_pdf_links_legacy(file: UploadFile = File(...)):
    """Remove links and hyperlinks from PDF"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    try:
        # Create temporary file
        input_path = create_temp_file('.pdf')
        output_path = create_temp_file('.pdf')
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Remove links using PyMuPDF
        doc = fitz.open(input_path)
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Get all links on the page
            links = page.get_links()
            
            # Remove each link
            for link in links:
                page.delete_link(link)
        
        # Save the modified PDF
        doc.save(output_path)
        doc.close()
        
        # Generate output filename
        output_filename = f"links_removed_{file.filename}"
        
        logger.info(f"Successfully removed links from {file.filename}")
        
        # Return the modified file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/pdf'
        )
        
    except Exception as e:
        logger.error(f"Error removing PDF links: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to remove links: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

# Adobe Enhanced Endpoints

@app.post("/adobe/convert/pdf-to-word")
async def adobe_convert_pdf_to_word(file: UploadFile = File(...)):
    """Convert PDF to Word document using Adobe PDF Services (Professional quality)"""
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
        
        # Use Adobe service for conversion
        success = await adobe_service.convert_pdf_to_word(input_path, output_path)
        
        if success:
            # Generate output filename
            output_filename = file.filename.replace('.pdf', '.docx')
            if not output_filename.endswith('.docx'):
                output_filename += '.docx'
            
            logger.info(f"Successfully converted {file.filename} to Word using Adobe")
            
            # Return the converted file
            return FileResponse(
                path=output_path,
                filename=output_filename,
                media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            raise HTTPException(status_code=500, detail="Adobe conversion failed. Please try again.")
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error converting PDF to Word with Adobe: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

@app.post("/adobe/convert/pdf-to-excel")
async def adobe_convert_pdf_to_excel(file: UploadFile = File(...)):
    """Convert PDF to Excel using Adobe PDF Services (Professional quality)"""
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
        
        # Use Adobe service for conversion
        success = await adobe_service.convert_pdf_to_excel(input_path, output_path)
        
        if success:
            # Generate output filename
            output_filename = file.filename.replace('.pdf', '.xlsx')
            if not output_filename.endswith('.xlsx'):
                output_filename += '.xlsx'
            
            logger.info(f"Successfully converted {file.filename} to Excel using Adobe")
            
            # Return the converted file
            return FileResponse(
                path=output_path,
                filename=output_filename,
                media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            raise HTTPException(status_code=500, detail="Adobe conversion failed. Please try again.")
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error converting PDF to Excel with Adobe: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

@app.post("/adobe/compress-pdf")
async def adobe_compress_pdf(
    file: UploadFile = File(...),
    compression_level: str = Form("medium")
):
    """Compress PDF using Adobe PDF Services (Professional quality)"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    if compression_level not in ['low', 'medium', 'high']:
        raise HTTPException(status_code=400, detail="Invalid compression level. Use: low, medium, or high.")
    
    try:
        # Create temporary files
        input_path = create_temp_file('.pdf')
        output_path = create_temp_file('.pdf')
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Use Adobe service for compression
        success = await adobe_service.compress_pdf(input_path, output_path, compression_level)
        
        if success:
            # Generate output filename
            output_filename = f"compressed_{file.filename}"
            
            logger.info(f"Successfully compressed {file.filename} using Adobe")
            
            # Return the compressed file
            return FileResponse(
                path=output_path,
                filename=output_filename,
                media_type='application/pdf'
            )
        else:
            raise HTTPException(status_code=500, detail="Adobe compression failed. Please try again.")
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error compressing PDF with Adobe: {e}")
        raise HTTPException(status_code=500, detail=f"Compression failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

@app.post("/adobe/optimize-pdf")
async def adobe_optimize_pdf(file: UploadFile = File(...)):
    """Optimize PDF for web viewing using Adobe PDF Services"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    try:
        # Create temporary files
        input_path = create_temp_file('.pdf')
        output_path = create_temp_file('.pdf')
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Use Adobe service for optimization
        success = await adobe_service.optimize_pdf_for_web(input_path, output_path)
        
        if success:
            # Generate output filename
            output_filename = f"optimized_{file.filename}"
            
            logger.info(f"Successfully optimized {file.filename} using Adobe")
            
            # Return the optimized file
            return FileResponse(
                path=output_path,
                filename=output_filename,
                media_type='application/pdf'
            )
        else:
            raise HTTPException(status_code=500, detail="Adobe optimization failed. Please try again.")
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error optimizing PDF with Adobe: {e}")
        raise HTTPException(status_code=500, detail=f"Optimization failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

@app.post("/adobe/extract-text")
async def adobe_extract_text(file: UploadFile = File(...)):
    """Extract text from PDF using OCR with Adobe Document Services"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size(file.size):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")
    
    try:
        # Create temporary file
        input_path = create_temp_file('.pdf')
        
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Use Adobe service for text extraction
        extracted_text = await adobe_service.extract_text_with_ocr(input_path)
        
        if extracted_text:
            # Generate output filename
            output_filename = file.filename.replace('.pdf', '.txt')
            if not output_filename.endswith('.txt'):
                output_filename += '.txt'
            
            # Create temporary text file
            output_path = create_temp_file('.txt')
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(extracted_text)
            
            logger.info(f"Successfully extracted text from {file.filename} using Adobe OCR")
            
            # Return the text file
            return FileResponse(
                path=output_path,
                filename=output_filename,
                media_type='text/plain'
            )
        else:
            raise HTTPException(status_code=500, detail="Adobe text extraction failed. Please try again.")
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error extracting text with Adobe: {e}")
        raise HTTPException(status_code=500, detail=f"Text extraction failed: {str(e)}")
    
    finally:
        # Cleanup temporary files
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=4000)