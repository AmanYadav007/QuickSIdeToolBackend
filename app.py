from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
import tempfile
import os
import shutil
import uuid
import logging
from typing import List, Optional
import asyncio
from datetime import datetime, timedelta
 

# Import conversion libraries
import pikepdf
import fitz  # PyMuPDF
 

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
    allow_origins=["*"],  # For production, restrict to your domains
    allow_credentials=False,  # Wildcard cannot be used with credentials
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
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
def get_upload_file_size(upload_file: UploadFile) -> int:
    """Safely determine the size of an UploadFile without consuming it"""
    try:
        current_position = upload_file.file.tell()
    except Exception:
        current_position = 0
    try:
        upload_file.file.seek(0, os.SEEK_END)
        size = upload_file.file.tell()
    finally:
        try:
            upload_file.file.seek(current_position, os.SEEK_SET)
        except Exception:
            # Best effort reset
            try:
                upload_file.file.seek(0)
            except Exception:
                pass
    return size

def validate_file_size_upload(upload_file: UploadFile) -> bool:
    """Validate file size for an UploadFile instance"""
    try:
        return get_upload_file_size(upload_file) <= MAX_FILE_SIZE
    except Exception:
        return False

 

def create_temp_file(extension: str = "") -> str:
    """Create a temporary file path"""
    unique_id = str(uuid.uuid4())
    return os.path.join(TEMP_DIR, f"{unique_id}{extension}")

def validate_pdf_file(file: UploadFile) -> bool:
    """Validate PDF file"""
    if not file.filename.lower().endswith('.pdf'):
        return False
    return True

 

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
            "convert_pdf_to_word": "/convert/pdf-to-word",
            "convert_pdf_to_excel": "/convert/pdf-to-excel",
            "convert_pdf_to_text": "/convert/pdf-to-text",
            "convert_word_to_pdf": "/convert/word-to-pdf",
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
    
    if not validate_file_size_upload(file):
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
    
    if not validate_file_size_upload(file):
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
        
        # Lock (encrypt) PDF using pikepdf
        with pikepdf.open(input_path) as pdf:
            encryption = pikepdf.Encryption(user=password, owner=password, R=4)
            pdf.save(output_path, encryption=encryption)
        
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
    
    if not validate_file_size_upload(file):
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

# Basic Endpoints

@app.post("/compress-pdf")
async def compress_pdf(file: UploadFile = File(...)):
    """Compress PDF using PyMuPDF"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")

    if not validate_file_size_upload(file):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")

    input_path = create_temp_file('.pdf')
    output_path = create_temp_file('.pdf')

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = fitz.open(input_path)
        # Save with garbage collection, deflation, and cleaning
        doc.save(output_path, garbage=4, deflate=True, clean=True)
        doc.close()

        output_filename = f"compressed_{file.filename}"
        logger.info(f"Successfully compressed {file.filename} using basic compression")

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/pdf'
        )
    except Exception as e:
        logger.error(f"Error compressing PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to compress PDF: {str(e)}")
    finally:
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

@app.post("/convert/pdf-to-word")
async def convert_pdf_to_word(file: UploadFile = File(...)):
    """Convert PDF to Word document using pdf2docx"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")

    if not validate_file_size_upload(file):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")

    input_path = create_temp_file('.pdf')
    output_path = create_temp_file('.docx')

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # First attempt: pdf2docx
        try:
            from pdf2docx import Converter
            cv = Converter(input_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            method = "pdf2docx"
        except Exception as primary_error:
            logger.warning(f"pdf2docx failed, falling back to text-based DOCX: {primary_error}")
            # Fallback: extract text with PyMuPDF and write a simple DOCX
            try:
                import fitz  # PyMuPDF
                from docx import Document  # python-docx

                document = Document()
                with fitz.open(input_path) as doc:
                    for page in doc:
                        text = page.get_text()
                        if text.strip():
                            for line in text.splitlines():
                                document.add_paragraph(line)
                        else:
                            document.add_paragraph(" ")
                document.save(output_path)
                method = "text-fallback"
            except Exception as fallback_error:
                logger.error(f"Fallback text DOCX creation failed: {fallback_error}")
                raise HTTPException(status_code=500, detail="Conversion failed. Try a different file.")

        output_filename = file.filename.replace('.pdf', '.docx')
        if not output_filename.endswith('.docx'):
            output_filename += '.docx'
        
        logger.info(f"Successfully converted {file.filename} to Word using {method}")

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error converting PDF to Word: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to convert PDF to Word: {str(e)}")
    finally:
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

@app.post("/convert/pdf-to-excel")
async def convert_pdf_to_excel(file: UploadFile = File(...)):
    """Convert PDF tables to Excel using pdfplumber + pandas (basic)"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")

    if not validate_file_size_upload(file):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")

    input_path = create_temp_file('.pdf')
    output_path = create_temp_file('.xlsx')

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        import pdfplumber
        import pandas as pd

        tables_found = 0
        with pdfplumber.open(input_path) as pdf:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                for page_index, page in enumerate(pdf.pages, start=1):
                    try:
                        # stream mode is more forgiving; lattice needs vector lines
                        extracted_tables = page.extract_tables()
                        if not extracted_tables:
                            # try explicit settings to improve detection
                            extracted_tables = page.extract_tables(table_settings={
                                'vertical_strategy': 'lines',
                                'horizontal_strategy': 'lines',
                                'intersection_tolerance': 5
                            })
                        for table_index, table in enumerate(extracted_tables, start=1):
                            if not table:
                                continue
                            df = pd.DataFrame(table)
                            sheet_name = f"p{page_index}_t{table_index}"
                            df.to_excel(writer, index=False, header=False, sheet_name=sheet_name)
                            tables_found += 1
                    except Exception:
                        # skip problematic pages but continue others
                        continue

        if tables_found == 0:
            raise HTTPException(status_code=400, detail="No tables detected in PDF. Try Adobe mode or another file.")

        output_filename = file.filename.replace('.pdf', '.xlsx')
        if not output_filename.endswith('.xlsx'):
            output_filename += '.xlsx'

        logger.info(f"Successfully converted {file.filename} to Excel using basic conversion with {tables_found} tables")

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error converting PDF to Excel: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to convert PDF to Excel: {str(e)}")
    finally:
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

@app.post("/convert/pdf-to-text")
async def convert_pdf_to_text(file: UploadFile = File(...)):
    """Extract text from PDF using pdfplumber"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")

    if not validate_file_size_upload(file):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")

    input_path = create_temp_file('.pdf')
    output_path = create_temp_file('.txt')

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        import pdfplumber
        text_parts = []
        with pdfplumber.open(input_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                text_parts.append(text)
        full_text = "\n\n".join(text_parts).strip()

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_text)

        output_filename = file.filename.replace('.pdf', '.txt')
        if not output_filename.endswith('.txt'):
            output_filename += '.txt'

        logger.info(f"Extracted text from {file.filename} ({len(full_text)} chars)")

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='text/plain'
        )
    except Exception as e:
        logger.error(f"Error extracting text: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to extract text: {str(e)}")
    finally:
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

@app.post("/convert/word-to-pdf")
async def convert_word_to_pdf(file: UploadFile = File(...)):
    """Convert Word document to PDF using LibreOffice headless"""
    allowed = ['.docx', '.doc']
    if not any(file.filename.lower().endswith(ext) for ext in allowed):
        raise HTTPException(status_code=400, detail="Invalid file type. Only .docx or .doc accepted.")

    if not validate_file_size_upload(file):
        raise HTTPException(status_code=400, detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.")

    import subprocess

    input_ext = os.path.splitext(file.filename)[1].lower()
    input_path = create_temp_file(input_ext)
    output_dir = TEMP_DIR

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Run LibreOffice headless conversion
        cmd = [
            "soffice", "--headless", "--convert-to", "pdf",
            "--outdir", output_dir, input_path
        ]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=120)
        if result.returncode != 0:
            logger.error(f"LibreOffice error: {result.stderr}")
            raise HTTPException(status_code=500, detail="LibreOffice conversion failed")

        output_path = os.path.splitext(input_path)[0] + '.pdf'
        if not os.path.exists(output_path):
            # Some LibreOffice versions keep original basename
            candidate = os.path.join(output_dir, os.path.basename(os.path.splitext(input_path)[0]) + '.pdf')
            if os.path.exists(candidate):
                output_path = candidate
            else:
                raise HTTPException(status_code=500, detail="Converted file not found")

        output_filename = file.filename.rsplit('.', 1)[0] + '.pdf'

        logger.info(f"Converted {file.filename} to PDF via LibreOffice")

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type='application/pdf'
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error converting Word to PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to convert Word to PDF: {str(e)}")
    finally:
        # Do not remove output_path here as FileResponse serves from disk; lifecycle cleanup task will remove old files
        if os.path.exists(input_path):
            try:
                os.remove(input_path)
            except:
                pass

# Adobe Enhanced Endpoints

@app.post("/adobe/convert/pdf-to-word")
async def adobe_convert_pdf_to_word(file: UploadFile = File(...)):
    """Convert PDF to Word document using Adobe PDF Services (Professional quality)"""
    if not validate_pdf_file(file):
        raise HTTPException(status_code=400, detail="Invalid file type. Only PDF files are accepted.")
    
    if not validate_file_size_upload(file):
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
    
    if not validate_file_size_upload(file):
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
    
    if not validate_file_size_upload(file):
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

 

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=4000)