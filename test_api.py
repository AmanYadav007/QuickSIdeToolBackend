#!/usr/bin/env python3
"""
Test script for QuickSideTool PDF Security API
"""

import requests
import os
import tempfile
from pathlib import Path

# API base URL
BASE_URL = "http://localhost:4000"

def test_health_check():
    """Test the health check endpoint"""
    print("🔍 Testing health check...")
    try:
        response = requests.get(f"{BASE_URL}/health")
        if response.status_code == 200:
            print("✅ Health check passed")
            print(f"   Response: {response.json()}")
        else:
            print(f"❌ Health check failed: {response.status_code}")
    except Exception as e:
        print(f"❌ Health check error: {e}")

def test_root_endpoint():
    """Test the root endpoint"""
    print("\n🔍 Testing root endpoint...")
    try:
        response = requests.get(f"{BASE_URL}/")
        if response.status_code == 200:
            print("✅ Root endpoint passed")
            print(f"   Response: {response.json()}")
        else:
            print(f"❌ Root endpoint failed: {response.status_code}")
    except Exception as e:
        print(f"❌ Root endpoint error: {e}")

def create_test_pdf():
    """Create a simple test PDF file"""
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        
        # Create temporary PDF
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_path = temp_file.name
        temp_file.close()
        
        # Generate simple PDF
        c = canvas.Canvas(temp_path, pagesize=letter)
        c.drawString(100, 750, "Test PDF Document")
        c.drawString(100, 700, "This is a test PDF for API testing.")
        c.drawString(100, 650, "It contains some sample text content.")
        c.save()
        
        return temp_path
    except ImportError:
        print("⚠️  reportlab not available, creating simple text file instead")
        # Create a simple text file as fallback
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.txt')
        temp_file.write(b"This is a test document for API testing.\nIt contains some sample text content.")
        temp_file.close()
        return temp_file.name

def test_pdf_unlock():
    """Test PDF unlock functionality"""
    print("\n🔍 Testing PDF unlock...")
    
    pdf_path = create_test_pdf()
    if not pdf_path:
        print("⚠️  Skipping PDF unlock test")
        return
    
    try:
        with open(pdf_path, 'rb') as f:
            files = {'file': ('test.pdf', f, 'application/pdf')}
            data = {'password': 'test123'}
            response = requests.post(f"{BASE_URL}/unlock-pdf", files=files, data=data)
        
        if response.status_code == 400 and "not encrypted" in response.text:
            print("✅ PDF unlock test passed (correctly detected unencrypted PDF)")
        else:
            print(f"❌ PDF unlock test failed: {response.status_code}")
            print(f"   Response: {response.text}")
    except Exception as e:
        print(f"❌ PDF unlock test error: {e}")
    finally:
        # Cleanup
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)

def test_pdf_lock():
    """Test PDF lock functionality"""
    print("\n🔍 Testing PDF lock...")
    
    pdf_path = create_test_pdf()
    if not pdf_path:
        print("⚠️  Skipping PDF lock test")
        return
    
    try:
        with open(pdf_path, 'rb') as f:
            files = {'file': ('test.pdf', f, 'application/pdf')}
            data = {'password': 'test123'}
            response = requests.post(f"{BASE_URL}/lock-pdf", files=files, data=data)
        
        if response.status_code == 200:
            print("✅ PDF lock test passed")
            print(f"   Response size: {len(response.content)} bytes")
        else:
            print(f"❌ PDF lock test failed: {response.status_code}")
            print(f"   Response: {response.text}")
    except Exception as e:
        print(f"❌ PDF lock test error: {e}")
    finally:
        # Cleanup
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)

def test_remove_pdf_links():
    """Test PDF link removal functionality"""
    print("\n🔍 Testing PDF link removal...")
    
    pdf_path = create_test_pdf()
    if not pdf_path:
        print("⚠️  Skipping PDF link removal test")
        return
    
    try:
        with open(pdf_path, 'rb') as f:
            files = {'file': ('test.pdf', f, 'application/pdf')}
            response = requests.post(f"{BASE_URL}/remove-pdf-links", files=files)
        
        if response.status_code == 200:
            print("✅ PDF link removal test passed")
            print(f"   Response size: {len(response.content)} bytes")
        else:
            print(f"❌ PDF link removal test failed: {response.status_code}")
            print(f"   Response: {response.text}")
    except Exception as e:
        print(f"❌ PDF link removal test error: {e}")
    finally:
        # Cleanup
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)

def test_invalid_file():
    """Test handling of invalid files"""
    print("\n🔍 Testing invalid file handling...")
    
    try:
        # Create a text file instead of PDF
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.txt')
        temp_file.write(b"This is not a PDF file")
        temp_file.close()
        
        with open(temp_file.name, 'rb') as f:
            files = {'file': ('test.txt', f, 'text/plain')}
            data = {'password': 'test123'}
            response = requests.post(f"{BASE_URL}/unlock-pdf", files=files, data=data)
        
        if response.status_code == 400:
            print("✅ Invalid file handling test passed")
        else:
            print(f"❌ Invalid file handling test failed: {response.status_code}")
            print(f"   Response: {response.text}")
    except Exception as e:
        print(f"❌ Invalid file handling test error: {e}")
    finally:
        # Cleanup
        if os.path.exists(temp_file.name):
            os.unlink(temp_file.name)

def test_missing_file():
    """Test handling of missing files"""
    print("\n🔍 Testing missing file handling...")
    
    try:
        response = requests.post(f"{BASE_URL}/unlock-pdf")
        
        if response.status_code == 422:  # FastAPI validation error
            print("✅ Missing file handling test passed")
        else:
            print(f"❌ Missing file handling test failed: {response.status_code}")
            print(f"   Response: {response.text}")
    except Exception as e:
        print(f"❌ Missing file handling test error: {e}")

def main():
    """Run all tests"""
    print("🚀 Starting QuickSideTool PDF Security API Tests")
    print("=" * 50)
    
    test_health_check()
    test_root_endpoint()
    test_pdf_unlock()
    test_pdf_lock()
    test_remove_pdf_links()
    test_invalid_file()
    test_missing_file()
    
    print("\n" + "=" * 50)
    print("✅ All tests completed!")

if __name__ == "__main__":
    main() 