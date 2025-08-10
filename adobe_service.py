"""
Adobe API Integration Service
Handles all Adobe PDF Services and Document Services API interactions
"""

import os
import asyncio
import aiohttp
import json
import logging
from typing import Optional, Dict, Any, List
from datetime import datetime, timedelta
import tempfile
import shutil

logger = logging.getLogger(__name__)

class AdobeService:
    """Adobe API Service for PDF and Document processing"""
    
    def __init__(self):
        self.client_id = os.getenv('ADOBE_CLIENT_ID')
        self.client_secret = os.getenv('ADOBE_CLIENT_SECRET')
        self.organization_id = os.getenv('ADOBE_ORGANIZATION_ID')
        self.account_id = os.getenv('ADOBE_ACCOUNT_ID')
        self.private_key_path = os.getenv('ADOBE_PRIVATE_KEY_PATH')
        
        # API endpoints
        self.auth_url = "https://ims-na1.adobelogin.com/ims/token/v3"
        self.pdf_services_url = "https://pdf-services.adobe.io"
        self.document_services_url = "https://document-services.adobe.io"
        
        # Token management
        self.access_token = None
        self.token_expires_at = None
        
        # Validate configuration
        self._validate_config()
    
    def _validate_config(self):
        """Validate Adobe API configuration"""
        required_vars = [
            'ADOBE_CLIENT_ID',
            'ADOBE_CLIENT_SECRET', 
            'ADOBE_ORGANIZATION_ID',
            'ADOBE_ACCOUNT_ID'
        ]
        
        missing_vars = [var for var in required_vars if not os.getenv(var)]
        
        if missing_vars:
            logger.warning(f"Missing Adobe API configuration: {missing_vars}")
            logger.warning("Adobe features will be disabled")
            self.enabled = False
        else:
            self.enabled = True
            logger.info("Adobe API service initialized successfully")
    
    async def _get_access_token(self) -> Optional[str]:
        """Get Adobe access token with JWT authentication"""
        if not self.enabled:
            return None
            
        # Check if we have a valid token
        if self.access_token and self.token_expires_at and datetime.now() < self.token_expires_at:
            return self.access_token
        
        try:
            # Create JWT token
            jwt_token = self._create_jwt_token()
            
            # Exchange JWT for access token
            async with aiohttp.ClientSession() as session:
                data = {
                    'grant_type': 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                    'client_id': self.client_id,
                    'client_secret': self.client_secret,
                    'assertion': jwt_token
                }
                
                async with session.post(self.auth_url, data=data) as response:
                    if response.status == 200:
                        token_data = await response.json()
                        self.access_token = token_data['access_token']
                        self.token_expires_at = datetime.now() + timedelta(seconds=token_data['expires_in'] - 300)  # 5 min buffer
                        logger.info("Adobe access token obtained successfully")
                        return self.access_token
                    else:
                        logger.error(f"Failed to get Adobe access token: {response.status}")
                        return None
                        
        except Exception as e:
            logger.error(f"Error getting Adobe access token: {e}")
            return None
    
    def _create_jwt_token(self) -> str:
        """Create JWT token for Adobe API authentication"""
        import jwt
        from datetime import datetime, timedelta
        
        # JWT payload
        payload = {
            'iss': self.organization_id,
            'sub': self.account_id,
            'aud': self.auth_url,
            'exp': datetime.utcnow() + timedelta(hours=1),
            'iat': datetime.utcnow()
        }
        
        # Read private key
        if self.private_key_path and os.path.exists(self.private_key_path):
            with open(self.private_key_path, 'r') as f:
                private_key = f.read()
        else:
            # Use environment variable for private key
            private_key = os.getenv('ADOBE_PRIVATE_KEY')
            if not private_key:
                raise ValueError("Adobe private key not found")
        
        # Create JWT
        token = jwt.encode(payload, private_key, algorithm='RS256')
        return token
    
    async def convert_pdf_to_word(self, pdf_path: str, output_path: str) -> bool:
        """Convert PDF to Word document using Adobe PDF Services"""
        if not self.enabled:
            return False
            
        try:
            access_token = await self._get_access_token()
            if not access_token:
                return False
            
            # Create job for PDF to Word conversion
            job_id = await self._create_conversion_job(pdf_path, 'pdf-to-word', access_token)
            if not job_id:
                return False
            
            # Wait for job completion and download result
            success = await self._wait_for_job_completion(job_id, output_path, access_token)
            return success
            
        except Exception as e:
            logger.error(f"Error converting PDF to Word: {e}")
            return False
    
    async def convert_pdf_to_excel(self, pdf_path: str, output_path: str) -> bool:
        """Convert PDF to Excel using Adobe PDF Services"""
        if not self.enabled:
            return False
            
        try:
            access_token = await self._get_access_token()
            if not access_token:
                return False
            
            # Create job for PDF to Excel conversion
            job_id = await self._create_conversion_job(pdf_path, 'pdf-to-excel', access_token)
            if not job_id:
                return False
            
            # Wait for job completion and download result
            success = await self._wait_for_job_completion(job_id, output_path, access_token)
            return success
            
        except Exception as e:
            logger.error(f"Error converting PDF to Excel: {e}")
            return False
    
    async def compress_pdf(self, pdf_path: str, output_path: str, compression_level: str = 'medium') -> bool:
        """Compress PDF using Adobe PDF Services"""
        if not self.enabled:
            return False
            
        try:
            access_token = await self._get_access_token()
            if not access_token:
                return False
            
            # Create job for PDF compression
            job_id = await self._create_compression_job(pdf_path, compression_level, access_token)
            if not job_id:
                return False
            
            # Wait for job completion and download result
            success = await self._wait_for_job_completion(job_id, output_path, access_token)
            return success
            
        except Exception as e:
            logger.error(f"Error compressing PDF: {e}")
            return False
    
    
    async def _create_conversion_job(self, pdf_path: str, operation: str, access_token: str) -> Optional[str]:
        """Create a conversion job with Adobe PDF Services"""
        try:
            async with aiohttp.ClientSession() as session:
                headers = {
                    'Authorization': f'Bearer {access_token}',
                    'x-api-key': self.client_id
                }
                
                # Upload the PDF file
                with open(pdf_path, 'rb') as f:
                    files = {'input': f}
                    async with session.post(f"{self.pdf_services_url}/assets", headers=headers, data=files) as response:
                        if response.status != 201:
                            logger.error(f"Failed to upload PDF: {response.status}")
                            return None
                        
                        upload_data = await response.json()
                        asset_id = upload_data['assetID']
                
                # Create conversion job
                job_data = {
                    'input': {
                        'href': f"{self.pdf_services_url}/assets/{asset_id}",
                        'type': 'application/pdf'
                    },
                    'targetFormat': operation.split('-')[1].upper()
                }
                
                async with session.post(f"{self.pdf_services_url}/operation/{operation}", 
                                      headers=headers, json=job_data) as response:
                    if response.status != 201:
                        logger.error(f"Failed to create conversion job: {response.status}")
                        return None
                    
                    job_data = await response.json()
                    return job_data['jobID']
                    
        except Exception as e:
            logger.error(f"Error creating conversion job: {e}")
            return None
    
    async def _create_compression_job(self, pdf_path: str, compression_level: str, access_token: str) -> Optional[str]:
        """Create a compression job with Adobe PDF Services"""
        try:
            async with aiohttp.ClientSession() as session:
                headers = {
                    'Authorization': f'Bearer {access_token}',
                    'x-api-key': self.client_id
                }
                
                # Upload the PDF file
                with open(pdf_path, 'rb') as f:
                    files = {'input': f}
                    async with session.post(f"{self.pdf_services_url}/assets", headers=headers, data=files) as response:
                        if response.status != 201:
                            logger.error(f"Failed to upload PDF: {response.status}")
                            return None
                        
                        upload_data = await response.json()
                        asset_id = upload_data['assetID']
                
                # Create compression job
                job_data = {
                    'input': {
                        'href': f"{self.pdf_services_url}/assets/{asset_id}",
                        'type': 'application/pdf'
                    },
                    'targetFormat': 'pdf',
                    'compressionLevel': compression_level
                }
                
                async with session.post(f"{self.pdf_services_url}/operation/compress-pdf", 
                                      headers=headers, json=job_data) as response:
                    if response.status != 201:
                        logger.error(f"Failed to create compression job: {response.status}")
                        return None
                    
                    job_data = await response.json()
                    return job_data['jobID']
                    
        except Exception as e:
            logger.error(f"Error creating compression job: {e}")
            return None
    
    
    async def _wait_for_job_completion(self, job_id: str, output_path: str, access_token: str) -> bool:
        """Wait for job completion and download result"""
        try:
            async with aiohttp.ClientSession() as session:
                headers = {
                    'Authorization': f'Bearer {access_token}',
                    'x-api-key': self.client_id
                }
                
                # Poll for job completion
                max_attempts = 30  # 5 minutes with 10-second intervals
                for attempt in range(max_attempts):
                    async with session.get(f"{self.pdf_services_url}/job/{job_id}", headers=headers) as response:
                        if response.status != 200:
                            logger.error(f"Failed to check job status: {response.status}")
                            return False
                        
                        job_status = await response.json()
                        status = job_status['status']
                        
                        if status == 'done':
                            # Download the result
                            result_asset_id = job_status['result']['assetID']
                            async with session.get(f"{self.pdf_services_url}/assets/{result_asset_id}", 
                                                 headers=headers) as download_response:
                                if download_response.status == 200:
                                    with open(output_path, 'wb') as f:
                                        async for chunk in download_response.content.iter_chunked(8192):
                                            f.write(chunk)
                                    logger.info(f"Job {job_id} completed successfully")
                                    return True
                                else:
                                    logger.error(f"Failed to download result: {download_response.status}")
                                    return False
                        
                        elif status == 'failed':
                            logger.error(f"Job {job_id} failed: {job_status.get('error', 'Unknown error')}")
                            return False
                        
                        # Wait before next poll
                        await asyncio.sleep(10)
                
                logger.error(f"Job {job_id} timed out")
                return False
                
        except Exception as e:
            logger.error(f"Error waiting for job completion: {e}")
            return False
    
    # Removed OCR and optimization helpers since not used by current frontend

# Global Adobe service instance
adobe_service = AdobeService()
