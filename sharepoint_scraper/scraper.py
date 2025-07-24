"""
SharePoint document scraper using Microsoft Graph API.
"""

import os
import logging
import requests
import time
from pathlib import Path
from typing import List, Dict, Optional
from urllib.parse import quote

from .auth import SharePointAuth
from .bedrock_integration import BedrockIntegration
from .exceptions import SharePointError, APIError, DownloadError, ConfigurationError

logger = logging.getLogger(__name__)


class SharePointScraper:
    """SharePoint document scraper with Microsoft Graph API integration"""
    
    def __init__(self, site_url: str, auth: SharePointAuth = None, bedrock: BedrockIntegration = None):
        """
        Initialize SharePoint scraper
        
        Args:
            site_url: SharePoint site URL
            auth: SharePointAuth instance (optional, will create one if not provided)
            bedrock: BedrockIntegration instance (optional, for knowledge base ingestion)
        """
        self.site_url = site_url.rstrip('/')
        self.auth = auth or SharePointAuth()
        self.bedrock = bedrock
        self.site_id = None
        
        # Validate site URL
        if not self.site_url or 'sharepoint.com' not in self.site_url:
            raise ConfigurationError("Invalid SharePoint site URL")
        
        logger.info(f"Initialized SharePoint scraper for: {self.site_url}")
        
        # Automatically set up site connection if auth is already authenticated
        if self.auth.is_authenticated():
            self._setup_site_connection()
    
    def authenticate(self) -> bool:
        """
        Authenticate with SharePoint
        
        Returns:
            bool: True if authentication successful
        """
        try:
            self.auth.authenticate()
            return self._setup_site_connection()
        except Exception as e:
            logger.error(f"Authentication failed: {str(e)}")
            return False
    
    def _setup_site_connection(self) -> bool:
        """Setup site connection and get site ID"""
        try:
            headers = self.auth.get_auth_headers()
            
            # Extract site information from URL
            if 'sharepoint.com' in self.site_url:
                # Parse URL like https://contoso.sharepoint.com/sites/sitename
                url_parts = self.site_url.replace('https://', '').split('/')
                hostname = url_parts[0]
                site_path = '/'.join(url_parts[1:]) if len(url_parts) > 1 else ''
                
                # Get site ID using Graph API
                if site_path:
                    site_url_encoded = f"{hostname}:/{site_path}"
                else:
                    site_url_encoded = hostname
                    
                api_url = f"https://graph.microsoft.com/v1.0/sites/{site_url_encoded}"
            else:
                raise ConfigurationError("Unsupported SharePoint URL format")
            
            # Test connection with retry logic
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    response = requests.get(api_url, headers=headers, timeout=30)
                    
                    if response.status_code == 200:
                        site_data = response.json()
                        self.site_id = site_data['id']
                        logger.info(f"Connected to SharePoint site: {site_data.get('displayName', 'Unknown')}")
                        return True
                    elif response.status_code == 404:
                        raise SharePointError("SharePoint site not found. Check the URL and permissions.")
                    elif response.status_code == 403:
                        raise SharePointError("Access denied. Check user permissions.")
                    else:
                        raise APIError(f"API error: {response.status_code}", response.status_code, response.json())
                        
                except requests.RequestException as e:
                    if attempt == max_retries - 1:
                        raise SharePointError(f"Connection failed after {max_retries} attempts: {str(e)}")
                    logger.warning(f"Connection attempt {attempt + 1} failed, retrying...")
                    time.sleep(2 ** attempt)
            
            return False
            
        except Exception as e:
            logger.error(f"Site connection setup failed: {str(e)}")
            return False
    
    def test_connection(self) -> bool:
        """
        Test SharePoint connection and permissions
        
        Returns:
            bool: True if connection is working
        """
        try:
            if not self.auth.is_authenticated():
                logger.error("Not authenticated. Call authenticate() first.")
                return False
            
            if not self.site_id:
                logger.error("Site connection not established.")
                return False
            
            logger.info("Testing SharePoint connection...")
            
            headers = self.auth.get_auth_headers()
            
            # Test basic site access
            site_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}"
            response = requests.get(site_url, headers=headers, timeout=30)
            
            if response.status_code != 200:
                logger.error(f"Site access test failed: {response.status_code}")
                return False
            
            # Test document library access
            drives_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives"
            response = requests.get(drives_url, headers=headers, timeout=30)
            
            if response.status_code != 200:
                logger.error(f"Document library access test failed: {response.status_code}")
                return False
            
            drives = response.json().get('value', [])
            logger.info(f"Found {len(drives)} document libraries")
            
            return True
            
        except Exception as e:
            logger.error(f"Connection test failed: {str(e)}")
            return False
    
    def get_documents(self) -> List[Dict]:
        """
        Get all documents from SharePoint site
        
        Returns:
            List[Dict]: List of document metadata
        """
        if not self.auth.is_authenticated() or not self.site_id:
            raise SharePointError("Not connected. Call authenticate() first.")
        
        documents = []
        
        try:
            logger.info("Scanning for documents...")
            headers = self.auth.get_auth_headers()
            
            # Get all drives (document libraries)
            drives_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives"
            response = requests.get(drives_url, headers=headers, timeout=30)
            
            if response.status_code != 200:
                raise APIError(f"Failed to get document libraries: {response.status_code}")
            
            drives = response.json().get('value', [])
            logger.info(f"Found {len(drives)} document libraries")
            
            # Process each document library
            for drive in drives:
                library_name = drive.get('name', 'Unknown')
                drive_id = drive.get('id')
                
                logger.info(f"Scanning library: {library_name}")
                
                try:
                    self._scan_drive_recursive(drive_id, '', library_name, documents, headers)
                except Exception as e:
                    logger.warning(f"Error scanning library {library_name}: {str(e)}")
            
            logger.info(f"Found {len(documents)} total documents")
            return documents
            
        except Exception as e:
            logger.error(f"Error getting documents: {str(e)}")
            raise SharePointError(f"Failed to get documents: {str(e)}")
    
    def _scan_drive_recursive(self, drive_id: str, folder_path: str, library_name: str, 
                             documents: List[Dict], headers: Dict, page_url: str = None):
        """Recursively scan drive items"""
        try:
            # Determine API URL
            if page_url:
                items_url = page_url
            elif folder_path:
                items_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{drive_id}/root:/{folder_path}:/children"
            else:
                items_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{drive_id}/root/children"
            
            response = requests.get(items_url, headers=headers, timeout=30)
            
            if response.status_code != 200:
                logger.warning(f"Failed to scan folder {folder_path}: {response.status_code}")
                return
            
            data = response.json()
            items = data.get('value', [])
            
            for item in items:
                if 'file' in item:  # It's a file
                    # Create document metadata
                    doc_metadata = {
                        'name': item.get('name', 'Unknown'),
                        'safe_name': self._sanitize_filename(item.get('name', 'Unknown')),
                        'id': item.get('id'),
                        'unique_id': item.get('id'),
                        'drive_id': drive_id,
                        'library': library_name,
                        'path': folder_path,
                        'size': item.get('size', 0),
                        'created': item.get('createdDateTime'),
                        'modified': item.get('lastModifiedDateTime'),
                        'download_url': item.get('@microsoft.graph.downloadUrl'),
                        'web_url': item.get('webUrl'),
                        'mime_type': item.get('file', {}).get('mimeType', 'application/octet-stream')
                    }
                    
                    documents.append(doc_metadata)
                    
                elif 'folder' in item:  # It's a folder
                    folder_name = item.get('name', 'Unknown')
                    new_path = f"{folder_path}/{folder_name}" if folder_path else folder_name
                    
                    # Recursively scan subfolder
                    self._scan_drive_recursive(drive_id, new_path, library_name, documents, headers)
            
            # Handle pagination
            next_link = data.get('@odata.nextLink')
            if next_link:
                self._scan_drive_recursive(drive_id, folder_path, library_name, documents, headers, next_link)
                
        except Exception as e:
            logger.error(f"Error scanning drive items: {str(e)}")
    
    def download_document(self, document: Dict, download_path: str = "downloads") -> Optional[str]:
        """
        Download a single document
        
        Args:
            document: Document metadata dictionary
            download_path: Local download directory
            
        Returns:
            Optional[str]: Path to downloaded file or None if failed
        """
        try:
            # Create directory structure
            library_path = Path(download_path) / self._sanitize_filename(document.get('library', 'default'))
            if document.get('path'):
                folder_path = self._sanitize_filename(document['path'])
                library_path = library_path / folder_path
            
            library_path.mkdir(parents=True, exist_ok=True)
            
            # Determine file path
            safe_filename = document.get('safe_name', self._sanitize_filename(document['name']))
            file_path = library_path / safe_filename
            
            logger.info(f"Downloading: {document['name']}")
            
            # Get download URL
            download_url = document.get('download_url')
            if not download_url:
                # Construct Graph API download URL
                drive_id = document.get('drive_id')
                item_id = document.get('id')
                
                if not drive_id or not item_id:
                    raise DownloadError("Missing drive_id or item_id for download")
                
                download_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{drive_id}/items/{item_id}/content"
            
            # Download with retry logic
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    if '@microsoft.graph.downloadUrl' in download_url:
                        # Direct download URL - no auth needed
                        response = requests.get(download_url, timeout=60)
                    else:
                        # Graph API URL - needs auth
                        headers = self.auth.get_auth_headers()
                        headers['Accept'] = 'application/octet-stream'
                        response = requests.get(download_url, headers=headers, timeout=60)
                    
                    if response.status_code == 200:
                        with open(file_path, 'wb') as f:
                            f.write(response.content)
                        
                        logger.info(f"Downloaded successfully: {file_path}")
                        return str(file_path)
                    else:
                        raise DownloadError(f"Download failed: {response.status_code}")
                        
                except requests.RequestException as e:
                    if attempt == max_retries - 1:
                        raise DownloadError(f"Download failed after {max_retries} attempts: {str(e)}")
                    logger.warning(f"Download attempt {attempt + 1} failed, retrying...")
                    time.sleep(2 ** attempt)
            
            return None
            
        except Exception as e:
            logger.error(f"Error downloading {document['name']}: {str(e)}")
            return None
    
    def download_and_ingest_document(self, document: Dict, download_path: str = "downloads") -> bool:
        """
        Download a document and ingest it into Bedrock knowledge base
        
        Args:
            document: Document metadata dictionary
            download_path: Local download directory
            
        Returns:
            bool: True if successful
        """
        if not self.bedrock:
            raise ConfigurationError("Bedrock integration not configured")
        
        try:
            # Download document
            file_path = self.download_document(document, download_path)
            if not file_path:
                return False
            
            # Ingest into Bedrock
            self.bedrock.ingest_document(
                file_path,
                document_id=document.get('id'),
                title=document.get('name')
            )
            
            logger.info(f"Successfully ingested document: {document['name']}")
            return True
            
        except Exception as e:
            logger.error(f"Error downloading and ingesting {document['name']}: {str(e)}")
            return False
    
    def bulk_download(self, documents: List[Dict], download_path: str = "downloads", 
                     progress_callback=None) -> Dict[str, str]:
        """
        Download multiple documents
        
        Args:
            documents: List of document metadata dictionaries
            download_path: Local download directory
            progress_callback: Optional callback for progress updates
            
        Returns:
            Dict[str, str]: Mapping of document names to local file paths
        """
        results = {}
        total = len(documents)
        
        for i, document in enumerate(documents):
            doc_name = document.get('name', 'Unknown')
            
            try:
                file_path = self.download_document(document, download_path)
                if file_path:
                    results[doc_name] = file_path
                else:
                    logger.warning(f"Failed to download: {doc_name}")
            except Exception as e:
                logger.error(f"Error downloading {doc_name}: {str(e)}")
            
            if progress_callback:
                progress_callback(i + 1, total)
        
        return results
    
    def bulk_download_and_ingest(self, documents: List[Dict], download_path: str = "downloads",
                                progress_callback=None) -> Dict[str, bool]:
        """
        Download and ingest multiple documents into Bedrock
        
        Args:
            documents: List of document metadata dictionaries
            download_path: Local download directory
            progress_callback: Optional callback for progress updates
            
        Returns:
            Dict[str, bool]: Mapping of document names to success status
        """
        if not self.bedrock:
            raise ConfigurationError("Bedrock integration not configured")
        
        results = {}
        total = len(documents)
        
        for i, document in enumerate(documents):
            doc_name = document.get('name', 'Unknown')
            
            try:
                success = self.download_and_ingest_document(document, download_path)
                results[doc_name] = success
            except Exception as e:
                logger.error(f"Error processing {doc_name}: {str(e)}")
                results[doc_name] = False
            
            if progress_callback:
                progress_callback(i + 1, total)
        
        return results
    
    def _sanitize_filename(self, filename: str) -> str:
        """Sanitize filename for safe file system storage"""
        if not filename:
            return 'unknown_file'
        
        # Remove invalid characters
        invalid_chars = '<>:"/\\|?*'
        sanitized = filename
        
        for char in invalid_chars:
            sanitized = sanitized.replace(char, '_')
        
        # Remove leading/trailing whitespace and dots
        sanitized = sanitized.strip('. ')
        
        # Ensure filename is not empty and not too long
        if not sanitized:
            sanitized = 'unknown_file'
        elif len(sanitized) > 200:
            name, ext = os.path.splitext(sanitized)
            sanitized = name[:200-len(ext)] + ext
        
        return sanitized
    
    def get_site_info(self) -> Dict:
        """
        Get SharePoint site information
        
        Returns:
            Dict: Site information
        """
        if not self.auth.is_authenticated() or not self.site_id:
            raise SharePointError("Not connected. Call authenticate() first.")
        
        try:
            headers = self.auth.get_auth_headers()
            site_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}"
            
            response = requests.get(site_url, headers=headers, timeout=30)
            
            if response.status_code == 200:
                return response.json()
            else:
                raise APIError(f"Failed to get site info: {response.status_code}")
                
        except Exception as e:
            logger.error(f"Error getting site info: {str(e)}")
            raise SharePointError(f"Failed to get site info: {str(e)}")
