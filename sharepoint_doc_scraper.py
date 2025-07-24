"""
SharePoint Connector and Document Scraper with User-Level Authentication
Uses Microsoft Graph API directly for user-level access
"""

import os
import logging
import json
from pathlib import Path
from typing import Optional, List, Dict
import requests
from dotenv import load_dotenv
import time
import urllib3
import webbrowser
import urllib.parse
import secrets
import hashlib
import base64
from http.server import HTTPServer, BaseHTTPRequestHandler
import threading
from dataclasses import dataclass
from ingest_kb_data_source import ingest_document

# Disable SSL warnings for corporate environments
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('sharepoint_scraper.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class AuthCallbackHandler(BaseHTTPRequestHandler):
    """HTTP handler for OAuth callback"""
    
    def do_GET(self):
        # Parse the callback URL for authorization code
        if '?' in self.path:
            query_params = urllib.parse.parse_qs(self.path.split('?')[1])
            if 'code' in query_params:
                self.server.auth_code = query_params['code'][0]
                self.send_response(200)
                self.send_header('Content-type', 'text/html')
                self.end_headers()
                self.wfile.write(b'''
                    <html>
                        <body>
                            <h2>Authentication Successful!</h2>
                            <p>You can now close this window and return to your application.</p>
                            <script>window.close();</script>
                        </body>
                    </html>
                ''')
            elif 'error' in query_params:
                self.server.auth_error = query_params.get('error_description', ['Unknown error'])[0]
                self.send_response(400)
                self.send_header('Content-type', 'text/html')
                self.end_headers()
                self.wfile.write(f'''
                    <html>
                        <body>
                            <h2>Authentication Failed</h2>
                            <p>Error: {self.server.auth_error}</p>
                        </body>
                    </html>
                '''.encode())
        else:
            self.send_response(404)
            self.end_headers()
    
    def log_message(self, format, *args):
        # Suppress default logging
        pass

@dataclass
class Token:
    tokenType: str
    accessToken: str

class UserLevelSharePointScraper:
    """SharePoint document scraper using user-level authentication with MFA support"""
    
    def __init__(self, site_url: str, client_id: str = None, tenant_id: str = None):
        """
        Initialize the SharePoint scraper with OAuth credentials
        
        Args:
            site_url: SharePoint site URL
            client_id: Azure AD App Registration Client ID
            tenant_id: Azure AD Tenant ID (optional, can be extracted from site URL)
        """
        self.site_url = site_url.rstrip('/')
        self.client_id = client_id or os.getenv('AZURE_CLIENT_ID')
        self.tenant_id = tenant_id or os.getenv('AZURE_TENANT_ID')
        self.redirect_uri = "http://localhost:8080/callback"

        self.access_token = None
        self.token_response = None
        self.site_id = None
        
        # Extract tenant from SharePoint URL if not provided
        if not self.tenant_id and 'sharepoint.com' in self.site_url:
            # Extract tenant from URL like https://contoso.sharepoint.com
            domain_parts = self.site_url.split('://')[1].split('.')[0]
            self.tenant_id = domain_parts.split('.')[0] if '.' in domain_parts else domain_parts
        
        logger.info(f"🔧 Initialized SharePoint scraper for: {self.site_url}")
        if self.client_id:
            logger.info(f"🔧 Using Azure AD App: {self.client_id}")
    
    def authenticate_user(self) -> bool:
        """
        Authenticate user using interactive OAuth 2.0 with PKCE flow (supports MFA)
        
        Returns:
            bool: True if authentication successful, False otherwise
        """
        try:
            if not self.client_id:
                logger.error("❌ Azure AD Client ID required for MFA authentication")
                logger.info("💡 Please set AZURE_CLIENT_ID in your .env file or pass it to the constructor")
                logger.info("💡 You need to register an app in Azure AD with appropriate SharePoint permissions")
                return False
            
            logger.info("🔐 Starting interactive OAuth authentication (supports MFA)...")
            
            # Generate PKCE parameters
            code_verifier = base64.urlsafe_b64encode(secrets.token_bytes(32)).decode('utf-8').rstrip('=')
            code_challenge = base64.urlsafe_b64encode(
                hashlib.sha256(code_verifier.encode('utf-8')).digest()
            ).decode('utf-8').rstrip('=')
            
            # Determine authorization endpoint
            if self.tenant_id:
                auth_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/authorize"
            else:
                auth_url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
            
            # Build authorization URL with SharePoint-specific scopes
            auth_params = {
                'client_id': self.client_id,
                'response_type': 'code',
                'redirect_uri': self.redirect_uri,
                'scope': 'https://graph.microsoft.com/Sites.Read.All https://graph.microsoft.com/Files.Read.All offline_access',
                'code_challenge': code_challenge,
                'code_challenge_method': 'S256',
                'prompt': 'select_account'
            }
            
            auth_url_full = f"{auth_url}?{urllib.parse.urlencode(auth_params)}"
            
            # Start local server to handle callback
            server = HTTPServer(('localhost', 8080), AuthCallbackHandler)
            server.auth_code = None
            server.auth_error = None
            
            # Start server in background thread
            server_thread = threading.Thread(target=server.serve_forever)
            server_thread.daemon = True
            server_thread.start()
            
            try:
                # Open browser for authentication
                logger.info("🌐 Opening browser for authentication...")
                logger.info("💡 Please complete the sign-in process in your browser (MFA supported)")
                webbrowser.open(auth_url_full)
                
                # Wait for callback with timeout
                timeout = 300  # 5 minutes
                start_time = time.time()
                
                while server.auth_code is None and server.auth_error is None:
                    if time.time() - start_time > timeout:
                        logger.error("❌ Authentication timeout - no response received")
                        return False
                    time.sleep(1)
                
                if server.auth_error:
                    logger.error(f"❌ Authentication failed: {server.auth_error}")
                    return False
                
                if not server.auth_code:
                    logger.error("❌ No authorization code received")
                    return False
                
                logger.info("✅ Authorization code received, exchanging for access token...")
                
                # Exchange authorization code for access token
                return self._exchange_code_for_token(server.auth_code, code_verifier)
                
            finally:
                server.shutdown()
                server.server_close()
                
        except Exception as e:
            logger.error(f"❌ Interactive authentication failed: {str(e)}")
            return False
    
    def _exchange_code_for_token(self, auth_code: str, code_verifier: str) -> bool:
        """Exchange authorization code for access token"""
        try:
            # Determine token endpoint
            if self.tenant_id:
                token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
            else:
                token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
            
            # Token request parameters
            token_data = {
                'client_secret': os.getenv('AZURE_CLIENT_SECRET'),
                'client_id': self.client_id,
                'grant_type': 'authorization_code',
                'code': auth_code,
                'redirect_uri': self.redirect_uri,
                'code_verifier': code_verifier,
                'scope': 'https://graph.microsoft.com/Sites.Read.All https://graph.microsoft.com/Files.Read.All'
            }
            
            # Make token request
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            response = requests.post(token_url, data=token_data, headers=headers)
            
            if response.status_code == 200:
                self.token_response = response.json()
                self.access_token = self.token_response['access_token']
                
                # Use Microsoft Graph API to access SharePoint
                return self._setup_sharepoint_context_with_graph()
                
            else:
                logger.error(f"❌ Token exchange failed: {response.status_code} - {response.text}")
                
                # Provide specific guidance based on error
                error_info = response.json() if response.headers.get('content-type', '').startswith('application/json') else {}
                error_code = error_info.get('error')
                
                if error_code == 'invalid_client':
                    logger.error("💡 Azure AD App configuration issue:")
                    logger.error("   1. Go to Azure Portal > App Registrations > Your App")
                    logger.error("   2. Go to Authentication section")
                    logger.error("   3. Enable 'Allow public client flows' = YES")
                    logger.error("   4. Add redirect URI: http://localhost:8080/callback")
                    logger.error("   5. Save the configuration")
                elif error_code == 'invalid_scope':
                    logger.error("💡 Scope/Permission issue:")
                    logger.error("   1. Go to API Permissions in your Azure AD App")
                    logger.error("   2. Add Microsoft Graph permissions:")
                    logger.error("      - Sites.Read.All (or Sites.ReadWrite.All)")
                    logger.error("      - Files.Read.All (or Files.ReadWrite.All)")
                    logger.error("   3. Grant admin consent")
                
                return False
                
        except Exception as e:
            logger.error(f"❌ Token exchange error: {str(e)}")
            return False
    
    def _wrap_token(self, token_response: dict) -> Token:
        """Wrap token response in the format expected by Office365-REST-Python-Client"""
        if "access_token" in token_response:
            return Token(tokenType="Bearer", accessToken=token_response["access_token"])
        else:
            raise Exception(
                f"Authentication error: {token_response.get('error')}, {token_response.get('error_description')}"
            )
    
    def _setup_sharepoint_context_with_graph(self) -> bool:
        """Setup SharePoint context using Microsoft Graph API token"""
        try:
            return self._setup_graph_api_access()
        except Exception as e:
            logger.error(f"❌ Graph API setup failed: {str(e)}")
            return False
    
    def _setup_graph_api_access(self) -> bool:
        """Setup using Microsoft Graph API for SharePoint access"""
        try:
            # Test Graph API access to SharePoint
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            }
            
            # Extract site information from URL
            if 'sharepoint.com' in self.site_url:
                # Parse SharePoint URL to get site path
                url_parts = self.site_url.split('/')
                hostname = url_parts[2]  # tenant.sharepoint.com
                site_path = '/'.join(url_parts[3:]) if len(url_parts) > 3 else ''
                
                # Get site using Graph API
                if site_path:
                    graph_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}"
                else:
                    graph_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}"
                
                response = requests.get(graph_url, headers=headers)
                
                if response.status_code == 200:
                    site_data = response.json()
                    logger.info("✅ Graph API authentication successful!")
                    logger.info(f"📍 Connected to site: {site_data.get('displayName', 'Unknown')}")
                    
                    # Store site ID for future Graph API calls
                    self.site_id = site_data.get('id')
                    return True
                else:
                    logger.error(f"❌ Graph API site access failed: {response.status_code}")
                    return False
            else:
                logger.error("❌ Invalid SharePoint URL format")
                return False
                
        except Exception as e:
            logger.error(f"❌ Graph API setup error: {str(e)}")
            return False
    
    def test_connection(self) -> bool:
        """
        Test the SharePoint connection and user permissions using Microsoft Graph API
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        try:
            if not self.access_token:
                logger.error("❌ Not authenticated - call authenticate_user() first")
                return False
            
            logger.info("🔍 Testing SharePoint connection via Microsoft Graph API...")
            
            # Setup Graph API headers
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            }
            
            # Test basic site access with retry logic
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    # Get site info using Microsoft Graph
                    if self.site_id:
                        graph_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}"
                    else:
                        # Extract site information from URL
                        url_parts = self.site_url.split('/')
                        hostname = url_parts[2]  # tenant.sharepoint.com
                        site_path = '/'.join(url_parts[3:]) if len(url_parts) > 3 else ''
                        
                        # Get site using hostname and path
                        if site_path:
                            graph_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}"
                        else:
                            graph_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}"
                    
                    response = requests.get(graph_url, headers=headers)
                    
                    if response.status_code == 200:
                        site_data = response.json()
                        logger.info(f"✅ Site access confirmed via Graph API")
                        logger.info(f"📖 Site Title: {site_data.get('displayName', 'Unknown Site')}")
                        
                        # Store site ID if not already stored
                        if not self.site_id:
                            self.site_id = site_data.get('id')
                        
                        break
                    else:
                        raise Exception(f"Graph API returned status: {response.status_code} - {response.text}")
                        
                except Exception as e:
                    if attempt < max_retries - 1:
                        logger.warning(f"⚠️  Connection attempt {attempt + 1} failed, retrying...")
                        time.sleep(1)
                        continue
                    else:
                        raise e
            
            # Test document library access using Graph API
            logger.info("🔍 Testing document library access...")
            try:
                # Get drives (document libraries) using Microsoft Graph
                drives_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives"
                response = requests.get(drives_url, headers=headers)
                
                if response.status_code == 200:
                    drives_data = response.json()
                    doc_libraries = drives_data.get('value', [])
                    
                    logger.info(f"📚 Found {len(doc_libraries)} document libraries")
                    for lib in doc_libraries[:3]:  # Show first 3
                        logger.info(f"  📂 {lib.get('name', 'Unknown')} ({lib.get('driveType', 'Unknown')})")
                else:
                    logger.warning(f"⚠️  Could not access document libraries: {response.status_code} - {response.text}")
                    logger.info("💡 You may have limited permissions")
            
            except Exception as e:
                logger.warning(f"⚠️  Could not access document libraries: {str(e)}")
                logger.info("💡 You may have limited permissions")
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Connection test failed: {str(e)}")
            
            # Enhanced error reporting
            if "401" in str(e):
                logger.error("💡 Authentication may have expired or credentials are invalid")
            elif "403" in str(e):
                logger.error("💡 Access denied - check site permissions")
            elif "404" in str(e):
                logger.error("💡 Site not found - check the SharePoint URL")
            elif "timeout" in str(e).lower():
                logger.error("💡 Network timeout - check connectivity")
            
            return False
    
    def get_documents_graph(self) -> List[Dict]:
        """
        Get all documents from SharePoint site using Microsoft Graph API
        
        Returns:
            List[Dict]: List of document metadata dictionaries
        """
        documents = []
        
        try:
            if not self.access_token or not self.site_id:
                logger.error("❌ Not authenticated or site ID not available")
                return documents
            
            logger.info("📄 Scanning for documents using Microsoft Graph API...")
            
            # Setup Graph API headers
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json'
            }
            
            # Get all drives (document libraries) from the site
            try:
                drives_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives"
                response = requests.get(drives_url, headers=headers)
                
                if response.status_code != 200:
                    logger.error(f"❌ Failed to get site drives: {response.status_code} - {response.text}")
                    return documents
                    
                drives = response.json().get('value', [])
                logger.info(f"📚 Found {len(drives)} document libraries")
                
            except Exception as e:
                logger.error(f"❌ Failed to get site drives: {str(e)}")
                return documents
            
            # Process each drive (document library)
            for drive in drives:
                try:
                    drive_id = drive.get('id')
                    drive_name = drive.get('name', 'Unknown Library')
                    logger.info(f"📂 Scanning library: {drive_name}")
                    
                    # Get root items first
                    self._scan_drive_items_recursive(drive_id, "", drive_name, documents, headers)
                    
                    # Rate limiting to avoid throttling
                    time.sleep(1.0)
                    
                except Exception as e:
                    logger.warning(f"⚠️  Failed to scan library {drive_name}: {str(e)}")
                    continue
            
            logger.info(f"📄 Found {len(documents)} total documents")
            return documents
            
        except Exception as e:
            logger.error(f"❌ Error getting documents: {str(e)}")
            return documents

    def _scan_drive_items_recursive(self, drive_id: str, folder_path: str, library_name: str, 
                                documents: List[Dict], headers: Dict, page_url: str = None):
        """
        Recursively scan drive items using Microsoft Graph API
        
        Args:
            drive_id: SharePoint drive ID
            folder_path: Current folder path in the API
            library_name: Name of the document library
            documents: List to append document metadata to
            headers: Graph API request headers
            page_url: URL for pagination (optional)
        """
        try:
            # Determine URL based on folder path and pagination
            if page_url:
                items_url = page_url
            elif folder_path:
                items_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_path}/children"
            else:
                items_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
            
            # Get items with error handling
            response = requests.get(items_url, headers=headers)
            
            if response.status_code != 200:
                logger.warning(f"⚠️  Error getting items: {response.status_code} - {response.text}")
                return
                
            data = response.json()
            items = data.get('value', [])
            
            # Process each item
            for item in items:
                try:
                    item_name = item.get('name', 'Unknown')
                    item_id = item.get('id', '')
                    
                    # Check if item is a folder
                    if 'folder' in item:
                        # Skip system folders
                        if item_name.startswith('_') or item_name in ['Forms']:
                            continue
                        
                        # Recursively scan subfolders
                        self._scan_drive_items_recursive(drive_id, item_id, library_name, documents, headers)
                        
                        # Rate limiting
                        time.sleep(0.5)
                        
                    # It's a file
                    elif 'file' in item:
                        file_path = item.get('parentReference', {}).get('path', '').split('root:/')[-1]
                        if file_path.startswith('/'):
                            file_path = file_path[1:]
                        full_path = f"{file_path}/{item_name}" if file_path else item_name
                        
                        # Create safe filename
                        safe_name = self._sanitize_filename(item_name)
                        
                        # Extract file metadata
                        doc_info = {
                            'name': item_name,
                            'safe_name': safe_name,
                            'library': library_name,
                            'path': full_path,
                            'size': item.get('size', 0),
                            'created': item.get('createdDateTime'),
                            'modified': item.get('lastModifiedDateTime'),
                            'author': item.get('createdBy', {}).get('user', {}).get('displayName', 'Unknown'),
                            'server_relative_url': item.get('webUrl', ''),
                            'file_type': os.path.splitext(item_name)[1].lower(),
                            'unique_id': item_id,
                            'web_url': item.get('webUrl', ''),
                            'download_url': item.get('@microsoft.graph.downloadUrl', ''),
                            'drive_id': drive_id
                        }
                        
                        documents.append(doc_info)
                
                except Exception as e:
                    logger.warning(f"⚠️  Error processing item {item.get('name', 'Unknown')}: {str(e)}")
                    continue
            
            # Handle pagination if more items exist
            next_link = data.get('@odata.nextLink')
            if next_link:
                # Rate limiting before fetching next page
                time.sleep(0.5)
                self._scan_drive_items_recursive(drive_id, folder_path, library_name, documents, headers, next_link)
                
        except Exception as e:
            logger.error(f"❌ Error scanning drive items: {str(e)}")

    def _sanitize_filename(self, filename: str) -> str:
        """
        Sanitize filename for safe file system storage
        
        Args:
            filename: Original filename
            
        Returns:
            str: Sanitized filename safe for file system
        """
        if not filename:
            return 'unknown_file'
            
        # Remove or replace invalid characters
        invalid_chars = '<>:"/\\|?*'
        sanitized = filename
        
        for char in invalid_chars:
            sanitized = sanitized.replace(char, '_')
        
        # Remove leading/trailing whitespace and dots
        sanitized = sanitized.strip('. ')
        
        # Ensure filename is not empty and not too long
        if not sanitized:
            sanitized = 'unknown_file'
        elif len(sanitized) > 200:  # Limit filename length
            name, ext = os.path.splitext(sanitized)
            sanitized = name[:200-len(ext)] + ext
        
        return sanitized


    def download_and_ingest_document(self, document: Dict, download_path: str = "downloads") -> bool:
        """
        Download a document using Microsoft Graph API with enhanced error handling and ingest it into Bedrock knowledge base using custom data source
        
        Args:
            document: Document metadata dictionary containing Graph API identifiers
            download_path: Local path to save the document
            
        Returns:
            bool: True if download successful, False otherwise
        """
        try:
            if not self.access_token:
                logger.error("❌ Not authenticated - missing access token")
                return False
            
            # Extract Graph API identifiers from document metadata
            site_id = self.site_id
            drive_id = document.get('drive_id') 
            item_id = document.get('unique_id')
            
            if not all([site_id, item_id]):
                logger.error(f"❌ Missing required Graph API identifiers for {document['name']}")
                return False
            
            # Create download directory structure
            library_path = Path(download_path) / self._sanitize_filename(document.get('library', 'default'))
            if document.get('path'):
                file_dir = os.path.dirname(document['path'])
                if file_dir:
                    sanitized_dir = '/'.join(self._sanitize_filename(part) for part in file_dir.split('/'))
                    library_path = library_path / sanitized_dir
            
            library_path.mkdir(parents=True, exist_ok=True)
            
            # Use sanitized filename
            safe_filename = document.get('safe_name', self._sanitize_filename(document['name']))
            file_path = library_path / safe_filename
            
            logger.info(f"⬇️  Downloading: {document['name']}")
            
            # Construct Graph API URL for file download
            if drive_id:
                graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}/content"
            else:
                graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/content"
            
            # Set up headers with authorization token
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/octet-stream'
            }
            
            # Download file content using Graph API with retry logic
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    # Send GET request to Graph API endpoint
                    response = requests.get(graph_url, headers=headers, stream=True)
                    
                    if response.status_code == 200:
                        # Save file content to local path
                        with open(file_path, 'wb') as local_file:
                            for chunk in response.iter_content(chunk_size=8192):
                                if chunk:
                                    local_file.write(chunk)
                        
                        logger.info(f"✅ Downloaded: {document['name']} -> {file_path}")
                        
                        # Ingest the document into Bedrock knowledge base
                        ingest_document(file_path, document['name'])
                        logger.info(f"📚 Ingested document into knowledge base: {document['name']}")
                        return True
                    
                    elif response.status_code == 401:
                        logger.error("❌ Authentication failed - token may be expired")
                        return False
                        
                    elif response.status_code == 403:
                        logger.error(f"❌ Access denied to {document['name']} - insufficient permissions")
                        return False
                        
                    elif response.status_code == 404:
                        logger.error(f"❌ File not found: {document['name']}")
                        return False
                        
                    else:
                        response.raise_for_status()
                        
                except requests.exceptions.RequestException as e:
                    if attempt < max_retries - 1:
                        logger.warning(f"⚠️  Download attempt {attempt + 1} failed, retrying...")
                        time.sleep(2)
                        continue
                    else:
                        raise e
            
        except Exception as e:
            logger.error(f"❌ Error downloading {document['name']}: {str(e)}")
            return False

    def get_file_metadata_from_graph(self, site_id: str, item_id: str) -> Dict:
        """
        Get file metadata using Microsoft Graph API
        
        Args:
            site_id: SharePoint site identifier
            item_id: File item identifier
            
        Returns:
            Dict: File metadata from Graph API
        """
        try:
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json'
            }
            
            metadata_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
            response = requests.get(metadata_url, headers=headers)
            
            if response.status_code == 200:
                return response.json()
            else:
                logger.error(f"Failed to get metadata: {response.status_code}")
                return {}
                
        except Exception as e:
            logger.error(f"Error getting file metadata: {str(e)}")
            return {}

    
    def get_user_permissions(self) -> Dict:
        """
        Get current user's permissions on the site
        
        Returns:
            Dict: User permission information
        """
        try:
            if self.access_token:
                try:
                    headers = {
                        'Authorization': f'Bearer {self.access_token}',
                        'Accept': 'application/json'
                    }
                    
                    # Try to get current user info from Graph API
                    graph_url = "https://graph.microsoft.com/v1.0/me"
                    response = requests.get(graph_url, headers=headers)
                    
                    if response.status_code == 200:
                        user_data = response.json()
                        user_info = {
                            'login_name': user_data.get('userPrincipalName', 'Unknown'),
                            'title': user_data.get('displayName', 'Unknown'),
                            'email': user_data.get('mail', user_data.get('userPrincipalName', 'Unknown')),
                            'user_id': user_data.get('id', 'Unknown')
                        }
                        
                        logger.info(f"👤 Current user: {user_info['title']} ({user_info['email']})")
                        return user_info
                except Exception as e:
                    logger.warning(f"⚠️ Failed to get user info from Graph API: {str(e)}")
            
        except Exception as e:
            logger.error(f"❌ Error getting user permissions: {str(e)}")
            return {}


def main():
    """Main function to run the user-level SharePoint scraper"""
    
    # Load configuration
    site_url = os.getenv('SHAREPOINT_SITE_URL')
    if not site_url:
        print("❌ Please set SHAREPOINT_SITE_URL in your .env file")
        return
    
    print("SharePoint Document Scraper (User-Level Authentication with MFA)")
    print("=" * 60)
    print(f"🔗 Connecting to: {site_url}")
    
    try:
        # Initialize scraper
        scraper = UserLevelSharePointScraper(site_url)
        
        # Authenticate with user credentials (MFA supported)
        print("\n🔐 Starting user authentication with MFA support...")
        if not scraper.authenticate_user():
            print("❌ User authentication failed")
            return
        
        # Get user info
        user_info = scraper.get_user_permissions()
        
        # Test connection
        print("\n🔍 Testing connection...")
        if not scraper.test_connection():
            print("❌ Connection test failed")
            return
        
        # Get documents
        print("\n📄 Scanning for documents...")
        documents = scraper.get_documents_graph()
        
        if documents:
            print(f"\n✅ Found {len(documents)} documents:")
            
            # Show summary by library
            libraries = {}
            total_size = 0
            
            for doc in documents:
                lib_name = doc['library']
                if lib_name not in libraries:
                    libraries[lib_name] = {'count': 0, 'size': 0}
                libraries[lib_name]['count'] += 1
                libraries[lib_name]['size'] += doc['size']
                total_size += doc['size']
            
            print(f"\n📊 Summary:")
            print(f"Total documents: {len(documents)}")
            print(f"Total size: {total_size / (1024 * 1024):.2f} MB")
            
            for lib_name, stats in libraries.items():
                print(f"  📂 {lib_name}: {stats['count']} documents ({stats['size'] / (1024 * 1024):.2f} MB)")
            
            # Show first few documents
            print(f"\n📄 Sample documents:")
            for doc in documents[:5]:
                size_mb = doc['size'] / (1024 * 1024) if doc['size'] > 0 else 0
                print(f"  📄 {doc['library']}/{doc['path']} ({size_mb:.2f} MB)")
            
            if len(documents) > 5:
                print(f"  ... and {len(documents) - 5} more documents")
            
            # Save results
            output_file = "sharepoint_documents.json"
            with open(output_file, 'w') as f:
                json.dump(documents, f, indent=2, default=str)
            print(f"\n💾 Document list saved to: {output_file}")
            
            # Ask about downloading
            download_choice = input("\nWould you like to download documents? (y/n): ").strip().lower()
            if download_choice in ['y', 'yes']:
                download_path = input("Enter download directory (default: downloads): ").strip() or "downloads"
                
                print(f"\n⬇️  Downloading {len(documents)} documents...")
                successful = 0
                failed = 0
                
                for i, doc in enumerate(documents, 1):
                    print(f"Progress: {i}/{len(documents)} - {doc['name']}", end='\r')
                    
                    if scraper.download_and_ingest_document(doc, download_path):
                        successful += 1
                    else:
                        failed += 1
                    
                    # Rate limiting to avoid SharePoint throttling
                    time.sleep(1.0)
                
                print(f"\n✅ Downloaded {successful}/{len(documents)} documents to '{download_path}/'")
                if failed > 0:
                    print(f"⚠️  {failed} downloads failed (check logs for details)")
                
        else:
            print("❌ No documents found or accessible")
            
    except KeyboardInterrupt:
        print("\n⏹️  Operation cancelled by user")
    except Exception as e:
        logger.error(f"❌ Unexpected error: {str(e)}")
        print(f"❌ An error occurred: {str(e)}")


if __name__ == "__main__":
    load_dotenv()
    main()