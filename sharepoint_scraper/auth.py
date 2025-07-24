"""
SharePoint authentication module using Microsoft Graph API with OAuth 2.0 and PKCE flow.
Supports MFA and interactive authentication.
"""

import os
import secrets
import hashlib
import base64
import threading
import webbrowser
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
from typing import Optional, Dict
import requests
import logging

from .exceptions import AuthenticationError, ConfigurationError

logger = logging.getLogger(__name__)


class AuthCallbackHandler(BaseHTTPRequestHandler):
    """HTTP handler for OAuth callback"""
    
    def do_GET(self):
        """Handle GET request for OAuth callback"""
        if '?' in self.path:
            query_params = urllib.parse.parse_qs(self.path.split('?')[1])
            if 'code' in query_params:
                self.server.auth_code = query_params['code'][0]
                self._send_success_response()
            elif 'error' in query_params:
                self.server.auth_error = query_params.get('error_description', ['Unknown error'])[0]
                self._send_error_response()
        else:
            self.send_response(404)
            self.end_headers()
    
    def _send_success_response(self):
        """Send success response to browser"""
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
    
    def _send_error_response(self):
        """Send error response to browser"""
        self.send_response(400)
        self.send_header('Content-type', 'text/html')
        self.end_headers()
        self.wfile.write(b'''
            <html>
                <body>
                    <h2>Authentication Failed!</h2>
                    <p>There was an error during authentication. Please try again.</p>
                </body>
            </html>
        ''')
    
    def log_message(self, format, *args):
        """Suppress default logging"""
        pass


class SharePointAuth:
    """Handles SharePoint authentication using Microsoft Graph API with OAuth 2.0"""
    
    def __init__(self, client_id: str = None, tenant_id: str = None, redirect_uri: str = "http://localhost:8080/callback"):
        """
        Initialize SharePoint authentication
        
        Args:
            client_id: Azure AD App Registration Client ID
            tenant_id: Azure AD Tenant ID  
            redirect_uri: OAuth redirect URI (default: http://localhost:8080/callback)
        """
        self.client_id = client_id or os.getenv('AZURE_CLIENT_ID')
        self.tenant_id = tenant_id or os.getenv('AZURE_TENANT_ID')
        self.redirect_uri = redirect_uri
        self.access_token = None
        self.token_response = None
        
        if not self.client_id:
            raise ConfigurationError("Azure Client ID is required. Set AZURE_CLIENT_ID environment variable or pass client_id parameter.")
    
    def authenticate(self) -> str:
        """
        Authenticate user using interactive OAuth 2.0 with PKCE flow (supports MFA)
        
        Returns:
            str: Access token
            
        Raises:
            AuthenticationError: If authentication fails
        """
        try:
            logger.info("Starting interactive OAuth authentication (supports MFA)...")
            
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
                logger.info(f"Opening browser for authentication: {auth_url_full}")
                webbrowser.open(auth_url_full)
                
                # Wait for callback
                timeout = 300  # 5 minutes
                waited = 0
                while server.auth_code is None and server.auth_error is None and waited < timeout:
                    threading.Event().wait(1)
                    waited += 1
                
                if server.auth_error:
                    raise AuthenticationError(f"OAuth error: {server.auth_error}")
                
                if server.auth_code is None:
                    raise AuthenticationError("Authentication timeout - no response received")
                
                # Exchange code for token
                if not self._exchange_code_for_token(server.auth_code, code_verifier):
                    raise AuthenticationError("Failed to exchange authorization code for token")
                
                logger.info("Authentication successful!")
                return self.access_token
                
            finally:
                server.shutdown()
                server_thread.join(timeout=5)
                
        except Exception as e:
            if isinstance(e, AuthenticationError):
                raise
            logger.error(f"Interactive authentication failed: {str(e)}")
            raise AuthenticationError(f"Authentication failed: {str(e)}")
    
    def _exchange_code_for_token(self, auth_code: str, code_verifier: str) -> bool:
        """
        Exchange authorization code for access token
        
        Args:
            auth_code: Authorization code from OAuth callback
            code_verifier: PKCE code verifier
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Determine token endpoint
            if self.tenant_id:
                token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
            else:
                token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
            
            # Token request parameters
            token_data = {
                'client_id': self.client_id,
                'grant_type': 'authorization_code',
                'code': auth_code,
                'redirect_uri': self.redirect_uri,
                'code_verifier': code_verifier,
                'scope': 'https://graph.microsoft.com/Sites.Read.All https://graph.microsoft.com/Files.Read.All'
            }
            
            # Add client secret if available
            client_secret = os.getenv('AZURE_CLIENT_SECRET')
            if client_secret:
                token_data['client_secret'] = client_secret
            
            # Make token request
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            response = requests.post(token_url, data=token_data, headers=headers)
            
            if response.status_code == 200:
                self.token_response = response.json()
                self.access_token = self.token_response['access_token']
                logger.info("Successfully obtained access token")
                return True
            else:
                error_data = response.json() if response.headers.get('content-type', '').startswith('application/json') else {}
                logger.error(f"Token exchange failed: {response.status_code} - {error_data}")
                return False
                
        except Exception as e:
            logger.error(f"Token exchange error: {str(e)}")
            return False
    
    def get_access_token(self) -> Optional[str]:
        """
        Get current access token
        
        Returns:
            Optional[str]: Current access token or None if not authenticated
        """
        return self.access_token
    
    def is_authenticated(self) -> bool:
        """
        Check if user is currently authenticated
        
        Returns:
            bool: True if authenticated, False otherwise
        """
        return self.access_token is not None
    
    def get_auth_headers(self) -> Dict[str, str]:
        """
        Get headers for authenticated requests
        
        Returns:
            Dict[str, str]: Headers with authorization token
            
        Raises:
            AuthenticationError: If not authenticated
        """
        if not self.access_token:
            raise AuthenticationError("Not authenticated. Call authenticate() first.")
        
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        }
