"""
Custom exceptions for SharePoint Scraper package.
"""


class SharePointError(Exception):
    """Base exception for SharePoint-related errors."""
    pass


class AuthenticationError(SharePointError):
    """Raised when authentication fails."""
    pass


class DownloadError(SharePointError):
    """Raised when document download fails."""
    pass


class ConfigurationError(SharePointError):
    """Raised when configuration is invalid or missing."""
    pass


class APIError(SharePointError):
    """Raised when Microsoft Graph API returns an error."""
    
    def __init__(self, message: str, status_code: int = None, response_data: dict = None):
        super().__init__(message)
        self.status_code = status_code
        self.response_data = response_data
