"""
SharePoint Document Scraper

A Python package for authenticating with SharePoint and downloading documents
using Microsoft Graph API with support for MFA and optional Bedrock knowledge base ingestion.
"""

from .scraper import SharePointScraper
from .auth import SharePointAuth
from .bedrock_integration import BedrockIntegration
from .exceptions import SharePointError, AuthenticationError, DownloadError

__version__ = "1.0.0"
__author__ = "Your Name"
__email__ = "your.email@example.com"

__all__ = [
    "SharePointScraper",
    "SharePointAuth", 
    "BedrockIntegration",
    "SharePointError",
    "AuthenticationError",
    "DownloadError"
]
