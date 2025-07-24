# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-07-24

### Added
- Initial release of SharePoint Scraper
- Microsoft Graph API integration for SharePoint access
- Interactive OAuth 2.0 authentication with PKCE flow
- Multi-factor authentication (MFA) support
- Document discovery and metadata extraction
- Bulk document downloading with progress tracking
- AWS Bedrock Knowledge Base integration
- Command-line interface (CLI) with multiple commands
- Python library API for programmatic use
- Comprehensive error handling and logging
- Document metadata export to JSON
- File size formatting and statistics
- Configuration management with environment variables
- Sanitized filename handling for cross-platform compatibility
- Recursive folder scanning
- Progress callbacks for long-running operations

### Features
- **Authentication**: Interactive OAuth flow with browser-based login
- **Document Discovery**: Scan all SharePoint sites and document libraries
- **Bulk Operations**: Download multiple documents efficiently
- **Bedrock Integration**: Ingest documents into AWS Bedrock knowledge bases
- **CLI Tools**: Command-line interface for easy automation
- **Python API**: Full programmatic access for custom integrations
- **Error Handling**: Comprehensive exception handling with specific error types
- **Logging**: Configurable logging levels and file output
- **Configuration**: Environment variable and file-based configuration
- **Progress Tracking**: Real-time progress updates for bulk operations

### CLI Commands
- `sharepoint-scraper test` - Test SharePoint connection
- `sharepoint-scraper scan` - Scan and list documents
- `sharepoint-scraper download` - Download documents
- `sharepoint-scraper config` - Show configuration status

### API Classes
- `SharePointScraper` - Main scraper class
- `SharePointAuth` - Authentication handling
- `BedrockIntegration` - AWS Bedrock integration
- `Config` - Configuration management

### Supported File Types
- Microsoft Office documents (Word, Excel, PowerPoint)
- PDF files
- Text files and markdown
- RTF documents
