# SharePoint Document Scraper

A Python package for authenticating with SharePoint and downloading documents using Microsoft Graph API with support for MFA and optional AWS Bedrock knowledge base ingestion.

## Features

- 🔐 **Easy Authentication**: Interactive OAuth 2.0 flow with MFA support
- 📄 **Document Discovery**: Scan and catalog all documents in SharePoint sites
- ⬇️ **Bulk Download**: Download multiple documents with progress tracking
- 🧠 **Bedrock Integration**: Optional ingestion into AWS Bedrock knowledge bases
- 🛠️ **CLI & Library**: Use as command-line tool or Python library
- 📊 **Rich Metadata**: Detailed document information and statistics

## Installation

```bash
pip install sharepoint-scraper
```

For AWS Bedrock integration:
```bash
pip install sharepoint-scraper[bedrock]
```

## Quick Start

### Command Line Usage

1. **Set up environment variables** (if you do not have the required Azure values please first follow the "Azure AD App Registration Setup" instructions from the Configuration section below):
```bash
export SHAREPOINT_SITE_URL="https://yourcompany.sharepoint.com/sites/yoursite"
export AZURE_CLIENT_ID="your-azure-app-client-id"
export AZURE_CLIENT_SECRET="your-azure-client-secret"
export AZURE_TENANT_ID="your-tenant-id"
```

2. **Test connection**:
```bash
sharepoint-scraper test
```

3. **Scan documents**:
```bash
sharepoint-scraper scan
```

4. **Download documents**:
```bash
sharepoint-scraper download
```

5. **Download and ingest to Bedrock**:
```bash
# Set Bedrock environment variables
export BEDROCK_KNOWLEDGE_BASE_ID="your-kb-id"
export BEDROCK_DATA_SOURCE_ID="your-data-source-id"
export AWS_REGION="us-east-1"

sharepoint-scraper download --bedrock
```

### Python Library Usage

```python
from sharepoint_scraper import SharePointScraper, SharePointAuth, BedrockIntegration

# Basic usage
scraper = SharePointScraper("https://yourcompany.sharepoint.com/sites/yoursite")

# Authenticate (opens browser for OAuth)
scraper.authenticate()

# Get all documents
documents = scraper.get_documents()

# Download documents
scraper.bulk_download(documents, "downloads/")

# With Bedrock integration
bedrock = BedrockIntegration(
    knowledge_base_id="your-kb-id",
    data_source_id="your-data-source-id"
)

scraper = SharePointScraper(site_url, bedrock=bedrock)
scraper.authenticate()

# Download and ingest to Bedrock
documents = scraper.get_documents()
scraper.bulk_download_and_ingest(documents, "downloads/")
```

## Configuration

### Environment Variables

#### Required for SharePoint
- `SHAREPOINT_SITE_URL`: Your SharePoint site URL
- `AZURE_CLIENT_ID`: Azure AD App Registration Client ID
- `AZURE_TENANT_ID`: Azure AD Tenant ID
- `AZURE_CLIENT_SECRET`: Azure AD Client Secret

#### Required for Bedrock Integration
- `BEDROCK_KNOWLEDGE_BASE_ID`: AWS Bedrock Knowledge Base ID
- `BEDROCK_DATA_SOURCE_ID`: AWS Bedrock Data Source ID
- `AWS_REGION`: AWS Region (default: us-east-1)

### Azure AD App Registration Setup

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Set these values:
   - **Name**: SharePoint Scraper
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: Web → `http://localhost:8080/callback` or configure it based on your application
4. After creation, note the **Application (client) ID**, **Directory (tenant) ID**, and navigate to Certificates & secrets > Create new client secret. Also note the **Client Secret ID**.
5. Go to "API permissions" → Add permission → Microsoft Graph → Delegated permissions
6. Add these permissions:
   - `Sites.Read.All`
   - `Files.Read.All` 
7. Click "Grant admin consent"

## CLI Reference

```bash
# Test connection
sharepoint-scraper test

# Scan and save document metadata
sharepoint-scraper scan --output documents.json

# Download documents
sharepoint-scraper download --output-dir downloads/

# Download and ingest to Bedrock
sharepoint-scraper download --bedrock

# Use existing metadata file
sharepoint-scraper download --metadata-file documents.json

# Show configuration status
sharepoint-scraper config

# Get help
sharepoint-scraper --help
sharepoint-scraper download --help
```

## Python API Reference

### SharePointScraper

Main class for SharePoint operations.

```python
from sharepoint_scraper import SharePointScraper

scraper = SharePointScraper(site_url, auth=None, bedrock=None)

# Authentication
scraper.authenticate() -> bool

# Document operations
scraper.get_documents() -> List[Dict]
scraper.download_document(document, download_path) -> Optional[str]
scraper.bulk_download(documents, download_path) -> Dict[str, str]

# Bedrock integration
scraper.download_and_ingest_document(document, download_path) -> bool
scraper.bulk_download_and_ingest(documents, download_path) -> Dict[str, bool]

# Connection testing
scraper.test_connection() -> bool
scraper.get_site_info() -> Dict
```

### SharePointAuth

Handles authentication with Microsoft Graph API.

```python
from sharepoint_scraper import SharePointAuth

auth = SharePointAuth(client_id, tenant_id, redirect_uri)

# Authenticate user
auth.authenticate() -> str  # Returns access token

# Check status
auth.is_authenticated() -> bool
auth.get_access_token() -> Optional[str]
auth.get_auth_headers() -> Dict[str, str]
```

### BedrockIntegration

Optional AWS Bedrock knowledge base integration.

```python
from sharepoint_scraper import BedrockIntegration

bedrock = BedrockIntegration(knowledge_base_id, data_source_id, region_name)

# Ingest single document
bedrock.ingest_document(document_path, document_id, title) -> Dict

# Batch ingest
bedrock.batch_ingest_documents(documents, progress_callback) -> Dict[str, Dict]
```

## Examples

### Example 1: Basic Document Download

```python
import os
from sharepoint_scraper import SharePointScraper

# Set up
os.environ['SHAREPOINT_SITE_URL'] = 'https://company.sharepoint.com/sites/mysite'
os.environ['AZURE_CLIENT_ID'] = 'your-client-id'

# Create scraper and authenticate
scraper = SharePointScraper(os.environ['SHAREPOINT_SITE_URL'])
scraper.authenticate()  # Opens browser for login

# Get and download all documents
documents = scraper.get_documents()
print(f"Found {len(documents)} documents")

# Download with progress
def progress(current, total):
    print(f"Progress: {current}/{total}")

results = scraper.bulk_download(documents, "downloads/", progress)
print(f"Downloaded {len(results)} documents")
```

### Example 2: Bedrock Integration

```python
import os
from sharepoint_scraper import SharePointScraper, BedrockIntegration

# Configure environment
os.environ.update({
    'SHAREPOINT_SITE_URL': 'https://company.sharepoint.com/sites/mysite',
    'AZURE_CLIENT_ID': 'your-client-id',
    'BEDROCK_KNOWLEDGE_BASE_ID': 'your-kb-id',
    'BEDROCK_DATA_SOURCE_ID': 'your-ds-id',
    'AWS_REGION': 'us-east-1'
})

# Set up Bedrock integration
bedrock = BedrockIntegration(
    knowledge_base_id=os.environ['BEDROCK_KNOWLEDGE_BASE_ID'],
    data_source_id=os.environ['BEDROCK_DATA_SOURCE_ID']
)

# Create scraper with Bedrock
scraper = SharePointScraper(
    site_url=os.environ['SHAREPOINT_SITE_URL'],
    bedrock=bedrock
)

# Authenticate and process
scraper.authenticate()
documents = scraper.get_documents()

# Download and ingest to Bedrock
results = scraper.bulk_download_and_ingest(documents, "downloads/")

successful = sum(1 for success in results.values() if success)
print(f"Successfully processed {successful}/{len(documents)} documents")
```

### Example 3: Custom Authentication

```python
from sharepoint_scraper import SharePointScraper, SharePointAuth

# Custom auth setup
auth = SharePointAuth(
    client_id="your-client-id",
    tenant_id="your-tenant-id",  # Optional
    redirect_uri="http://localhost:8080/callback"  # Default
)

# Authenticate
token = auth.authenticate()
print(f"Access token: {token[:20]}...")

# Use with scraper
scraper = SharePointScraper(
    site_url="https://company.sharepoint.com/sites/mysite",
    auth=auth
)

# Now scraper is already authenticated
documents = scraper.get_documents()
```

## Error Handling

The package includes comprehensive error handling:

```python
from sharepoint_scraper import (
    SharePointScraper, 
    SharePointError, 
    AuthenticationError, 
    DownloadError,
    ConfigurationError
)

try:
    scraper = SharePointScraper(site_url)
    scraper.authenticate()
    documents = scraper.get_documents()
    
except ConfigurationError as e:
    print(f"Configuration issue: {e}")
except AuthenticationError as e:
    print(f"Authentication failed: {e}")
except SharePointError as e:
    print(f"SharePoint error: {e}")
except Exception as e:
    print(f"Unexpected error: {e}")
```

## Troubleshooting

### Common Issues

1. **Authentication fails**
   - Ensure Azure AD app has correct permissions
   - Check that redirect URI is set to `http://localhost:8080/callback`
   - Verify client ID and tenant ID

2. **"Access denied" errors**
   - User needs read permissions on SharePoint site
   - Azure AD app needs admin consent for permissions

3. **Documents not found**
   - Check SharePoint site URL format
   - Ensure user has access to document libraries
   - Site might have restricted access

4. **Bedrock ingestion fails**
   - Verify AWS credentials are configured
   - Check Bedrock knowledge base and data source IDs
   - Ensure proper IAM permissions for Bedrock

### Debug Mode

Enable debug logging:

```bash
sharepoint-scraper --log-level DEBUG scan
```

Or in Python:
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Make changes and add tests
4. Run tests: `pytest`
5. Submit a pull request

## License

MIT License - see LICENSE file for details.

## Changelog

### v1.0.0
- Initial release
- Microsoft Graph API integration
- Interactive OAuth authentication with MFA support
- Document scanning and downloading
- AWS Bedrock knowledge base integration
- Command-line interface
- Comprehensive error handling and logging
