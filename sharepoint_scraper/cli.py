"""
Command Line Interface for SharePoint Scraper.
"""

import argparse
import sys
from pathlib import Path
from typing import Optional

from .scraper import SharePointScraper
from .auth import SharePointAuth
from .bedrock_integration import BedrockIntegration
from .config import Config
from .utils import (
    setup_logging, 
    save_documents_metadata, 
    print_document_summary,
    create_download_progress_callback
)
from .exceptions import SharePointError, ConfigurationError


def create_parser() -> argparse.ArgumentParser:
    """Create command line argument parser"""
    parser = argparse.ArgumentParser(
        description="SharePoint Document Scraper with Microsoft Graph API",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic usage - scan documents
  sharepoint-scraper --site-url https://contoso.sharepoint.com/sites/mysite scan

  # Download all documents
  sharepoint-scraper --site-url https://contoso.sharepoint.com/sites/mysite download

  # Download and ingest to Bedrock
  sharepoint-scraper --site-url https://contoso.sharepoint.com/sites/mysite download --bedrock

  # Test connection
  sharepoint-scraper --site-url https://contoso.sharepoint.com/sites/mysite test

Environment Variables:
  SHAREPOINT_SITE_URL      SharePoint site URL
  AZURE_CLIENT_ID          Azure AD App Registration Client ID
  AZURE_TENANT_ID          Azure AD Tenant ID (optional)
  AZURE_CLIENT_SECRET      Azure AD Client Secret (optional)
  BEDROCK_KNOWLEDGE_BASE_ID  Bedrock Knowledge Base ID
  BEDROCK_DATA_SOURCE_ID   Bedrock Data Source ID
  AWS_REGION               AWS Region (default: us-east-1)
        """
    )
    
    # Global options
    parser.add_argument(
        '--site-url', 
        help='SharePoint site URL'
    )
    parser.add_argument(
        '--client-id',
        help='Azure AD App Registration Client ID'
    )
    parser.add_argument(
        '--tenant-id',
        help='Azure AD Tenant ID'
    )
    parser.add_argument(
        '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='Logging level (default: INFO)'
    )
    parser.add_argument(
        '--log-file',
        help='Log file path'
    )
    parser.add_argument(
        '--config-file',
        help='Configuration file path'
    )
    
    # Subcommands
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # Test command
    test_parser = subparsers.add_parser('test', help='Test SharePoint connection')
    
    # Scan command
    scan_parser = subparsers.add_parser('scan', help='Scan and list documents')
    scan_parser.add_argument(
        '--output',
        default='sharepoint_documents.json',
        help='Output file for document metadata (default: sharepoint_documents.json)'
    )
    
    # Download command
    download_parser = subparsers.add_parser('download', help='Download documents')
    download_parser.add_argument(
        '--output-dir',
        default='downloads',
        help='Download directory (default: downloads)'
    )
    download_parser.add_argument(
        '--bedrock',
        action='store_true',
        help='Also ingest documents into Bedrock knowledge base'
    )
    download_parser.add_argument(
        '--metadata-file',
        help='Use existing metadata file instead of scanning'
    )
    
    # Config command
    config_parser = subparsers.add_parser('config', help='Show configuration status')
    
    return parser


def setup_scraper(args) -> tuple[SharePointScraper, Optional[BedrockIntegration]]:
    """Setup scraper instances based on arguments"""
    # Load configuration
    config = Config(args.config_file)
    
    # Get SharePoint config
    sharepoint_config = config.get_sharepoint_config()
    
    # Override with command line arguments
    if args.site_url:
        sharepoint_config.site_url = args.site_url
    if args.client_id:
        sharepoint_config.client_id = args.client_id
    if args.tenant_id:
        sharepoint_config.tenant_id = args.tenant_id
    
    # Setup authentication
    auth = SharePointAuth(
        client_id=sharepoint_config.client_id,
        tenant_id=sharepoint_config.tenant_id,
        redirect_uri=sharepoint_config.redirect_uri
    )
    
    # Setup Bedrock integration if configured
    bedrock = None
    bedrock_config = config.get_bedrock_config()
    if bedrock_config:
        try:
            bedrock = BedrockIntegration(
                knowledge_base_id=bedrock_config.knowledge_base_id,
                data_source_id=bedrock_config.data_source_id,
                region_name=bedrock_config.region_name
            )
        except Exception as e:
            print(f"Warning: Bedrock integration not available: {str(e)}")
    
    # Create scraper
    scraper = SharePointScraper(
        site_url=sharepoint_config.site_url,
        auth=auth,
        bedrock=bedrock
    )
    
    return scraper, bedrock


def cmd_test(args) -> int:
    """Test SharePoint connection"""
    try:
        scraper, _ = setup_scraper(args)
        
        print("Testing SharePoint connection...")
        print(f"Site URL: {scraper.site_url}")
        
        # Authenticate
        print("\n🔐 Authenticating...")
        if not scraper.authenticate():
            print("❌ Authentication failed")
            return 1
        
        print("✅ Authentication successful")
        
        # Test connection
        print("\n🔍 Testing connection...")
        if not scraper.test_connection():
            print("❌ Connection test failed")
            return 1
        
        print("✅ Connection test successful")
        
        # Get site info
        try:
            site_info = scraper.get_site_info()
            print(f"\n📋 Site Information:")
            print(f"  Name: {site_info.get('displayName', 'Unknown')}")
            print(f"  Description: {site_info.get('description', 'N/A')}")
            print(f"  Web URL: {site_info.get('webUrl', 'N/A')}")
        except Exception as e:
            print(f"Warning: Could not get site info: {str(e)}")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return 1


def cmd_scan(args) -> int:
    """Scan and list documents"""
    try:
        scraper, _ = setup_scraper(args)
        
        print("Scanning SharePoint documents...")
        print(f"Site URL: {scraper.site_url}")
        
        # Authenticate
        print("\n🔐 Authenticating...")
        if not scraper.authenticate():
            print("❌ Authentication failed")
            return 1
        
        # Get documents
        print("\n📄 Scanning for documents...")
        documents = scraper.get_documents()
        
        if not documents:
            print("No documents found")
            return 0
        
        # Show summary
        print_document_summary(documents)
        
        # Save metadata
        save_documents_metadata(documents, args.output)
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return 1


def cmd_download(args) -> int:
    """Download documents"""
    try:
        scraper, bedrock = setup_scraper(args)
        
        # Check Bedrock requirement
        if args.bedrock and not bedrock:
            print("❌ Bedrock integration not configured but --bedrock flag was used")
            return 1
        
        documents = []
        
        if args.metadata_file:
            # Load from existing metadata file
            from .utils import load_documents_metadata
            print(f"Loading documents from metadata file: {args.metadata_file}")
            documents = load_documents_metadata(args.metadata_file)
        else:
            # Scan for documents
            print("Scanning SharePoint documents...")
            print(f"Site URL: {scraper.site_url}")
            
            # Authenticate
            print("\n🔐 Authenticating...")
            if not scraper.authenticate():
                print("❌ Authentication failed")
                return 1
            
            # Get documents
            print("\n📄 Scanning for documents...")
            documents = scraper.get_documents()
        
        if not documents:
            print("No documents found")
            return 0
        
        # Show summary
        print_document_summary(documents)
        
        # Confirm download
        response = input(f"\nDownload {len(documents)} documents? (y/N): ").strip().lower()
        if response not in ['y', 'yes']:
            print("Download cancelled")
            return 0
        
        # Setup progress callback
        progress_callback = create_download_progress_callback()
        
        # Download documents
        print(f"\n⬇️  Downloading to: {args.output_dir}")
        
        if args.bedrock:
            print("📚 Will also ingest into Bedrock knowledge base")
            results = scraper.bulk_download_and_ingest(
                documents, 
                args.output_dir,
                progress_callback
            )
            
            successful = sum(1 for success in results.values() if success)
            print(f"\n✅ Downloaded and ingested {successful}/{len(documents)} documents")
        else:
            results = scraper.bulk_download(
                documents,
                args.output_dir, 
                progress_callback
            )
            
            print(f"\n✅ Downloaded {len(results)}/{len(documents)} documents")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return 1


def cmd_config(args) -> int:
    """Show configuration status"""
    try:
        config = Config(args.config_file)
        validation_results = config.validate_config()
        
        print("Configuration Status:")
        print("=" * 50)
        
        for component, results in validation_results.items():
            status = "✅ Valid" if results['valid'] else "❌ Invalid"
            print(f"{component.title()}: {status}")
            
            if results['errors']:
                for error in results['errors']:
                    print(f"  • {error}")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return 1


def main():
    """Main CLI entry point"""
    parser = create_parser()
    args = parser.parse_args()
    
    # Setup logging
    setup_logging(args.log_level, args.log_file)
    
    # Check for command
    if not args.command:
        parser.print_help()
        return 1
    
    # Route to command handlers
    try:
        if args.command == 'test':
            return cmd_test(args)
        elif args.command == 'scan':
            return cmd_scan(args)
        elif args.command == 'download':
            return cmd_download(args)
        elif args.command == 'config':
            return cmd_config(args)
        else:
            print(f"Unknown command: {args.command}")
            return 1
            
    except KeyboardInterrupt:
        print("\n⏹️  Operation cancelled by user")
        return 1
    except Exception as e:
        print(f"❌ Unexpected error: {str(e)}")
        return 1


if __name__ == '__main__':
    sys.exit(main())
