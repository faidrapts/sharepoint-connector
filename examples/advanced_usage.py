"""
Example: Custom authentication and advanced usage
"""

import os
from dotenv import load_dotenv
from sharepoint_scraper import SharePointScraper, SharePointAuth
from sharepoint_scraper.utils import save_documents_metadata, load_documents_metadata

def main():
    """Advanced SharePoint scraper usage with custom authentication"""
    
    site_url = os.getenv('SHAREPOINT_SITE_URL')
    client_id = os.getenv('AZURE_CLIENT_ID')
    tenant_id = os.getenv('AZURE_TENANT_ID')  # Optional
    
    print(f"SharePoint Site: {site_url}")
    print(f"Azure Client ID: {client_id}")
    if tenant_id:
        print(f"Azure Tenant ID: {tenant_id}")
    
    try:
        # Custom authentication setup
        print("\nSetting up custom authentication...")
        auth = SharePointAuth(
            client_id=client_id,
            tenant_id=tenant_id,
            redirect_uri="http://localhost:8080/callback"
        )
        
        # Authenticate first
        print("Starting authentication...")
        access_token = auth.authenticate()
        print(f"Access token obtained: {access_token[:20]}...")
        
        # Create scraper with pre-authenticated session
        scraper = SharePointScraper(site_url, auth=auth)
        
        # Test connection (already authenticated)
        print("Testing connection...")
        if not scraper.test_connection():
            print("Connection test failed")
            return
        
        print("Connection successful!")
        
        # Get detailed site information
        site_info = scraper.get_site_info()
        print(f"\nSite Details:")
        print(f"  Name: {site_info.get('displayName', 'Unknown')}")
        print(f"  Description: {site_info.get('description', 'N/A')}")
        print(f"  Web URL: {site_info.get('webUrl', 'N/A')}")
        print(f"  Created: {site_info.get('createdDateTime', 'N/A')}")
        print(f"  Last Modified: {site_info.get('lastModifiedDateTime', 'N/A')}")
        
        # Check if we have cached document metadata
        metadata_file = "sharepoint_documents.json"
        
        if os.path.exists(metadata_file):
            print(f"\nFound existing metadata file: {metadata_file}")
            response = input("Use cached metadata? (Y/n): ").strip().lower()
            
            if response in ['', 'y', 'yes']:
                print("Loading cached metadata...")
                documents = load_documents_metadata(metadata_file)
                print(f"Loaded {len(documents)} documents from cache")
            else:
                documents = scan_documents(scraper, metadata_file)
        else:
            documents = scan_documents(scraper, metadata_file)
        
        if not documents:
            print("No documents found")
            return
        
        # Show detailed statistics
        show_detailed_stats(documents)
        
        # Interactive menu
        while True:
            print(f"\nOptions:")
            print("1. Download all documents")
            print("2. Download specific library")
            print("3. Download by file type")
            print("4. Show document details")
            print("5. Refresh document list")
            print("6. Exit")
            
            choice = input("\nSelect option (1-6): ").strip()
            
            if choice == '1':
                download_all_documents(scraper, documents)
            elif choice == '2':
                download_by_library(scraper, documents)
            elif choice == '3':
                download_by_file_type(scraper, documents)
            elif choice == '4':
                show_document_details(documents)
            elif choice == '5':
                documents = scan_documents(scraper, metadata_file)
                show_detailed_stats(documents)
            elif choice == '6':
                print("Goodbye!")
                break
            else:
                print("Invalid option")
    
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
    except Exception as e:
        print(f"Error: {e}")

def scan_documents(scraper, metadata_file):
    """Scan for documents and save metadata"""
    print("Scanning for documents...")
    documents = scraper.get_documents()
    
    if documents:
        save_documents_metadata(documents, metadata_file)
        print(f"Saved metadata to {metadata_file}")
    
    return documents

def show_detailed_stats(documents):
    """Show detailed document statistics"""
    from sharepoint_scraper.utils import summarize_documents, format_file_size
    
    summary = summarize_documents(documents)
    
    print(f"\n📊 Detailed Statistics:")
    print(f"Total documents: {summary['total']:,}")
    print(f"Total size: {summary['total_size_formatted']}")
    
    print(f"\n📚 Libraries:")
    for lib_name, stats in summary['libraries'].items():
        print(f"  • {lib_name}: {stats['count']} files ({format_file_size(stats['size'])})")
    
    print(f"\n📄 File types:")
    sorted_types = sorted(summary['file_types'].items(), key=lambda x: x[1], reverse=True)
    for ext, count in sorted_types:
        print(f"  • .{ext}: {count} files")

def download_all_documents(scraper, documents):
    """Download all documents"""
    response = input(f"Download all {len(documents)} documents? (y/N): ").strip().lower()
    if response in ['y', 'yes']:
        
        def progress_callback(current, total):
            percentage = (current / total) * 100
            print(f"\rDownloading: {current}/{total} ({percentage:.1f}%)", end="")
            if current == total:
                print()
        
        print("Downloading all documents...")
        results = scraper.bulk_download(documents, "downloads/", progress_callback)
        print(f"Downloaded {len(results)} files")

def download_by_library(scraper, documents):
    """Download documents from specific library"""
    # Get unique libraries
    libraries = set(doc.get('library', 'Unknown') for doc in documents)
    libraries = sorted(list(libraries))
    
    print("\nAvailable libraries:")
    for i, lib in enumerate(libraries, 1):
        lib_docs = [doc for doc in documents if doc.get('library') == lib]
        print(f"  {i}. {lib} ({len(lib_docs)} documents)")
    
    try:
        choice = int(input(f"\nSelect library (1-{len(libraries)}): "))
        if 1 <= choice <= len(libraries):
            selected_lib = libraries[choice - 1]
            lib_docs = [doc for doc in documents if doc.get('library') == selected_lib]
            
            print(f"\nDownloading {len(lib_docs)} documents from '{selected_lib}'...")
            
            def progress_callback(current, total):
                percentage = (current / total) * 100
                print(f"\rDownloading: {current}/{total} ({percentage:.1f}%)", end="")
                if current == total:
                    print()
            
            results = scraper.bulk_download(lib_docs, f"downloads/{selected_lib}/", progress_callback)
            print(f"Downloaded {len(results)} files to downloads/{selected_lib}/")
        else:
            print("Invalid selection")
    except ValueError:
        print("Invalid input")

def download_by_file_type(scraper, documents):
    """Download documents by file type"""
    # Get file types
    file_types = {}
    for doc in documents:
        name = doc.get('name', '')
        if '.' in name:
            ext = name.split('.')[-1].lower()
            file_types[ext] = file_types.get(ext, 0) + 1
    
    sorted_types = sorted(file_types.items(), key=lambda x: x[1], reverse=True)
    
    print("\nAvailable file types:")
    for i, (ext, count) in enumerate(sorted_types, 1):
        print(f"  {i}. .{ext} ({count} files)")
    
    try:
        choice = int(input(f"\nSelect file type (1-{len(sorted_types)}): "))
        if 1 <= choice <= len(sorted_types):
            selected_ext = sorted_types[choice - 1][0]
            filtered_docs = [
                doc for doc in documents 
                if doc.get('name', '').lower().endswith(f'.{selected_ext}')
            ]
            
            print(f"\nDownloading {len(filtered_docs)} .{selected_ext} files...")
            
            def progress_callback(current, total):
                percentage = (current / total) * 100
                print(f"\rDownloading: {current}/{total} ({percentage:.1f}%)", end="")
                if current == total:
                    print()
            
            results = scraper.bulk_download(filtered_docs, f"downloads/{selected_ext}/", progress_callback)
            print(f"Downloaded {len(results)} files to downloads/{selected_ext}/")
        else:
            print("Invalid selection")
    except ValueError:
        print("Invalid input")

def show_document_details(documents):
    """Show details for specific documents"""
    print(f"\nShowing first 10 documents:")
    for i, doc in enumerate(documents[:10], 1):
        size_mb = doc.get('size', 0) / (1024 * 1024)
        print(f"  {i:2d}. {doc['name']}")
        print(f"      Library: {doc.get('library', 'Unknown')}")
        print(f"      Size: {size_mb:.1f} MB")
        print(f"      Modified: {doc.get('modified', 'Unknown')}")
        if doc.get('path'):
            print(f"      Path: {doc['path']}")
        print()

if __name__ == "__main__":
    load_dotenv()  # Load environment variables from .env file
    main()
