"""
Example: Basic SharePoint document download
"""

import os
from sharepoint_scraper import SharePointScraper

def main():
    """Download documents from SharePoint"""
    
    # Set up environment (you can also use .env file)
    site_url = os.getenv('SHAREPOINT_SITE_URL', 'https://company.sharepoint.com/sites/mysite')
    client_id = os.getenv('AZURE_CLIENT_ID', 'your-client-id')
    
    if not site_url or not client_id:
        print("Please set SHAREPOINT_SITE_URL and AZURE_CLIENT_ID environment variables")
        return
    
    print(f"SharePoint Site: {site_url}")
    
    # Create scraper
    scraper = SharePointScraper(site_url)
    
    try:
        # Authenticate (opens browser)
        print("Starting authentication...")
        if not scraper.authenticate():
            print("Authentication failed")
            return
        
        print("Authentication successful!")
        
        # Test connection
        print("Testing connection...")
        if not scraper.test_connection():
            print("Connection test failed")
            return
        
        # Get site info
        site_info = scraper.get_site_info()
        print(f"Connected to: {site_info.get('displayName', 'Unknown Site')}")
        
        # Get all documents
        print("Scanning for documents...")
        documents = scraper.get_documents()
        
        if not documents:
            print("No documents found")
            return
        
        print(f"Found {len(documents)} documents")
        
        # Show first few documents
        print("\nSample documents:")
        for i, doc in enumerate(documents[:5]):
            size_mb = doc.get('size', 0) / (1024 * 1024)
            print(f"  {i+1}. {doc['name']} ({size_mb:.1f} MB)")
        
        if len(documents) > 5:
            print(f"  ... and {len(documents) - 5} more")
        
        # Ask user if they want to download
        response = input(f"\nDownload all {len(documents)} documents? (y/N): ").strip().lower()
        if response in ['y', 'yes']:
            
            # Download with progress tracking
            def progress_callback(current, total):
                percentage = (current / total) * 100
                print(f"\rDownloading: {current}/{total} ({percentage:.1f}%)", end="")
                if current == total:
                    print()  # New line when complete
            
            print(f"Downloading to 'downloads/' directory...")
            results = scraper.bulk_download(documents, "downloads/", progress_callback)
            
            print(f"\nDownload complete! Downloaded {len(results)} files.")
            
            # Show download summary
            total_size = sum(os.path.getsize(path) for path in results.values() if os.path.exists(path))
            print(f"Total downloaded: {total_size / (1024 * 1024):.1f} MB")
        
        else:
            print("Download cancelled")
    
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
