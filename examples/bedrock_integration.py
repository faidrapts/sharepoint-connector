"""
Example: SharePoint to Bedrock Knowledge Base integration
"""

import os
from sharepoint_scraper import SharePointScraper, BedrockIntegration

def main():
    """Download SharePoint documents and ingest into Bedrock Knowledge Base"""
    
    # Check required environment variables
    required_vars = [
        'SHAREPOINT_SITE_URL',
        'AZURE_CLIENT_ID', 
        'BEDROCK_KNOWLEDGE_BASE_ID',
        'BEDROCK_DATA_SOURCE_ID'
    ]
    
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    if missing_vars:
        print(f"Missing required environment variables: {', '.join(missing_vars)}")
        print("\nPlease set these variables or create a .env file with:")
        for var in missing_vars:
            print(f"  {var}=your-value")
        return
    
    # Get configuration from environment
    site_url = os.getenv('SHAREPOINT_SITE_URL')
    knowledge_base_id = os.getenv('BEDROCK_KNOWLEDGE_BASE_ID')
    data_source_id = os.getenv('BEDROCK_DATA_SOURCE_ID')
    aws_region = os.getenv('AWS_REGION', 'us-east-1')
    
    print(f"SharePoint Site: {site_url}")
    print(f"Bedrock Knowledge Base: {knowledge_base_id}")
    print(f"AWS Region: {aws_region}")
    
    try:
        # Set up Bedrock integration
        print("\nSetting up Bedrock integration...")
        bedrock = BedrockIntegration(
            knowledge_base_id=knowledge_base_id,
            data_source_id=data_source_id,
            region_name=aws_region
        )
        
        # Create scraper with Bedrock integration
        scraper = SharePointScraper(site_url, bedrock=bedrock)
        
        # Authenticate
        print("Starting authentication...")
        if not scraper.authenticate():
            print("Authentication failed")
            return
        
        print("Authentication successful!")
        
        # Get documents
        print("Scanning for documents...")
        documents = scraper.get_documents()
        
        if not documents:
            print("No documents found")
            return
        
        print(f"Found {len(documents)} documents")
        
        # Filter documents (optional - e.g., only PDFs and Word docs)
        supported_types = ['.pdf', '.docx', '.doc', '.txt', '.md']
        filtered_docs = []
        
        for doc in documents:
            name = doc.get('name', '').lower()
            if any(name.endswith(ext) for ext in supported_types):
                filtered_docs.append(doc)
        
        print(f"Found {len(filtered_docs)} supported documents for Bedrock ingestion")
        
        if not filtered_docs:
            print("No supported document types found")
            return
        
        # Show sample documents
        print("\nSample documents to be ingested:")
        for i, doc in enumerate(filtered_docs[:5]):
            size_mb = doc.get('size', 0) / (1024 * 1024)
            print(f"  {i+1}. {doc['name']} ({size_mb:.1f} MB)")
        
        if len(filtered_docs) > 5:
            print(f"  ... and {len(filtered_docs) - 5} more")
        
        # Confirm with user
        response = input(f"\nDownload and ingest {len(filtered_docs)} documents into Bedrock? (y/N): ").strip().lower()
        if response not in ['y', 'yes']:
            print("Operation cancelled")
            return
        
        # Download and ingest with progress tracking
        def progress_callback(current, total):
            percentage = (current / total) * 100
            print(f"\rProcessing: {current}/{total} ({percentage:.1f}%)", end="")
            if current == total:
                print()
        
        print(f"\nDownloading and ingesting documents...")
        results = scraper.bulk_download_and_ingest(
            filtered_docs, 
            "downloads/", 
            progress_callback
        )
        
        # Show results
        successful = sum(1 for success in results.values() if success)
        failed = len(results) - successful
        
        print(f"\nIngestion complete!")
        print(f"✅ Successfully processed: {successful}")
        print(f"❌ Failed: {failed}")
        
        if failed > 0:
            print("\nFailed documents:")
            for doc_name, success in results.items():
                if not success:
                    print(f"  • {doc_name}")
        
        print(f"\nDocuments are now available in your Bedrock Knowledge Base: {knowledge_base_id}")
        
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
