import boto3
import base64
import os
import logging
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def sanitize_filename(filename):
    """
    Sanitize the filename to prevent path traversal or other file-related security issues.
    
    Args:
        filename (str): Original filename
        
    Returns:
        str: Sanitized filename
    """
    return Path(filename).name

def ingest_document(document_path, document_id=None, title=None, region_name=None):
    """
    Ingest a document into a Bedrock knowledge base.
    
    Args:
        document_path (str): Path to the document file
        document_id (str, optional): Unique ID for the document. If None, uses sanitized filename
        title (str, optional): Title of the document. If None, uses sanitized filename
        region_name (str, optional): AWS region name. If None, uses the environment variable
        
    Returns:
        dict: Response from the Bedrock agent API
        
    Raises:
        ValueError: If required environment variables are missing
        FileNotFoundError: If the document file doesn't exist
        Exception: For other errors during ingestion
    """
    try:
        # Check for required environment variables
        knowledge_base_id = os.environ.get('BEDROCK_KNOWLEDGE_BASE_ID')
        data_source_id = os.environ.get('BEDROCK_DATA_SOURCE_ID')
        
        if not knowledge_base_id or not data_source_id:
            raise ValueError("Environment variables BEDROCK_KNOWLEDGE_BASE_ID and BEDROCK_DATA_SOURCE_ID must be set")
        
        # Use provided region or get from environment
        aws_region = region_name or os.environ.get('AWS_REGION', 'us-east-1')
        
        # Initialize the Bedrock agent client
        bedrock_agent = boto3.client('bedrock-agent', region_name=aws_region)
        
        # Sanitize document path and extract filename
        print(f"📄 Ingesting document: {document_path}")
        # Check if file exists
        if not Path(document_path).is_file():
            raise FileNotFoundError(f"Document not found: {document_path}")
        
        # Set document ID and title if not provided
        doc_id = document_id or Path(document_path).stem
        doc_title = title or Path(document_path).stem
        
        # Determine mime type based on file extension
        mime_type = "application/pdf"  # Default
        if document_id.endswith('.docx'):
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif document_id.endswith('.txt'):
            mime_type = "text/plain"
        logger.info(f"Ingesting document: {document_path}")
        
        # Read and encode the file
        with open(document_path, 'rb') as file:
            file_content = file.read()
            base64_content = base64.b64encode(file_content).decode('utf-8')
        
        # Ingest the document
        response = bedrock_agent.ingest_knowledge_base_documents(
            knowledgeBaseId=knowledge_base_id,
            dataSourceId=data_source_id,
            documents=[
                {
                    'content': {
                        'dataSourceType': 'CUSTOM',
                        'custom': {
                            'customDocumentIdentifier': {
                                'id': doc_id
                            },
                            'inlineContent': {
                                'byteContent': {
                                    'data': base64_content,
                                    'mimeType': mime_type
                                },
                                'type': 'BYTE'
                            },
                            'sourceType': 'IN_LINE'
                        }
                    },
                    'metadata': {
                        'inlineAttributes': [
                            {
                                'key': 'title',
                                'value': {
                                    'stringValue': doc_title,
                                    'type': 'STRING'
                                }
                            }
                        ],
                        'type': 'IN_LINE_ATTRIBUTE'
                    }
                }
            ]
        )
        
        logger.info(f"Document ingestion successful: {doc_id}")
        return response
        
    except ValueError as ve:
        logger.error(f"Configuration error: {ve}")
        raise
    except FileNotFoundError as fe:
        logger.error(f"File error: {fe}")
        raise
    except Exception as e:
        logger.error(f"Ingestion failed: {e}")
        raise
