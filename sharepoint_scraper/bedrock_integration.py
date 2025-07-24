"""
AWS Bedrock Knowledge Base integration for SharePoint documents.
"""

import boto3
import base64
import os
import logging
from pathlib import Path
from typing import Dict, Optional

from .exceptions import ConfigurationError

logger = logging.getLogger(__name__)


class BedrockIntegration:
    """Handles ingestion of documents into AWS Bedrock Knowledge Base"""
    
    def __init__(self, knowledge_base_id: str = None, data_source_id: str = None, region_name: str = None):
        """
        Initialize Bedrock integration
        
        Args:
            knowledge_base_id: Bedrock Knowledge Base ID
            data_source_id: Bedrock Data Source ID  
            region_name: AWS region name
        """
        self.knowledge_base_id = knowledge_base_id or os.environ.get('BEDROCK_KNOWLEDGE_BASE_ID')
        self.data_source_id = data_source_id or os.environ.get('BEDROCK_DATA_SOURCE_ID')
        self.region_name = region_name or os.environ.get('AWS_REGION', 'us-east-1')
        
        if not self.knowledge_base_id or not self.data_source_id:
            raise ConfigurationError(
                "Bedrock Knowledge Base ID and Data Source ID are required. "
                "Set BEDROCK_KNOWLEDGE_BASE_ID and BEDROCK_DATA_SOURCE_ID environment variables "
                "or pass them as parameters."
            )
        
        try:
            self.bedrock_agent = boto3.client('bedrock-agent', region_name=self.region_name)
        except Exception as e:
            raise ConfigurationError(f"Failed to initialize Bedrock client: {str(e)}")
    
    def ingest_document(self, document_path: str, document_id: str = None, title: str = None) -> Dict:
        """
        Ingest a document into Bedrock Knowledge Base
        
        Args:
            document_path: Path to the document file
            document_id: Unique ID for the document (defaults to filename stem)
            title: Title of the document (defaults to filename stem)
            
        Returns:
            Dict: Response from Bedrock agent API
            
        Raises:
            FileNotFoundError: If document file doesn't exist
            ConfigurationError: If Bedrock configuration is invalid
            Exception: For other ingestion errors
        """
        try:
            document_path = Path(document_path)
            
            # Validate file exists
            if not document_path.is_file():
                raise FileNotFoundError(f"Document not found: {document_path}")
            
            # Set defaults
            doc_id = document_id or document_path.stem
            doc_title = title or document_path.stem
            
            # Determine MIME type
            mime_type = self._get_mime_type(document_path)
            
            logger.info(f"Ingesting document into Bedrock: {document_path}")
            
            # Read and encode file
            with open(document_path, 'rb') as file:
                file_content = file.read()
                base64_content = base64.b64encode(file_content).decode('utf-8')
            
            # Prepare document for ingestion
            document_payload = {
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
                        },
                        {
                            'key': 'source',
                            'value': {
                                'stringValue': 'SharePoint',
                                'type': 'STRING'
                            }
                        }
                    ],
                    'type': 'IN_LINE_ATTRIBUTE'
                }
            }
            
            # Ingest document
            response = self.bedrock_agent.ingest_knowledge_base_documents(
                knowledgeBaseId=self.knowledge_base_id,
                dataSourceId=self.data_source_id,
                documents=[document_payload]
            )
            
            logger.info(f"Document ingestion successful: {doc_id}")
            return response
            
        except FileNotFoundError:
            raise
        except Exception as e:
            logger.error(f"Document ingestion failed: {str(e)}")
            raise
    
    def _get_mime_type(self, file_path: Path) -> str:
        """
        Determine MIME type based on file extension
        
        Args:
            file_path: Path to the file
            
        Returns:
            str: MIME type string
        """
        extension = file_path.suffix.lower()
        
        mime_types = {
            '.pdf': 'application/pdf',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.doc': 'application/msword',
            '.txt': 'text/plain',
            '.md': 'text/markdown',
            '.rtf': 'application/rtf',
            '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            '.ppt': 'application/vnd.ms-powerpoint',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.xls': 'application/vnd.ms-excel'
        }
        
        return mime_types.get(extension, 'application/octet-stream')
    
    def batch_ingest_documents(self, documents: list, progress_callback=None) -> Dict[str, Dict]:
        """
        Ingest multiple documents in batch
        
        Args:
            documents: List of document paths or dicts with metadata
            progress_callback: Optional callback function for progress updates
            
        Returns:
            Dict[str, Dict]: Results mapping document IDs to ingestion results
        """
        results = {}
        total = len(documents)
        
        for i, document in enumerate(documents):
            if isinstance(document, (str, Path)):
                doc_path = document
                doc_id = None
                doc_title = None
            else:
                doc_path = document.get('path')
                doc_id = document.get('id')
                doc_title = document.get('title')
            
            try:
                result = self.ingest_document(doc_path, doc_id, doc_title)
                results[doc_id or Path(doc_path).stem] = {'success': True, 'response': result}
            except Exception as e:
                results[doc_id or Path(doc_path).stem] = {'success': False, 'error': str(e)}
            
            if progress_callback:
                progress_callback(i + 1, total)
        
        return results
