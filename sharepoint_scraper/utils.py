"""
Utility functions for SharePoint Scraper.
"""

import os
import json
import logging
from pathlib import Path
from typing import List, Dict, Any
from datetime import datetime


def setup_logging(log_level: str = "INFO", log_file: str = None) -> logging.Logger:
    """
    Setup logging configuration
    
    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_file: Optional log file path
        
    Returns:
        logging.Logger: Configured logger
    """
    # Convert string level to logging constant
    numeric_level = getattr(logging, log_level.upper(), logging.INFO)
    
    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Setup handlers
    handlers = []
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    handlers.append(console_handler)
    
    # File handler if specified
    if log_file:
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(formatter)
        handlers.append(file_handler)
    
    # Configure root logger
    logging.basicConfig(
        level=numeric_level,
        handlers=handlers,
        force=True
    )
    
    return logging.getLogger(__name__)


def save_documents_metadata(documents: List[Dict], output_file: str = "sharepoint_documents.json"):
    """
    Save document metadata to JSON file
    
    Args:
        documents: List of document metadata dictionaries
        output_file: Output file path
    """
    try:
        # Add timestamp
        metadata = {
            'timestamp': datetime.now().isoformat(),
            'total_documents': len(documents),
            'documents': documents
        }
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2, ensure_ascii=False)
        
        print(f"Document metadata saved to: {output_file}")
        
    except Exception as e:
        print(f"Error saving metadata: {str(e)}")


def load_documents_metadata(input_file: str) -> List[Dict]:
    """
    Load document metadata from JSON file
    
    Args:
        input_file: Input file path
        
    Returns:
        List[Dict]: List of document metadata dictionaries
    """
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        if isinstance(data, dict) and 'documents' in data:
            return data['documents']
        elif isinstance(data, list):
            return data
        else:
            raise ValueError("Invalid metadata file format")
            
    except Exception as e:
        print(f"Error loading metadata: {str(e)}")
        return []


def format_file_size(size_bytes: int) -> str:
    """
    Format file size in human readable format
    
    Args:
        size_bytes: Size in bytes
        
    Returns:
        str: Formatted size string
    """
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB", "TB"]
    import math
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    
    return f"{s} {size_names[i]}"


def summarize_documents(documents: List[Dict]) -> Dict[str, Any]:
    """
    Create summary statistics for documents
    
    Args:
        documents: List of document metadata dictionaries
        
    Returns:
        Dict[str, Any]: Summary statistics
    """
    if not documents:
        return {'total': 0, 'libraries': {}, 'total_size': 0, 'file_types': {}}
    
    libraries = {}
    total_size = 0
    file_types = {}
    
    for doc in documents:
        # Library statistics
        library = doc.get('library', 'Unknown')
        if library not in libraries:
            libraries[library] = {'count': 0, 'size': 0}
        
        libraries[library]['count'] += 1
        libraries[library]['size'] += doc.get('size', 0)
        
        # Total size
        total_size += doc.get('size', 0)
        
        # File type statistics
        name = doc.get('name', '')
        if '.' in name:
            ext = name.split('.')[-1].lower()
            file_types[ext] = file_types.get(ext, 0) + 1
    
    return {
        'total': len(documents),
        'libraries': libraries,
        'total_size': total_size,
        'total_size_formatted': format_file_size(total_size),
        'file_types': file_types
    }


def print_document_summary(documents: List[Dict]):
    """
    Print formatted document summary
    
    Args:
        documents: List of document metadata dictionaries
    """
    summary = summarize_documents(documents)
    
    print(f"\n📊 Document Summary:")
    print(f"Total documents: {summary['total']:,}")
    print(f"Total size: {summary['total_size_formatted']}")
    
    if summary['libraries']:
        print(f"\n📚 Libraries:")
        for lib_name, stats in summary['libraries'].items():
            print(f"  • {lib_name}: {stats['count']} files ({format_file_size(stats['size'])})")
    
    if summary['file_types']:
        print(f"\n📄 File types:")
        sorted_types = sorted(summary['file_types'].items(), key=lambda x: x[1], reverse=True)
        for ext, count in sorted_types[:10]:  # Show top 10
            print(f"  • .{ext}: {count} files")
        
        if len(sorted_types) > 10:
            print(f"  • ... and {len(sorted_types) - 10} more types")


def create_download_progress_callback():
    """
    Create a progress callback for download operations
    
    Returns:
        Callable: Progress callback function
    """
    def progress_callback(current: int, total: int):
        percentage = (current / total) * 100 if total > 0 else 0
        print(f"\rProgress: {current}/{total} ({percentage:.1f}%)", end="")
        if current == total:
            print()  # New line when complete
    
    return progress_callback


def validate_url(url: str) -> bool:
    """
    Validate SharePoint URL format
    
    Args:
        url: URL to validate
        
    Returns:
        bool: True if valid SharePoint URL
    """
    if not url:
        return False
    
    url = url.lower()
    return (
        url.startswith('https://') and
        'sharepoint.com' in url and
        not url.endswith('/')
    )


def sanitize_path(path: str) -> str:
    """
    Sanitize file/folder path for safe file system usage
    
    Args:
        path: Original path
        
    Returns:
        str: Sanitized path
    """
    if not path:
        return 'unknown'
    
    # Replace invalid characters
    invalid_chars = '<>:"/\\|?*'
    sanitized = path
    
    for char in invalid_chars:
        sanitized = sanitized.replace(char, '_')
    
    # Remove leading/trailing whitespace and dots
    sanitized = sanitized.strip('. ')
    
    # Split into components and sanitize each
    components = sanitized.split('/')
    sanitized_components = []
    
    for component in components:
        if component and component not in ['.', '..']:
            # Limit component length
            if len(component) > 100:
                component = component[:100]
            sanitized_components.append(component)
    
    return '/'.join(sanitized_components) if sanitized_components else 'unknown'
