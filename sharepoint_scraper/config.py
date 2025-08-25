"""
Configuration management for SharePoint Scraper.
"""

import os
from typing import Optional, Dict, Any
from dataclasses import dataclass
from pathlib import Path

from .exceptions import ConfigurationError


@dataclass
class SharePointConfig:
    """SharePoint configuration settings"""
    site_url: str
    client_id: str
    tenant_id: Optional[str] = None
    client_secret: Optional[str] = None
    redirect_uri: str = None


@dataclass 
class BedrockConfig:
    """Bedrock configuration settings"""
    knowledge_base_id: str
    data_source_id: str
    region_name: str = "us-east-1"


class Config:
    """Configuration manager for SharePoint Scraper"""
    
    def __init__(self, config_file: str = None):
        """
        Initialize configuration
        
        Args:
            config_file: Optional path to configuration file
        """
        self.config_file = config_file
        self._load_config()
    
    def _load_config(self):
        """Load configuration from environment variables and config file"""
        # Load from .env file if it exists
        env_file = Path('.env')
        if env_file.exists():
            self._load_env_file(env_file)
    
    def _load_env_file(self, env_file: Path):
        """Load environment variables from .env file"""
        try:
            with open(env_file, 'r') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and '=' in line:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        value = value.strip().strip('"').strip("'")
                        os.environ[key] = value
        except Exception as e:
            raise ConfigurationError(f"Error loading .env file: {str(e)}")
    
    def get_sharepoint_config(self) -> SharePointConfig:
        """
        Get SharePoint configuration
        
        Returns:
            SharePointConfig: SharePoint configuration object
            
        Raises:
            ConfigurationError: If required configuration is missing
        """
        site_url = os.getenv('SHAREPOINT_SITE_URL')
        client_id = os.getenv('AZURE_CLIENT_ID')
        
        if not site_url:
            raise ConfigurationError("SHAREPOINT_SITE_URL environment variable is required")
        
        if not client_id:
            raise ConfigurationError("AZURE_CLIENT_ID environment variable is required")
        
        return SharePointConfig(
            site_url=site_url,
            client_id=client_id,
            tenant_id=os.getenv('AZURE_TENANT_ID'),
            client_secret=os.getenv('AZURE_CLIENT_SECRET'),
            redirect_uri=os.getenv('AZURE_REDIRECT_URI', "http://localhost:8080/callback")
        )
    
    def get_bedrock_config(self) -> Optional[BedrockConfig]:
        """
        Get Bedrock configuration
        
        Returns:
            Optional[BedrockConfig]: Bedrock configuration object or None if not configured
        """
        knowledge_base_id = os.getenv('BEDROCK_KNOWLEDGE_BASE_ID')
        data_source_id = os.getenv('BEDROCK_DATA_SOURCE_ID')
        
        if not knowledge_base_id or not data_source_id:
            return None
        
        return BedrockConfig(
            knowledge_base_id=knowledge_base_id,
            data_source_id=data_source_id,
            region_name=os.getenv('AWS_REGION', 'us-east-1')
        )
    
    def validate_config(self) -> Dict[str, Any]:
        """
        Validate configuration and return status
        
        Returns:
            Dict[str, Any]: Validation results
        """
        results = {
            'sharepoint': {'valid': False, 'errors': []},
            'bedrock': {'valid': False, 'errors': []},
            'aws': {'valid': False, 'errors': []}
        }
        
        # Validate SharePoint config
        try:
            self.get_sharepoint_config()
            results['sharepoint']['valid'] = True
        except ConfigurationError as e:
            results['sharepoint']['errors'].append(str(e))
        
        # Validate Bedrock config
        bedrock_config = self.get_bedrock_config()
        if bedrock_config:
            results['bedrock']['valid'] = True
        else:
            results['bedrock']['errors'].append("Bedrock configuration not found")
        
        # Validate AWS credentials
        if os.getenv('AWS_ACCESS_KEY_ID') or os.getenv('AWS_PROFILE'):
            results['aws']['valid'] = True
        else:
            results['aws']['errors'].append("AWS credentials not found")
        
        return results
