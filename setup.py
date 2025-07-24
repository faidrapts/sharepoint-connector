"""
SharePoint Document Scraper

A Python package for authenticating with SharePoint and downloading documents
using Microsoft Graph API with support for MFA and optional Bedrock knowledge base ingestion.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the contents of README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding='utf-8')

# Read requirements
requirements = []
requirements_file = this_directory / "requirements.txt"
if requirements_file.exists():
    requirements = requirements_file.read_text().strip().split('\n')
    requirements = [req.strip() for req in requirements if req.strip() and not req.startswith('#')]

setup(
    name="sharepoint-scraper",
    version="1.0.0",
    author="Faidra Anastasia Patsatzi",
    author_email="faidrapatsatzi@gmail.com",
    description="SharePoint document scraper with Microsoft Graph API and optional Bedrock integration",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/faidrapts/sharepoint-scraper",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Internet :: WWW/HTTP",
        "Topic :: Office/Business",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    python_requires=">=3.8",
    install_requires=requirements,
    extras_require={
        "bedrock": ["boto3>=1.34.0"],
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
            "mypy>=1.0.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "sharepoint-scraper=sharepoint_scraper.cli:main",
        ],
    },
    keywords="sharepoint, microsoft, graph, api, documents, scraper, bedrock, aws",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/sharepoint-scraper/issues",
        "Source": "https://github.com/yourusername/sharepoint-scraper",
        "Documentation": "https://github.com/yourusername/sharepoint-scraper#readme",
    },
)
