"""
PDF Document Loader for STRIX
"""
from typing import List
from langchain.schema import Document
from langchain.document_loaders import PyPDFLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from .base_loader import BaseDocumentLoader
import logging
from ..config import CHUNK_SIZE, CHUNK_OVERLAP

logger = logging.getLogger(__name__)

class STRIXPDFLoader(BaseDocumentLoader):
    """PDF document loader with Korean text optimization"""
    
    def __init__(self, doc_type: str = "internal"):
        super().__init__(doc_type)
        self.text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=CHUNK_SIZE,
            chunk_overlap=CHUNK_OVERLAP,
            separators=["\n\n", "\n", ".", "。", "!", "?", ";", ":", " ", ""],
            length_function=len
        )
    
    def load(self, file_path: str) -> List[Document]:
        """Load and process PDF document"""
        try:
            # Load PDF
            loader = PyPDFLoader(file_path)
            pages = loader.load()
            
            # Combine all pages
            full_text = "\n\n".join([page.page_content for page in pages])
            
            # Extract title (first non-empty line)
            lines = full_text.split('\n')
            title = next((line.strip() for line in lines if line.strip()), "Untitled")
            
            # Split into chunks
            chunks = self.text_splitter.split_text(full_text)
            
            # Create documents
            documents = []
            for i, chunk in enumerate(chunks):
                metadata = self.extract_metadata(file_path, full_text)
                metadata.update({
                    "chunk_index": i,
                    "total_chunks": len(chunks),
                    "title": title,
                    "page_count": len(pages),
                    "file_type": "PDF"
                })
                
                doc = Document(
                    page_content=chunk,
                    metadata=metadata
                )
                documents.append(doc)
            
            logger.info(f"Loaded {len(documents)} chunks from PDF: {file_path}")
            return documents
            
        except Exception as e:
            logger.error(f"Error loading PDF {file_path}: {e}")
            raise