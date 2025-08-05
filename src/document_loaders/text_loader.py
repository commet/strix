"""
Text Document Loader for STRIX
"""
from typing import List
from langchain.schema import Document
from langchain.text_splitter import RecursiveCharacterTextSplitter
from .base_loader import BaseDocumentLoader
import logging
from config import CHUNK_SIZE, CHUNK_OVERLAP

logger = logging.getLogger(__name__)

class STRIXTextLoader(BaseDocumentLoader):
    """Text document loader for testing"""
    
    def __init__(self, doc_type: str = "internal"):
        super().__init__(doc_type)
        self.text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=CHUNK_SIZE,
            chunk_overlap=CHUNK_OVERLAP,
            separators=["\n\n", "\n", ".", "。", "!", "?", ";", ":", " ", ""],
            length_function=len
        )
    
    def load(self, file_path: str) -> List[Document]:
        """Load and process text document"""
        try:
            # Read file
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Extract title (first line or filename)
            lines = content.split('\n')
            title = next((line.strip() for line in lines if line.strip()), "Untitled")
            
            # For very short title, use filename
            if len(title) < 10:
                import os
                title = os.path.splitext(os.path.basename(file_path))[0]
            
            # Split into chunks
            chunks = self.text_splitter.split_text(content)
            
            # Create documents
            documents = []
            for i, chunk in enumerate(chunks):
                metadata = self.extract_metadata(file_path, content)
                metadata.update({
                    "chunk_index": i,
                    "total_chunks": len(chunks),
                    "title": title,
                    "file_type": "TXT"
                })
                
                doc = Document(
                    page_content=chunk,
                    metadata=metadata
                )
                documents.append(doc)
            
            logger.info(f"Loaded {len(documents)} chunks from text file: {file_path}")
            return documents
            
        except Exception as e:
            logger.error(f"Error loading text file {file_path}: {e}")
            raise