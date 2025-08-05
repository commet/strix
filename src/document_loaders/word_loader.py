"""
Word Document Loader for STRIX
"""
from typing import List
from langchain.schema import Document
from langchain.document_loaders import Docx2txtLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from .base_loader import BaseDocumentLoader
import logging
from ..config import CHUNK_SIZE, CHUNK_OVERLAP

logger = logging.getLogger(__name__)

class STRIXWordLoader(BaseDocumentLoader):
    """Word document loader with Korean text optimization"""
    
    def __init__(self, doc_type: str = "internal"):
        super().__init__(doc_type)
        self.text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=CHUNK_SIZE,
            chunk_overlap=CHUNK_OVERLAP,
            separators=["\n\n", "\n", ".", "。", "!", "?", ";", ":", " ", ""],
            length_function=len
        )
    
    def load(self, file_path: str) -> List[Document]:
        """Load and process Word document"""
        try:
            # Load Word document
            loader = Docx2txtLoader(file_path)
            raw_documents = loader.load()
            
            # Get full text
            full_text = raw_documents[0].page_content if raw_documents else ""
            
            # Extract title
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
                    "file_type": "DOCX"
                })
                
                doc = Document(
                    page_content=chunk,
                    metadata=metadata
                )
                documents.append(doc)
            
            logger.info(f"Loaded {len(documents)} chunks from Word: {file_path}")
            return documents
            
        except Exception as e:
            logger.error(f"Error loading Word document {file_path}: {e}")
            raise