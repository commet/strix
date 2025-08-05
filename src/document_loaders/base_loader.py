"""
Base Document Loader for STRIX
"""
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional
from langchain.schema import Document
import logging
from datetime import datetime
import hashlib

logger = logging.getLogger(__name__)

class BaseDocumentLoader(ABC):
    """Base class for all document loaders"""
    
    def __init__(self, doc_type: str = "internal"):
        self.doc_type = doc_type  # 'internal' or 'external'
        
    @abstractmethod
    def load(self, file_path: str) -> List[Document]:
        """Load documents from file"""
        pass
    
    def extract_metadata(self, file_path: str, content: str) -> Dict[str, Any]:
        """Extract metadata from document"""
        metadata = {
            "file_path": file_path,
            "type": self.doc_type,
            "loaded_at": datetime.now().isoformat(),
            "content_hash": self._generate_hash(content)
        }
        
        # Extract additional metadata based on file path and content
        if self.doc_type == "internal":
            metadata.update(self._extract_internal_metadata(file_path, content))
        else:
            metadata.update(self._extract_external_metadata(file_path, content))
            
        return metadata
    
    def _generate_hash(self, content: str) -> str:
        """Generate hash for content deduplication"""
        return hashlib.md5(content.encode()).hexdigest()
    
    def _extract_internal_metadata(self, file_path: str, content: str) -> Dict[str, Any]:
        """Extract metadata specific to internal documents"""
        metadata = {}
        
        # Extract organization from file path
        path_parts = file_path.split('/')
        for part in path_parts:
            if part in ["전략기획", "R&D", "경영지원", "생산", "영업마케팅"]:
                metadata["organization"] = part
                break
        
        # Extract date from filename (예: 2024_Q1_보고서.pptx)
        import re
        date_pattern = r'(\d{4})[_-]?([Q\d]{1,2})'
        date_match = re.search(date_pattern, file_path)
        if date_match:
            year = date_match.group(1)
            quarter_or_month = date_match.group(2)
            metadata["report_period"] = f"{year}_{quarter_or_month}"
        
        # Detect if it's executive interest
        exec_keywords = ["경영진", "CEO", "대표", "이사회", "전략", "중장기"]
        if any(keyword in content[:1000] for keyword in exec_keywords):
            metadata["exec_interest"] = True
            
        return metadata
    
    def _extract_external_metadata(self, file_path: str, content: str) -> Dict[str, Any]:
        """Extract metadata specific to external news"""
        from config import CATEGORY_KEYWORDS
        metadata = {}
        
        # Extract source from file path
        if "PR팀" in file_path:
            metadata["source"] = "PR팀"
        elif "Google" in file_path:
            metadata["source"] = "Google Alert"
        elif "Naver" in file_path:
            metadata["source"] = "Naver News"
        
        # Auto-categorize based on keywords
        content_lower = content.lower()
        category_scores = {}
        
        for category, keywords in CATEGORY_KEYWORDS.items():
            score = sum(1 for keyword in keywords if keyword.lower() in content_lower)
            if score > 0:
                category_scores[category] = score
        
        if category_scores:
            # Get top category
            top_category = max(category_scores, key=category_scores.get)
            metadata["category"] = top_category
            metadata["category_confidence"] = category_scores[top_category]
            
            # Also store other matching categories
            if len(category_scores) > 1:
                metadata["secondary_categories"] = [
                    cat for cat in category_scores if cat != top_category
                ]
        
        return metadata
    
    def create_documents(self, file_path: str, chunks: List[str]) -> List[Document]:
        """Create Document objects from chunks"""
        documents = []
        
        for i, chunk in enumerate(chunks):
            metadata = self.extract_metadata(file_path, chunk)
            metadata["chunk_index"] = i
            metadata["total_chunks"] = len(chunks)
            
            doc = Document(
                page_content=chunk,
                metadata=metadata
            )
            documents.append(doc)
            
        logger.info(f"Created {len(documents)} documents from {file_path}")
        return documents