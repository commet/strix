"""
Test Document Ingestion Pipeline
"""
import os
import sys
import asyncio
from dotenv import load_dotenv

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from document_loaders.text_loader import STRIXTextLoader
from langchain_openai import OpenAIEmbeddings
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DocumentIngestion:
    def __init__(self):
        self.client = SupabaseClient()
        self.embeddings = OpenAIEmbeddings()
        
    def ingest_document(self, file_path: str, doc_type: str = "internal"):
        """Ingest a single document"""
        try:
            # Load document
            loader = STRIXTextLoader(doc_type=doc_type)
            documents = loader.load(file_path)
            logger.info(f"Loaded {len(documents)} chunks from {file_path}")
            
            # Extract metadata from first chunk
            if documents:
                first_doc = documents[0]
                
                # Insert document record
                doc_data = {
                    "type": doc_type,
                    "source": first_doc.metadata.get("source", "Unknown"),
                    "title": first_doc.metadata.get("title", "Untitled"),
                    "organization": first_doc.metadata.get("organization"),
                    "category": first_doc.metadata.get("category"),
                    "file_path": file_path,
                    "metadata": first_doc.metadata
                }
                
                doc_id = self.client.insert_document(doc_data)
                logger.info(f"Inserted document with ID: {doc_id}")
                
                # Insert chunks
                chunks_data = []
                for i, doc in enumerate(documents):
                    chunk_data = {
                        "document_id": doc_id,
                        "content": doc.page_content,
                        "chunk_index": i,
                        "metadata": doc.metadata
                    }
                    chunks_data.append(chunk_data)
                
                chunk_results = self.client.insert_chunks(chunks_data)
                logger.info(f"Inserted {len(chunk_results)} chunks")
                
                # Generate and insert embeddings
                embeddings_data = []
                for i, (chunk, doc) in enumerate(zip(chunk_results, documents)):
                    embedding = self.embeddings.embed_query(doc.page_content)
                    embedding_data = {
                        "chunk_id": chunk['id'],
                        "embedding": embedding
                    }
                    embeddings_data.append(embedding_data)
                
                self.client.insert_embeddings(embeddings_data)
                logger.info(f"Inserted {len(embeddings_data)} embeddings")
                
                return doc_id
                
        except Exception as e:
            logger.error(f"Error ingesting document: {e}")
            raise
    
    def ingest_directory(self, directory_path: str, doc_type: str = "internal"):
        """Ingest all documents in a directory"""
        ingested_count = 0
        
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                if file.endswith('.txt'):  # For now, only process text files
                    file_path = os.path.join(root, file)
                    try:
                        self.ingest_document(file_path, doc_type)
                        ingested_count += 1
                    except Exception as e:
                        logger.error(f"Failed to ingest {file_path}: {e}")
        
        return ingested_count

def main():
    """Main test function"""
    load_dotenv()
    
    print("STRIX Document Ingestion Test")
    print("="*50)
    
    # Initialize ingestion
    ingestion = DocumentIngestion()
    
    # Check if mock data exists
    if not os.path.exists("./mock_data"):
        print("[ERROR] Mock data not found. Run 'python create_mock_data.py' first.")
        return
    
    # Test ingesting a single document
    print("\n1. Testing single document ingestion...")
    test_file = "./mock_data/internal/전략기획/2024_Q1_배터리사업_중장기전략.txt"
    if os.path.exists(test_file):
        try:
            doc_id = ingestion.ingest_document(test_file, "internal")
            print(f"[OK] Successfully ingested document: {doc_id}")
        except Exception as e:
            print(f"[ERROR] Failed to ingest document: {e}")
    
    # Test ingesting all documents
    print("\n2. Testing batch ingestion...")
    auto_ingest = "--auto" in sys.argv
    
    if auto_ingest:
        user_input = 'y'
    else:
        user_input = input("Ingest all mock documents? (y/n): ")
    
    if user_input.lower() == 'y':
        # Ingest internal documents
        print("\nIngesting internal documents...")
        internal_count = ingestion.ingest_directory("./mock_data/internal", "internal")
        print(f"[OK] Ingested {internal_count} internal documents")
        
        # Ingest external news
        print("\nIngesting external news...")
        external_count = ingestion.ingest_directory("./mock_data/external", "external")
        print(f"[OK] Ingested {external_count} external documents")
        
        print(f"\n[OK] Total documents ingested: {internal_count + external_count}")
    
    # Show statistics
    print("\n3. Database statistics:")
    try:
        stats = ingestion.client.client.rpc('get_document_stats').execute()
        for stat in stats.data:
            print(f"  - {stat['stat_type']}: {stat['stat_value']}")
    except:
        print("  (Statistics function not available)")

if __name__ == "__main__":
    main()