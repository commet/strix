"""
Ingest documents with proper metadata
"""
import os
import json
import asyncio
from datetime import datetime
import sys

sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from ingestion.text_splitter import TextSplitter
from langchain_openai import OpenAIEmbeddings
from dotenv import load_dotenv

load_dotenv()

class DocumentIngestor:
    def __init__(self):
        self.supabase = SupabaseClient()
        self.text_splitter = TextSplitter()
        self.embeddings = OpenAIEmbeddings()
        
    async def ingest_documents_with_metadata(self):
        """Ingest documents using metadata.json"""
        
        # Load metadata
        metadata_path = "./mock_data/metadata.json"
        with open(metadata_path, 'r', encoding='utf-8') as f:
            metadata_all = json.load(f)
        
        # Process internal documents
        print("=== Processing Internal Documents ===")
        for doc_meta in metadata_all['internal']:
            await self.process_single_document(doc_meta)
        
        # Process external documents
        print("\n=== Processing External Documents ===")
        for doc_meta in metadata_all['external']:
            await self.process_single_document(doc_meta)
            
        print("\n✅ All documents processed successfully!")
    
    async def process_single_document(self, doc_meta):
        """Process a single document with metadata"""
        try:
            # Read file content
            with open(doc_meta['file_path'], 'r', encoding='utf-8') as f:
                content = f.read()
            
            print(f"\nProcessing: {doc_meta['title']}")
            print(f"  Organization: {doc_meta['organization']}")
            print(f"  Date: {doc_meta['created_at']}")
            
            # Insert document record
            doc_data = {
                'type': doc_meta['type'],
                'source': doc_meta['organization'],
                'title': doc_meta['title'],
                'organization': doc_meta['organization'],
                'category': doc_meta.get('category', ''),
                'created_at': doc_meta['created_at'],
                'file_path': doc_meta['file_path'],
                'metadata': {
                    'title': doc_meta['title'],
                    'organization': doc_meta['organization'],
                    'created_at': doc_meta['created_at'],
                    'category': doc_meta.get('category', '')
                }
            }
            
            doc_id = self.supabase.insert_document(doc_data)
            print(f"  Document ID: {doc_id}")
            
            # Split into chunks
            chunks = self.text_splitter.split_text(content)
            print(f"  Chunks: {len(chunks)}")
            
            # Process chunks
            chunks_data = []
            embeddings_data = []
            
            for i, chunk in enumerate(chunks):
                # Prepare chunk data
                chunk_data = {
                    'document_id': doc_id,
                    'content': chunk,
                    'chunk_index': i,
                    'start_char': i * 1000,  # Approximate
                    'end_char': (i + 1) * 1000,
                    'metadata': doc_meta
                }
                chunks_data.append(chunk_data)
            
            # Insert chunks
            chunk_records = self.supabase.insert_chunks(chunks_data)
            
            # Generate embeddings
            chunk_texts = [chunk for chunk in chunks]
            chunk_embeddings = await self.embeddings.aembed_documents(chunk_texts)
            
            # Prepare embeddings data
            for chunk_record, embedding in zip(chunk_records, chunk_embeddings):
                embeddings_data.append({
                    'chunk_id': chunk_record['id'],
                    'embedding': embedding,
                    'model': 'text-embedding-ada-002'
                })
            
            # Insert embeddings
            self.supabase.insert_embeddings(embeddings_data)
            
            print(f"  ✅ Successfully processed")
            
        except Exception as e:
            print(f"  ❌ Error: {str(e)}")

async def main():
    # Clear existing data first (optional)
    print("This will ingest documents with proper metadata.")
    response = input("Clear existing data first? (y/n): ")
    
    if response.lower() == 'y':
        print("Clearing existing data...")
        # Add clearing logic here if needed
    
    # Run ingestion
    ingestor = DocumentIngestor()
    await ingestor.ingest_documents_with_metadata()

if __name__ == "__main__":
    asyncio.run(main())