"""
Check documents metadata in Supabase
"""
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from dotenv import load_dotenv

load_dotenv()

# Initialize Supabase client
client = SupabaseClient()

# Get all documents
response = client.client.table('documents').select('*').limit(10).execute()

print("=== Documents in Database ===\n")
for doc in response.data:
    print(f"ID: {doc['id'][:8]}...")
    print(f"Title: {doc.get('title', 'N/A')}")
    print(f"Type: {doc.get('type', 'N/A')}")
    print(f"Organization: {doc.get('organization', 'N/A')}")
    print(f"Created_at: {doc.get('created_at', 'N/A')}")
    print(f"Category: {doc.get('category', 'N/A')}")
    print(f"File Path: {doc.get('file_path', 'N/A')}")
    print("-" * 50)

# Get sample chunks with metadata
print("\n=== Sample Chunks ===\n")
chunks_response = client.client.table('chunks').select('*, document:documents(*)').limit(5).execute()

for chunk in chunks_response.data:
    doc = chunk.get('document', {})
    print(f"Chunk: {chunk['content'][:100]}...")
    print(f"Document Title: {doc.get('title', 'N/A')}")
    print(f"Document Org: {doc.get('organization', 'N/A')}")
    print("-" * 50)