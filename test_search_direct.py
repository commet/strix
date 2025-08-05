"""
Direct test of search functionality
"""
import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from langchain_openai import OpenAIEmbeddings
from dotenv import load_dotenv

load_dotenv()

# Initialize
client = SupabaseClient()
embeddings = OpenAIEmbeddings()

# Test query
test_query = "전고체 배터리 투자 계획"
print(f"Testing search for: {test_query}")

# Generate embedding
query_embedding = embeddings.embed_query(test_query)
print(f"Generated embedding (length: {len(query_embedding)})")

# Search
results = client.search_similar_chunks(query_embedding, limit=5)
print(f"\nFound {len(results)} results:")

for i, result in enumerate(results):
    print(f"\n--- Result {i+1} ---")
    print(f"Similarity: {result.get('similarity', 'N/A')}")
    print(f"Title: {result.get('doc_title', 'N/A')}")
    print(f"Type: {result.get('doc_type', 'N/A')}")
    print(f"Content preview: {result.get('content', '')[:200]}...")