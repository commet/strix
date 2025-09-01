"""
Debug search functionality
"""
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from langchain_openai import OpenAIEmbeddings
import json

# Initialize
client = SupabaseClient()
embeddings = OpenAIEmbeddings()

# Test 1: Check what's in the database
print("=== CHECKING DATABASE CONTENTS ===")
docs_response = client.client.table('documents').select('id, title, type').execute()
print(f"Total documents: {len(docs_response.data)}")
print("\nDocument types:")
doc_types = {}
for doc in docs_response.data:
    doc_type = doc.get('type', 'unknown')
    doc_types[doc_type] = doc_types.get(doc_type, 0) + 1
    
for dtype, count in doc_types.items():
    print(f"  - {dtype}: {count} documents")

# Test 2: Check embeddings
print("\n=== CHECKING EMBEDDINGS ===")
embeddings_response = client.client.table('embeddings').select('id').limit(10).execute()
print(f"Total embeddings (sample): {len(embeddings_response.data)}")

# Test 3: Test search with different queries
test_queries = [
    "전고체 배터리",
    "SK온 합병",
    "CATL",
    "배터리 시장",
    "리튬 가격"
]

print("\n=== TESTING SEARCH QUERIES ===")
for query in test_queries:
    print(f"\nQuery: '{query}'")
    
    # Generate embedding
    query_embedding = embeddings.embed_query(query)
    
    # Search using RPC function
    results = client.search_similar_chunks(query_embedding, limit=3)
    
    if results:
        print(f"  Found {len(results)} results:")
        for i, result in enumerate(results, 1):
            title = result.get('doc_title', 'No title')
            doc_type = result.get('doc_type', 'unknown')
            similarity = result.get('similarity', 0)
            content_preview = result.get('content', '')[:100]
            print(f"    {i}. [{doc_type}] {title} (sim: {similarity:.3f})")
            print(f"       Preview: {content_preview}...")
    else:
        print("  No results found")

# Test 4: Direct RPC call
print("\n=== TESTING DIRECT RPC CALL ===")
test_embedding = embeddings.embed_query("배터리")
try:
    rpc_result = client.client.rpc('search_chunks', {
        'query_embedding': test_embedding,
        'match_threshold': 0.5,  # Lower threshold
        'match_count': 5
    }).execute()
    
    print(f"Direct RPC returned {len(rpc_result.data)} results")
    if rpc_result.data:
        for i, res in enumerate(rpc_result.data[:3], 1):
            print(f"  {i}. {res.get('doc_title', 'No title')} (sim: {res.get('similarity', 0):.3f})")
except Exception as e:
    print(f"RPC call failed: {e}")

# Test 5: Check if search function exists
print("\n=== CHECKING RPC FUNCTION ===")
try:
    # Try to call with minimal parameters
    test_result = client.client.rpc('search_chunks', {
        'query_embedding': [0.0] * 1536,  # Dummy embedding
        'match_threshold': 0.1,
        'match_count': 1
    }).execute()
    print("RPC function 'search_chunks' exists and is callable")
except Exception as e:
    print(f"RPC function error: {e}")
    print("\nPossible issues:")
    print("1. The RPC function 'search_chunks' may not exist in Supabase")
    print("2. The function parameters may be different")
    print("3. The function may have permission issues")