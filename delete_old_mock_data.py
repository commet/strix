"""
Delete old mock data with English types (internal/external)
Keep new data with Korean types (내부/외부)
"""
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient

def delete_old_mock_data():
    client = SupabaseClient()
    
    print("=== Deleting Old Mock Data ===\n")
    
    # Get all documents with English type
    try:
        # Delete documents with type='internal'
        internal_docs = client.client.table('documents').select('id, title, type').eq('type', 'internal').execute()
        print(f"Found {len(internal_docs.data)} documents with type='internal'")
        
        for doc in internal_docs.data:
            print(f"  Deleting: {doc['title']}")
            client.client.table('documents').delete().eq('id', doc['id']).execute()
        
        # Delete documents with type='external'
        external_docs = client.client.table('documents').select('id, title, type').eq('type', 'external').execute()
        print(f"\nFound {len(external_docs.data)} documents with type='external'")
        
        for doc in external_docs.data:
            print(f"  Deleting: {doc['title']}")
            client.client.table('documents').delete().eq('id', doc['id']).execute()
        
        print(f"\nTotal deleted: {len(internal_docs.data) + len(external_docs.data)} documents")
        
        # Check remaining documents
        remaining = client.client.table('documents').select('id, title, type').execute()
        print(f"\nRemaining documents: {len(remaining.data)}")
        
        # Show types of remaining documents
        types = {}
        for doc in remaining.data:
            doc_type = doc.get('type', 'unknown')
            types[doc_type] = types.get(doc_type, 0) + 1
        
        print("\nRemaining document types:")
        for dtype, count in types.items():
            print(f"  - {dtype}: {count} documents")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    delete_old_mock_data()