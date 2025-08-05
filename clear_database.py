"""
Clear all data from Supabase tables
"""
import os
import sys
from dotenv import load_dotenv

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient

def clear_all_data(auto_confirm=False):
    """Clear all data from database"""
    load_dotenv()
    
    client = SupabaseClient()
    
    print("This will DELETE ALL DATA from the database!")
    
    if not auto_confirm:
        confirm = input("Are you sure? (yes/no): ")
        if confirm.lower() != "yes":
            print("Cancelled.")
            return
    
    try:
        # Delete in order due to foreign key constraints
        tables = ['embeddings', 'chunks', 'correlations', 'documents', 'search_logs', 'keyword_learning']
        
        for table in tables:
            response = client.client.table(table).delete().neq('id', '00000000-0000-0000-0000-000000000000').execute()
            print(f"[OK] Cleared table: {table}")
        
        print("\n[OK] All data cleared successfully!")
        
    except Exception as e:
        print(f"[ERROR] Failed to clear data: {e}")

if __name__ == "__main__":
    import sys
    auto = "--auto" in sys.argv
    clear_all_data(auto_confirm=auto)