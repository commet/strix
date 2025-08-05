"""
Test Supabase Connection and Setup
"""
import os
import sys
from dotenv import load_dotenv

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from config import OPENAI_API_KEY

def test_connection():
    """Test Supabase connection"""
    print("Testing Supabase connection...")
    
    try:
        # Initialize client
        client = SupabaseClient()
        print("[OK] Supabase client initialized")
        
        # Test simple query
        response = client.client.table('documents').select("*").limit(1).execute()
        print("[OK] Successfully connected to Supabase")
        
        # Check OpenAI API key
        if OPENAI_API_KEY and OPENAI_API_KEY.startswith("sk-"):
            print("[OK] OpenAI API key loaded")
        else:
            print("[ERROR] OpenAI API key not found or invalid")
            
        return True
        
    except Exception as e:
        print(f"[ERROR] Connection failed: {e}")
        return False

def check_tables():
    """Check if tables exist"""
    print("\nChecking tables...")
    
    client = SupabaseClient()
    tables_to_check = ['documents', 'chunks', 'embeddings', 'correlations']
    
    for table in tables_to_check:
        try:
            response = client.client.table(table).select("*").limit(1).execute()
            print(f"[OK] Table '{table}' exists")
        except Exception as e:
            print(f"[ERROR] Table '{table}' not found: {e}")

def print_setup_instructions():
    """Print setup instructions"""
    print("\n" + "="*50)
    print("SETUP INSTRUCTIONS")
    print("="*50)
    print("\n1. Go to your Supabase dashboard:")
    print("   https://app.supabase.com/project/qxrwyfxwwihskktsmjhj")
    print("\n2. Navigate to SQL Editor")
    print("\n3. Copy and run the contents of 'setup_supabase.sql'")
    print("\n4. After running the SQL, run this test again")
    print("="*50)

if __name__ == "__main__":
    # Load environment variables
    load_dotenv()
    
    print("STRIX Connection Test")
    print("="*50)
    
    # Test connection
    if test_connection():
        check_tables()
    else:
        print_setup_instructions()