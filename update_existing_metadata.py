"""
Update metadata for existing documents in database
"""
import sys
import os
import re
from datetime import datetime

sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from dotenv import load_dotenv

load_dotenv()

# Organization mapping from file path
ORGANIZATION_MAP = {
    "전략기획": "전략기획",
    "R&D": "R&D",
    "경영지원": "경영지원",
    "생산": "생산",
    "영업마케팅": "영업마케팅",
    "PR팀_AM": "PR팀_AM",
    "PR팀_PM": "PR팀_PM",
    "Google_Alert": "Google Alert",
    "Naver_News": "Naver News"
}

def extract_date_from_filename(filename):
    """Extract date from filename"""
    # Pattern: 2024_01_15, 2024-01-15, 2024_Q1
    patterns = [
        r'(\d{4})[_-](\d{1,2})[_-](\d{1,2})',  # YYYY_MM_DD
        r'(\d{4})[_-]Q(\d)',  # YYYY_Q1
        r'(\d{4})[_-](\d{1,2})'  # YYYY_MM
    ]
    
    for pattern in patterns:
        match = re.search(pattern, filename)
        if match:
            if len(match.groups()) == 3:
                # Full date
                year, month, day = match.groups()
                return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            elif 'Q' in pattern:
                # Quarter
                year, quarter = match.groups()
                month = str((int(quarter) - 1) * 3 + 1).zfill(2)
                return f"{year}-{month}-01"
            else:
                # Year and month only
                year, month = match.groups()
                return f"{year}-{month.zfill(2)}-01"
    
    return None

def extract_title_from_filename(filename):
    """Extract title from filename"""
    # Remove extension
    name = os.path.splitext(filename)[0]
    
    # Remove date pattern
    name = re.sub(r'\d{4}[_-]\d{1,2}[_-]\d{1,2}[_-]?', '', name)
    name = re.sub(r'\d{4}[_-]Q\d[_-]?', '', name)
    name = re.sub(r'\d{4}[_-]\d{1,2}[_-]?', '', name)
    
    # Replace underscores with spaces
    name = name.replace('_', ' ')
    
    # Clean up
    name = ' '.join(name.split())
    
    return name.strip()

def update_document_metadata():
    """Update metadata for existing documents"""
    
    client = SupabaseClient()
    
    # Get all documents
    print("Fetching all documents...")
    response = client.client.table('documents').select('*').execute()
    documents = response.data
    
    print(f"Found {len(documents)} documents\n")
    
    updated_count = 0
    
    for doc in documents:
        try:
            file_path = doc.get('file_path', '')
            if not file_path:
                continue
                
            # Extract info from file path
            filename = os.path.basename(file_path)
            folder_parts = file_path.split('/')
            
            # Find organization
            organization = None
            for part in folder_parts:
                if part in ORGANIZATION_MAP:
                    organization = ORGANIZATION_MAP[part]
                    break
            
            # Extract date
            date = extract_date_from_filename(filename)
            
            # Extract or improve title
            if not doc.get('title') or doc.get('title') == '제목 없음':
                title = extract_title_from_filename(filename)
            else:
                title = doc.get('title')
            
            # Prepare update data
            update_data = {}
            
            if organization and (not doc.get('organization') or doc.get('organization') == '조직 미상'):
                update_data['organization'] = organization
                
            if date and not doc.get('created_at'):
                update_data['created_at'] = date
                
            if title and title != doc.get('title'):
                update_data['title'] = title
            
            # Update metadata field
            metadata = doc.get('metadata', {})
            if organization:
                metadata['organization'] = organization
            if date:
                metadata['created_at'] = date
            if title:
                metadata['title'] = title
                
            update_data['metadata'] = metadata
            
            # Perform update if there are changes
            if update_data:
                print(f"Updating: {filename}")
                print(f"  Organization: {update_data.get('organization', 'no change')}")
                print(f"  Date: {update_data.get('created_at', 'no change')}")
                print(f"  Title: {update_data.get('title', 'no change')}")
                
                # Update in database
                client.client.table('documents').update(update_data).eq('id', doc['id']).execute()
                updated_count += 1
                print("  [OK] Updated\n")
                
        except Exception as e:
            print(f"Error updating document {doc.get('id')}: {str(e)}\n")
    
    print(f"\n[COMPLETE] Updated {updated_count} documents")

if __name__ == "__main__":
    print("=== Update Existing Document Metadata ===\n")
    update_document_metadata()