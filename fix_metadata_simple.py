"""
Simple fix for document metadata
"""
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from dotenv import load_dotenv

load_dotenv()

# Define proper metadata for known documents
DOCUMENT_METADATA = {
    "전고체": {
        "title": "전고체 배터리 개발 현황 보고",
        "organization": "R&D",
        "created_at": "2024-01-20"
    },
    "중장기전략": {
        "title": "2024년 1분기 배터리 사업 중장기 전략 보고서",
        "organization": "전략기획",
        "created_at": "2024-01-15"
    },
    "리스크": {
        "title": "2024년 1분기 리스크 관리 현황",
        "organization": "경영지원",
        "created_at": "2024-01-25"
    },
    "CATL": {
        "title": "[경쟁사] CATL 유럽 신공장 착공",
        "organization": "Google Alert",
        "created_at": "2024-01-15"
    },
    "배터리시장동향": {
        "title": "[산업] 글로벌 배터리 수요 급증",
        "organization": "PR팀_AM",
        "created_at": "2024-01-16"
    },
    "정책브리핑분석": {
        "title": "[정책] 정부 K-배터리 지원책 발표",
        "organization": "PR팀_PM",
        "created_at": "2024-01-17"
    }
}

def fix_metadata():
    client = SupabaseClient()
    
    # Get all documents
    print("Fetching documents...")
    response = client.client.table('documents').select('*').execute()
    documents = response.data
    
    print(f"Found {len(documents)} documents\n")
    
    for doc in documents:
        try:
            # Find matching metadata
            file_path = doc.get('file_path', '')
            content_snippet = doc.get('title', '') + ' ' + str(doc.get('metadata', {}))
            
            matched = False
            for keyword, metadata in DOCUMENT_METADATA.items():
                if keyword in file_path or keyword in content_snippet:
                    print(f"Updating document: {doc['id'][:8]}...")
                    print(f"  Title: {metadata['title']}")
                    print(f"  Organization: {metadata['organization']}")
                    print(f"  Date: {metadata['created_at']}")
                    
                    # Update document
                    update_data = {
                        'title': metadata['title'],
                        'organization': metadata['organization'],
                        'created_at': metadata['created_at'],
                        'metadata': metadata
                    }
                    
                    client.client.table('documents').update(update_data).eq('id', doc['id']).execute()
                    print("  [OK] Updated\n")
                    matched = True
                    break
            
            if not matched:
                print(f"No match for document: {doc['id'][:8]}")
                print(f"  Current title: {doc.get('title', 'N/A')}")
                print(f"  File path: {file_path}\n")
                
        except Exception as e:
            print(f"Error: {str(e)}\n")
    
    print("[COMPLETE] Metadata fix completed")

if __name__ == "__main__":
    fix_metadata()