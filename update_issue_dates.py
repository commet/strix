"""
이슈 날짜를 최근으로 업데이트
"""
import sys
import os
from datetime import datetime, timedelta
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
from database.supabase_client import SupabaseClient

supabase = SupabaseClient()

# 모든 이슈의 날짜를 최근으로 업데이트
updates = [
    ('ISS-2024-001', (datetime.now() - timedelta(days=60)).date().isoformat()),
    ('ISS-2024-002', (datetime.now() - timedelta(days=45)).date().isoformat()),
    ('ISS-2024-003', (datetime.now() - timedelta(days=75)).date().isoformat()),
    ('ISS-2024-004', (datetime.now() - timedelta(days=30)).date().isoformat()),
    ('ISS-2024-005', (datetime.now() - timedelta(days=15)).date().isoformat()),
]

print("이슈 날짜 업데이트 중...")

for issue_key, new_date in updates:
    try:
        result = supabase.client.table('issues')\
            .update({
                'first_mentioned_date': new_date,
                'last_updated': datetime.now().date().isoformat()
            })\
            .eq('issue_key', issue_key)\
            .execute()
        print(f"[OK] {issue_key}: {new_date}")
    except Exception as e:
        print(f"[ERROR] {issue_key}: {str(e)}")

# 확인
print("\n업데이트된 이슈 확인:")
issues = supabase.client.table('issues').select('title, first_mentioned_date').execute()
for issue in issues.data:
    print(f"  - {issue['title']}: {issue['first_mentioned_date']}")