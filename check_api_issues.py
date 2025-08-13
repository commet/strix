"""
API와 데이터베이스 이슈 확인
"""
import requests
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
from database.supabase_client import SupabaseClient

print("=== 데이터베이스 직접 확인 ===")
supabase = SupabaseClient()

# 데이터베이스에서 직접 조회
try:
    db_issues = supabase.client.table('issues').select('*').execute()
    print(f"DB에 있는 이슈: {len(db_issues.data)}개")
    for issue in db_issues.data:
        print(f"  - {issue['title']} [{issue['status']}]")
except Exception as e:
    print(f"DB 조회 실패: {e}")

print("\n=== API 응답 확인 ===")
# API 호출
try:
    response = requests.get('http://localhost:5000/api/issues')
    if response.status_code == 200:
        api_issues = response.json()
        print(f"API가 반환한 이슈: {len(api_issues)}개")
        for issue in api_issues:
            print(f"  - {issue.get('title', 'N/A')} [{issue.get('status', 'N/A')}]")
    else:
        print(f"API 오류: {response.status_code}")
except Exception as e:
    print(f"API 호출 실패: {e}")

print("\n=== issue_summary 뷰 확인 ===")
# issue_summary 뷰 확인 (API가 사용하는 뷰)
try:
    view_issues = supabase.client.table('issue_summary').select('*').execute()
    print(f"issue_summary 뷰의 이슈: {len(view_issues.data)}개")
    for issue in view_issues.data:
        print(f"  - {issue.get('title', 'N/A')} [{issue.get('status', 'N/A')}]")
except Exception as e:
    print(f"뷰 조회 실패: {e}")
    print("issue_summary 뷰가 없을 수 있습니다.")