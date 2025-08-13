"""
추가 테스트 이슈 생성
"""
import os
import sys
from datetime import datetime, timedelta
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
from database.supabase_client import SupabaseClient

# Supabase 클라이언트
supabase = SupabaseClient()

# 추가 테스트 이슈들
additional_issues = [
    {
        'issue_key': 'ISS-2024-002',
        'title': '원자재 가격 변동성 리스크',
        'category': '리스크',
        'priority': 'HIGH',
        'status': 'MONITORING',
        'first_mentioned_date': '2024-01-25',
        'last_updated': datetime.now().date().isoformat(),
        'department': '경영지원',
        'owner': '이영희',
        'description': '리튬 가격 급등으로 수익성 악화 우려'
    },
    {
        'issue_key': 'ISS-2024-003',
        'title': 'CATL 유럽 공장 대응 전략',
        'category': '경쟁사',
        'priority': 'MEDIUM',
        'status': 'RESOLVED',
        'first_mentioned_date': '2024-01-10',
        'last_updated': datetime.now().date().isoformat(),
        'department': '전략기획',
        'owner': '박민수',
        'description': 'CATL 유럽 진출 대응 전략 수립'
    },
    {
        'issue_key': 'ISS-2024-004',
        'title': 'ESG 평가 등급 개선',
        'category': '전략',
        'priority': 'MEDIUM',
        'status': 'IN_PROGRESS',
        'first_mentioned_date': '2024-02-01',
        'last_updated': datetime.now().date().isoformat(),
        'department': 'ESG팀',
        'owner': '정수진',
        'description': '2024년 ESG A등급 목표'
    },
    {
        'issue_key': 'ISS-2024-005',
        'title': '배터리 안전성 규제 대응',
        'category': '정책',
        'priority': 'HIGH',
        'status': 'OPEN',
        'first_mentioned_date': '2024-02-15',
        'last_updated': datetime.now().date().isoformat(),
        'department': '품질관리',
        'owner': '최준호',
        'description': '새로운 안전성 규제 대응'
    }
]

print("추가 이슈 생성 중...")

for issue in additional_issues:
    try:
        # Upsert (있으면 업데이트, 없으면 생성)
        result = supabase.client.table('issues').upsert(
            issue,
            on_conflict='issue_key'
        ).execute()
        print(f"[OK] {issue['title']}")
    except Exception as e:
        print(f"[ERROR] {issue['title']}: {str(e)}")

# 확인
try:
    all_issues = supabase.client.table('issues').select('title, status').execute()
    print(f"\n총 {len(all_issues.data)}개 이슈:")
    for issue in all_issues.data:
        print(f"  - {issue['title']} [{issue['status']}]")
except Exception as e:
    print(f"확인 실패: {str(e)}")