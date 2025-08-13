"""
테스트용 이슈 데이터 생성
"""
import os
import sys
from datetime import datetime, timedelta
import uuid

sys.path.append(os.path.dirname(__file__))
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
from database.supabase_client import SupabaseClient

def create_test_issues():
    """테스트용 이슈 데이터 생성"""
    supabase = SupabaseClient()
    
    # 테스트 이슈 데이터
    test_issues = [
        {
            'issue_key': f'ISS-2024-001',
            'title': '전고체 배터리 양산 기술 개발',
            'category': '기술',
            'priority': 'HIGH',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=60)).date().isoformat(),
            'last_updated': datetime.now().date().isoformat(),
            'department': 'R&D',
            'owner': '김철수',
            'description': '2024년 하반기 파일럿 생산을 목표로 전고체 배터리 기술 개발 진행',
            'metadata': {'keywords': ['전고체', '배터리', 'R&D']}
        },
        {
            'issue_key': f'ISS-2024-002',
            'title': '원자재 가격 변동성 리스크 관리',
            'category': '리스크',
            'priority': 'HIGH',
            'status': 'MONITORING',
            'first_mentioned_date': (datetime.now() - timedelta(days=45)).date().isoformat(),
            'last_updated': datetime.now().date().isoformat(),
            'department': '경영지원',
            'owner': '이영희',
            'description': '리튬 및 코발트 가격 급등에 따른 수익성 악화 리스크 모니터링',
            'metadata': {'keywords': ['원자재', '리스크', '가격']}
        },
        {
            'issue_key': f'ISS-2024-003',
            'title': 'CATL 유럽 공장 대응 전략 수립',
            'category': '경쟁사',
            'priority': 'MEDIUM',
            'status': 'RESOLVED',
            'first_mentioned_date': (datetime.now() - timedelta(days=75)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=10)).date().isoformat(),
            'resolution_date': (datetime.now() - timedelta(days=10)).date().isoformat(),
            'department': '전략기획',
            'owner': '박민수',
            'description': 'CATL 유럽 진출에 대응한 현지 파트너십 전략 수립 완료',
            'metadata': {'keywords': ['CATL', '경쟁사', '유럽']}
        },
        {
            'issue_key': f'ISS-2024-004',
            'title': 'ESG 평가 등급 개선 프로젝트',
            'category': '전략',
            'priority': 'MEDIUM',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=30)).date().isoformat(),
            'last_updated': datetime.now().date().isoformat(),
            'department': 'ESG팀',
            'owner': '정수진',
            'description': '2024년 ESG 평가 A등급 달성을 위한 개선 프로젝트',
            'metadata': {'keywords': ['ESG', '지속가능성', '평가']}
        },
        {
            'issue_key': f'ISS-2024-005',
            'title': '배터리 안전성 규제 대응',
            'category': '정책',
            'priority': 'HIGH',
            'status': 'OPEN',
            'first_mentioned_date': (datetime.now() - timedelta(days=15)).date().isoformat(),
            'last_updated': datetime.now().date().isoformat(),
            'department': '품질관리',
            'owner': '최준호',
            'description': '새로운 배터리 안전성 규제 기준 대응 방안 수립 필요',
            'metadata': {'keywords': ['규제', '안전성', '인증']}
        }
    ]
    
    print("테스트 이슈 생성 중...")
    
    for issue in test_issues:
        try:
            # 이슈 삽입
            result = supabase.client.table('issues').upsert(issue).execute()
            print(f"[OK] 이슈 생성: {issue['title']}")
        except Exception as e:
            print(f"[ERROR] 이슈 생성 실패: {issue['title']} - {str(e)}")
    
    print("\n테스트 이슈 생성 완료!")
    
    # 생성된 이슈 확인
    try:
        issues = supabase.client.table('issues').select('*').execute()
        print(f"\n총 {len(issues.data)}개의 이슈가 데이터베이스에 있습니다.")
    except Exception as e:
        print(f"이슈 확인 실패: {str(e)}")

if __name__ == "__main__":
    create_test_issues()