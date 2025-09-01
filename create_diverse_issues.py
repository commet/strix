"""
다양한 테스트 이슈 생성 스크립트
20개 이상의 현실적인 배터리 산업 이슈 생성
"""
import os
import sys
from datetime import datetime, timedelta
import random
import uuid

sys.path.append(os.path.dirname(__file__))
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
from src.database.supabase_client import SupabaseClient

def create_diverse_issues():
    """다양한 테스트 이슈 데이터 생성"""
    supabase = SupabaseClient()
    
    # 먼저 기존 이슈 삭제 (clean slate)
    try:
        supabase.client.table('issues').delete().neq('id', '00000000-0000-0000-0000-000000000000').execute()
        print("기존 이슈 삭제 완료")
    except:
        pass
    
    # 다양한 이슈 데이터
    test_issues = [
        # 기술 카테고리
        {
            'issue_key': 'ISS-2025-001',
            'title': '전고체 배터리 양산 기술 개발 - 2025년 목표',
            'category': '기술',
            'priority': 'CRITICAL',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=120)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=2)).date().isoformat(),
            'department': 'R&D센터',
            'owner': '김철수',
            'description': '2025년 하반기 파일럿 생산을 목표로 전고체 배터리 기술 개발',
        },
        {
            'issue_key': 'ISS-2025-002',
            'title': '46파이 원통형 배터리 개발 프로젝트',
            'category': '기술',
            'priority': 'HIGH',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=90)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=5)).date().isoformat(),
            'department': 'R&D센터',
            'owner': '박영희',
            'description': '테슬라 요구사항에 맞춘 46파이 배터리 개발',
        },
        {
            'issue_key': 'ISS-2025-003',
            'title': 'LFP 배터리 기술 도입 검토',
            'category': '기술',
            'priority': 'MEDIUM',
            'status': 'OPEN',
            'first_mentioned_date': (datetime.now() - timedelta(days=30)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=1)).date().isoformat(),
            'department': '기술전략팀',
            'owner': '이준호',
            'description': '중국 경쟁사 대응을 위한 LFP 기술 도입 타당성 검토',
        },
        {
            'issue_key': 'ISS-2025-004',
            'title': '실리콘 음극재 적용 연구',
            'category': '기술',
            'priority': 'HIGH',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=60)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=7)).date().isoformat(),
            'department': 'R&D센터',
            'owner': '정민수',
            'description': '에너지 밀도 향상을 위한 실리콘 음극재 적용',
        },
        {
            'issue_key': 'ISS-2025-005',
            'title': '급속충전 기술 개발 (10분 충전 80%)',
            'category': '기술',
            'priority': 'CRITICAL',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=45)).date().isoformat(),
            'last_updated': datetime.now().date().isoformat(),
            'department': 'R&D센터',
            'owner': '김서연',
            'description': 'BYD 5분 충전 기술 대응',
        },
        
        # 리스크 카테고리
        {
            'issue_key': 'ISS-2025-006',
            'title': '리튬 가격 변동성 리스크 관리',
            'category': '리스크',
            'priority': 'HIGH',
            'status': 'MONITORING',
            'first_mentioned_date': (datetime.now() - timedelta(days=150)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=3)).date().isoformat(),
            'department': '경영지원실',
            'owner': '이영희',
            'description': '원자재 가격 급등에 따른 수익성 악화 리스크',
        },
        {
            'issue_key': 'ISS-2025-007',
            'title': '전기차 캐즘 심화로 인한 수요 급감',
            'category': '리스크',
            'priority': 'CRITICAL',
            'status': 'MONITORING',
            'first_mentioned_date': (datetime.now() - timedelta(days=60)).date().isoformat(),
            'last_updated': datetime.now().date().isoformat(),
            'department': '전략기획팀',
            'owner': '최현우',
            'description': '글로벌 전기차 판매 둔화 대응',
        },
        {
            'issue_key': 'ISS-2025-008',
            'title': '화재 사고 리스크 및 리콜 대응',
            'category': '리스크',
            'priority': 'CRITICAL',
            'status': 'RESOLVED',
            'first_mentioned_date': (datetime.now() - timedelta(days=180)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=30)).date().isoformat(),
            'resolution_date': (datetime.now() - timedelta(days=30)).date().isoformat(),
            'department': '품질관리팀',
            'owner': '박지훈',
            'description': '배터리 화재 사고 예방 시스템 구축 완료',
        },
        {
            'issue_key': 'ISS-2025-009',
            'title': '공급망 다변화 필요성',
            'category': '리스크',
            'priority': 'HIGH',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=75)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=4)).date().isoformat(),
            'department': '구매팀',
            'owner': '강민정',
            'description': '중국 의존도 감소 및 공급망 안정화',
        },
        
        # 경쟁사 카테고리
        {
            'issue_key': 'ISS-2025-010',
            'title': 'CATL 유럽시장 점유율 37.9% 대응',
            'category': '경쟁사',
            'priority': 'HIGH',
            'status': 'MONITORING',
            'first_mentioned_date': (datetime.now() - timedelta(days=90)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=2)).date().isoformat(),
            'department': '전략기획팀',
            'owner': '박민수',
            'description': 'CATL 유럽 공장 가동에 따른 대응 전략',
        },
        {
            'issue_key': 'ISS-2025-011',
            'title': 'BYD 수직계열화 모델 분석',
            'category': '경쟁사',
            'priority': 'MEDIUM',
            'status': 'RESOLVED',
            'first_mentioned_date': (datetime.now() - timedelta(days=120)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=15)).date().isoformat(),
            'resolution_date': (datetime.now() - timedelta(days=15)).date().isoformat(),
            'department': '시장분석팀',
            'owner': '윤서진',
            'description': 'BYD 비즈니스 모델 벤치마킹 완료',
        },
        {
            'issue_key': 'ISS-2025-012',
            'title': 'LG에너지솔루션 위기경영 선언 모니터링',
            'category': '경쟁사',
            'priority': 'MEDIUM',
            'status': 'MONITORING',
            'first_mentioned_date': (datetime.now() - timedelta(days=20)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=1)).date().isoformat(),
            'department': '시장분석팀',
            'owner': '김태현',
            'description': '국내 경쟁사 동향 파악 및 시사점 도출',
        },
        {
            'issue_key': 'ISS-2025-013',
            'title': '테슬라 자체 배터리 생산 확대',
            'category': '경쟁사',
            'priority': 'HIGH',
            'status': 'OPEN',
            'first_mentioned_date': (datetime.now() - timedelta(days=40)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=3)).date().isoformat(),
            'department': '영업팀',
            'owner': '이상호',
            'description': '주요 고객사의 내재화 전략 대응',
        },
        
        # 전략 카테고리
        {
            'issue_key': 'ISS-2025-014',
            'title': 'SK온-SK이노베이션 합병 PMI',
            'category': '전략',
            'priority': 'CRITICAL',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=10)).date().isoformat(),
            'last_updated': datetime.now().date().isoformat(),
            'department': '경영기획팀',
            'owner': '정수진',
            'description': '11월 1일 통합법인 출범 준비',
        },
        {
            'issue_key': 'ISS-2025-015',
            'title': '2030년 EBITDA 20조원 달성 로드맵',
            'category': '전략',
            'priority': 'HIGH',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=5)).date().isoformat(),
            'last_updated': datetime.now().date().isoformat(),
            'department': '전략기획팀',
            'owner': '한지민',
            'description': '중장기 성장 전략 수립',
        },
        {
            'issue_key': 'ISS-2025-016',
            'title': '북미 제2공장 투자 결정',
            'category': '전략',
            'priority': 'HIGH',
            'status': 'OPEN',
            'first_mentioned_date': (datetime.now() - timedelta(days=35)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=2)).date().isoformat(),
            'department': '해외사업팀',
            'owner': '조현준',
            'description': '조지아 제2공장 증설 타당성 검토',
        },
        {
            'issue_key': 'ISS-2025-017',
            'title': 'ESG 평가등급 A등급 달성',
            'category': '전략',
            'priority': 'MEDIUM',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=100)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=10)).date().isoformat(),
            'department': 'ESG팀',
            'owner': '박소연',
            'description': '2025년 ESG 평가 대응',
        },
        {
            'issue_key': 'ISS-2025-018',
            'title': '5조원 규모 자본확충 진행',
            'category': '전략',
            'priority': 'CRITICAL',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=15)).date().isoformat(),
            'last_updated': datetime.now().date().isoformat(),
            'department': '재무팀',
            'owner': '김동현',
            'description': '유상증자 및 전략적 투자자 유치',
        },
        
        # 정책 카테고리
        {
            'issue_key': 'ISS-2025-019',
            'title': '트럼프 2기 IRA 폐지 가능성 대응',
            'category': '정책',
            'priority': 'CRITICAL',
            'status': 'MONITORING',
            'first_mentioned_date': (datetime.now() - timedelta(days=25)).date().isoformat(),
            'last_updated': datetime.now().date().isoformat(),
            'department': '정책대응팀',
            'owner': '최준호',
            'description': 'AMPC 세액공제 축소/폐지 시나리오 대응',
        },
        {
            'issue_key': 'ISS-2025-020',
            'title': 'EU 배터리 규제(CBAM) 대응',
            'category': '정책',
            'priority': 'HIGH',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=85)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=6)).date().isoformat(),
            'department': '규제대응팀',
            'owner': '이수진',
            'description': '탄소국경조정메커니즘 대응 준비',
        },
        {
            'issue_key': 'ISS-2025-021',
            'title': '중국 배터리 백서 규제',
            'category': '정책',
            'priority': 'MEDIUM',
            'status': 'RESOLVED',
            'first_mentioned_date': (datetime.now() - timedelta(days=200)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=45)).date().isoformat(),
            'resolution_date': (datetime.now() - timedelta(days=45)).date().isoformat(),
            'department': '중국사업팀',
            'owner': '장현석',
            'description': '중국 현지 인증 획득 완료',
        },
        {
            'issue_key': 'ISS-2025-022',
            'title': '배터리 여권(Battery Passport) 도입',
            'category': '정책',
            'priority': 'HIGH',
            'status': 'IN_PROGRESS',
            'first_mentioned_date': (datetime.now() - timedelta(days=50)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=8)).date().isoformat(),
            'department': 'IT혁신팀',
            'owner': '백승민',
            'description': '2026년 EU 배터리 여권 의무화 대응',
        },
        {
            'issue_key': 'ISS-2025-023',
            'title': '폐배터리 재활용 의무화',
            'category': '정책',
            'priority': 'MEDIUM',
            'status': 'OPEN',
            'first_mentioned_date': (datetime.now() - timedelta(days=70)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=12)).date().isoformat(),
            'department': '환경안전팀',
            'owner': '송미래',
            'description': '폐배터리 재활용 체계 구축',
        },
        
        # 추가 이슈들
        {
            'issue_key': 'ISS-2025-024',
            'title': '인도시장 진출 전략 수립',
            'category': '전략',
            'priority': 'MEDIUM',
            'status': 'OPEN',
            'first_mentioned_date': (datetime.now() - timedelta(days=28)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=3)).date().isoformat(),
            'department': '신시장개발팀',
            'owner': '나현우',
            'description': '인도 전기차 시장 진출 방안',
        },
        {
            'issue_key': 'ISS-2025-025',
            'title': 'AI 기반 품질예측 시스템 구축',
            'category': '기술',
            'priority': 'MEDIUM',
            'status': 'RESOLVED',
            'first_mentioned_date': (datetime.now() - timedelta(days=150)).date().isoformat(),
            'last_updated': (datetime.now() - timedelta(days=20)).date().isoformat(),
            'resolution_date': (datetime.now() - timedelta(days=20)).date().isoformat(),
            'department': 'DX추진팀',
            'owner': '오지훈',
            'description': 'AI/ML 활용 불량률 예측 시스템 구축 완료',
        },
    ]
    
    print(f"총 {len(test_issues)}개 이슈 생성 시작...")
    
    for idx, issue in enumerate(test_issues, 1):
        try:
            # 상태별 메타데이터 추가
            metadata = {
                'progress_percentage': random.randint(10, 95) if issue['status'] == 'IN_PROGRESS' else (100 if issue['status'] == 'RESOLVED' else 0),
                'risk_level': random.choice(['LOW', 'MEDIUM', 'HIGH', 'CRITICAL']),
                'impact_score': random.randint(1, 10),
                'urgency_score': random.randint(1, 10),
            }
            issue['metadata'] = metadata
            
            # 이슈 삽입
            result = supabase.client.table('issues').upsert(issue).execute()
            print(f"[{idx}/{len(test_issues)}] OK: {issue['title'][:30]}...")
            
        except Exception as e:
            print(f"[{idx}/{len(test_issues)}] FAIL: {str(e)}")
    
    print(f"\n테스트 이슈 생성 완료!")
    
    # 생성된 이슈 확인
    try:
        issues = supabase.client.table('issues').select('*').execute()
        print(f"총 {len(issues.data)}개의 이슈가 데이터베이스에 있습니다.")
        
        # 상태별 집계
        status_count = {}
        for issue in issues.data:
            status = issue['status']
            status_count[status] = status_count.get(status, 0) + 1
        
        print("\n상태별 분포:")
        for status, count in status_count.items():
            print(f"  - {status}: {count}개")
            
    except Exception as e:
        print(f"이슈 확인 실패: {str(e)}")

if __name__ == "__main__":
    create_diverse_issues()