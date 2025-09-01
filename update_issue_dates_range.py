"""
이슈 날짜를 2025년 6월-10월 범위로 업데이트
"""
import os
import sys
from datetime import datetime, timedelta
import random

sys.path.append(os.path.dirname(__file__))
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
from src.database.supabase_client import SupabaseClient

def update_issue_dates():
    """이슈 날짜를 2025년 6-10월 범위로 업데이트"""
    supabase = SupabaseClient()
    
    # 모든 이슈 가져오기
    result = supabase.client.table('issues').select('*').execute()
    issues = result.data
    
    print(f"총 {len(issues)}개 이슈 날짜 업데이트 시작...")
    
    # 날짜 범위 설정 (2025년 6월 1일 ~ 10월 31일)
    start_date = datetime(2025, 6, 1)
    end_date = datetime(2025, 10, 31)
    
    for idx, issue in enumerate(issues, 1):
        # 상태별로 날짜 설정
        if issue['status'] == 'RESOLVED':
            # 해결된 이슈: 6-7월에 시작, 7-8월에 종료
            first_date = start_date + timedelta(days=random.randint(0, 60))
            last_date = first_date + timedelta(days=random.randint(15, 45))
        elif issue['status'] == 'IN_PROGRESS':
            # 진행중: 6-8월에 시작, 아직 진행중
            first_date = start_date + timedelta(days=random.randint(0, 90))
            last_date = datetime(2025, 8, 27)  # 현재
        elif issue['status'] == 'MONITORING':
            # 모니터링: 6-7월에 시작, 계속 모니터링중
            first_date = start_date + timedelta(days=random.randint(0, 60))
            last_date = datetime(2025, 8, 27)  # 현재
        else:  # OPEN
            # 미해결: 7-8월에 시작
            first_date = start_date + timedelta(days=random.randint(30, 90))
            last_date = datetime(2025, 8, 27)  # 현재
        
        # 날짜 업데이트
        update_data = {
            'first_mentioned_date': first_date.date().isoformat(),
            'last_updated': last_date.date().isoformat()
        }
        
        # RESOLVED 상태면 resolution_date도 추가
        if issue['status'] == 'RESOLVED':
            update_data['resolution_date'] = last_date.date().isoformat()
        
        try:
            result = supabase.client.table('issues')\
                .update(update_data)\
                .eq('id', issue['id'])\
                .execute()
            
            print(f"[{idx}/{len(issues)}] {issue['issue_key']}: {first_date.strftime('%Y-%m-%d')} ~ {last_date.strftime('%Y-%m-%d')}")
        except Exception as e:
            print(f"[{idx}/{len(issues)}] 실패: {str(e)}")
    
    print("\n날짜 업데이트 완료!")
    
    # 업데이트된 이슈 확인
    result = supabase.client.table('issues')\
        .select('issue_key, status, first_mentioned_date, last_updated')\
        .order('first_mentioned_date')\
        .execute()
    
    print("\n업데이트된 날짜 확인:")
    for issue in result.data[:5]:  # 처음 5개만 출력
        print(f"  {issue['issue_key']}: {issue['first_mentioned_date']} ({issue['status']})")

if __name__ == "__main__":
    update_issue_dates()