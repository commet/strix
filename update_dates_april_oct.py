"""
이슈 날짜를 2025년 4월-10월 범위로 재분산
"""
import os
import sys
from datetime import datetime, timedelta
import random

sys.path.append(os.path.dirname(__file__))
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
from src.database.supabase_client import SupabaseClient

def update_issue_dates_wide_range():
    """이슈 날짜를 2025년 4-10월 범위로 분산"""
    supabase = SupabaseClient()
    
    # 모든 이슈 가져오기
    result = supabase.client.table('issues').select('*').execute()
    issues = result.data
    
    print(f"총 {len(issues)}개 이슈 날짜 업데이트 시작...")
    
    # 날짜 범위 설정 (2025년 4월 1일 ~ 10월 31일)
    start_date = datetime(2025, 4, 1)
    end_date = datetime(2025, 10, 31)
    today = datetime(2025, 8, 27)  # 현재 날짜 고정
    
    # 이슈를 상태별로 분류
    resolved_issues = [i for i in issues if i['status'] == 'RESOLVED']
    in_progress_issues = [i for i in issues if i['status'] == 'IN_PROGRESS']
    monitoring_issues = [i for i in issues if i['status'] == 'MONITORING']
    open_issues = [i for i in issues if i['status'] == 'OPEN']
    
    # RESOLVED: 4-6월 시작, 5-7월 종료
    for idx, issue in enumerate(resolved_issues, 1):
        first_date = start_date + timedelta(days=random.randint(0, 60))  # 4-5월
        last_date = first_date + timedelta(days=random.randint(20, 60))  # 1-2개월 후 종료
        
        update_data = {
            'first_mentioned_date': first_date.date().isoformat(),
            'last_updated': last_date.date().isoformat(),
            'resolution_date': last_date.date().isoformat()
        }
        
        supabase.client.table('issues').update(update_data).eq('id', issue['id']).execute()
        print(f"RESOLVED {idx}/{len(resolved_issues)}: {issue['issue_key']} -> {first_date.strftime('%m/%d')} ~ {last_date.strftime('%m/%d')}")
    
    # IN_PROGRESS: 5-7월 시작, 현재까지
    for idx, issue in enumerate(in_progress_issues, 1):
        first_date = start_date + timedelta(days=random.randint(30, 120))  # 5-7월
        
        update_data = {
            'first_mentioned_date': first_date.date().isoformat(),
            'last_updated': today.date().isoformat()
        }
        
        supabase.client.table('issues').update(update_data).eq('id', issue['id']).execute()
        print(f"IN_PROGRESS {idx}/{len(in_progress_issues)}: {issue['issue_key']} -> {first_date.strftime('%m/%d')} ~ 진행중")
    
    # MONITORING: 4-6월 시작, 계속 모니터링
    for idx, issue in enumerate(monitoring_issues, 1):
        first_date = start_date + timedelta(days=random.randint(0, 90))  # 4-6월
        
        update_data = {
            'first_mentioned_date': first_date.date().isoformat(),
            'last_updated': today.date().isoformat()
        }
        
        supabase.client.table('issues').update(update_data).eq('id', issue['id']).execute()
        print(f"MONITORING {idx}/{len(monitoring_issues)}: {issue['issue_key']} -> {first_date.strftime('%m/%d')} ~ 모니터링")
    
    # OPEN: 7-8월 시작
    for idx, issue in enumerate(open_issues, 1):
        first_date = start_date + timedelta(days=random.randint(90, 145))  # 7-8월
        
        update_data = {
            'first_mentioned_date': first_date.date().isoformat(),
            'last_updated': today.date().isoformat()
        }
        
        supabase.client.table('issues').update(update_data).eq('id', issue['id']).execute()
        print(f"OPEN {idx}/{len(open_issues)}: {issue['issue_key']} -> {first_date.strftime('%m/%d')} ~ 미해결")
    
    print("\n날짜 업데이트 완료!")
    
    # 분포 확인
    result = supabase.client.table('issues').select('first_mentioned_date, status').execute()
    month_counts = {}
    for issue in result.data:
        month = issue['first_mentioned_date'][:7]
        month_counts[month] = month_counts.get(month, 0) + 1
    
    print("\n월별 이슈 분포:")
    for month in sorted(month_counts.keys()):
        print(f"  {month}: {month_counts[month]}개")

if __name__ == "__main__":
    update_issue_dates_wide_range()