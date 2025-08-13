"""
Issue Tracker 데이터베이스 설정 스크립트
"""
import os
from supabase import create_client, Client
from dotenv import load_dotenv

load_dotenv()

# Supabase 클라이언트 생성
supabase: Client = create_client(
    os.getenv("SUPABASE_URL"),
    os.getenv("SUPABASE_KEY")
)

print("=== STRIX Issue Tracker 설정 ===\n")
print("1. Supabase Dashboard에 로그인하세요")
print("2. SQL Editor로 이동하세요")
print("3. create_issue_tracking_tables.sql 파일의 내용을 실행하세요")
print("\n또는 아래 명령어를 실행하세요:")
print("\n테이블이 생성되었는지 확인:")

# 테이블 존재 확인
try:
    result = supabase.table('issues').select('*').limit(1).execute()
    print("✓ issues 테이블이 이미 존재합니다")
except:
    print("✗ issues 테이블이 없습니다. SQL 스크립트를 실행해주세요")

print("\n=== 사용 방법 ===")
print("\n1. API 서버 실행:")
print("   python api_server_with_issues.py")
print("\n2. Excel에서 Issue Timeline 대시보드 생성:")
print("   - Excel 파일 열기")
print("   - Alt+F11로 VBA 편집기 열기")
print("   - modIssueTimeline 모듈에서 CreateIssueTimelineDashboard 실행")
print("\n3. 기능 사용:")
print("   - 문서에서 이슈 자동 추출")
print("   - 이슈 타임라인 시각화")
print("   - AI 기반 예측 및 분석")
print("   - 상태별/카테고리별 필터링")