"""
Issue Extraction 테스트
"""
import requests
import json

# API 서버 URL
API_URL = "http://localhost:5000"

print("=== STRIX Issue Tracker API 테스트 ===\n")

# 1. 헬스 체크
print("1. API 서버 상태 확인...")
try:
    response = requests.get(f"{API_URL}/health")
    if response.status_code == 200:
        print(f"[OK] 서버 상태: {response.json()}\n")
    else:
        print(f"[ERROR] 서버 응답 오류: {response.status_code}\n")
except Exception as e:
    print(f"[ERROR] 서버 연결 실패: {e}\n")
    exit(1)

# 2. 이슈 목록 조회
print("2. 이슈 목록 조회...")
try:
    response = requests.get(f"{API_URL}/api/issues")
    if response.status_code == 200:
        issues = response.json()
        print(f"[OK] 총 {len(issues)}개의 이슈 발견")
        
        # 처음 3개 이슈만 표시
        for issue in issues[:3]:
            print(f"  - {issue.get('title', 'N/A')} [{issue.get('status', 'N/A')}]")
        print()
    else:
        print(f"[ERROR] 이슈 조회 실패: {response.status_code}\n")
except Exception as e:
    print(f"[ERROR] 오류: {e}\n")

# 3. 대시보드 요약 조회
print("3. 대시보드 요약 정보...")
try:
    response = requests.get(f"{API_URL}/api/issues/dashboard-summary")
    if response.status_code == 200:
        summary = response.json()
        print(f"[OK] 대시보드 데이터 수신")
        
        stats = summary.get('statistics', {})
        print(f"  - 전체 이슈: {stats.get('total_issues', 0)}")
        print(f"  - 미해결: {stats.get('open_issues', 0)}")
        print(f"  - 진행중: {stats.get('in_progress', 0)}")
        print(f"  - 해결됨: {stats.get('resolved', 0)}")
        print()
    else:
        print(f"[ERROR] 대시보드 조회 실패: {response.status_code}\n")
except Exception as e:
    print(f"[ERROR] 오류: {e}\n")

# 4. AI 예측 테스트
print("4. AI 예측 생성 테스트...")
try:
    response = requests.post(f"{API_URL}/api/issues/predict")
    if response.status_code == 200:
        result = response.json()
        print(f"[OK] AI 예측 생성 완료: {result}\n")
    else:
        print(f"[ERROR] AI 예측 실패: {response.status_code}\n")
except Exception as e:
    print(f"[ERROR] 오류: {e}\n")

print("=== 테스트 완료 ===")
print("\nExcel에서 Issue Timeline Dashboard를 생성하여")
print("실제 UI에서 이슈를 관리할 수 있습니다.")