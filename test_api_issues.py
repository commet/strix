"""
API 이슈 조회 테스트
"""
import requests
import json

# API 테스트
print("=== API 이슈 조회 테스트 ===\n")

# 1. 이슈 목록 조회
url = "http://localhost:5000/api/issues"
print(f"요청 URL: {url}\n")

try:
    response = requests.get(url)
    print(f"응답 상태: {response.status_code}")
    
    if response.status_code == 200:
        data = response.json()
        print(f"응답 데이터 타입: {type(data)}")
        print(f"데이터 길이: {len(data)}")
        
        if len(data) > 0:
            print("\n첫 번째 이슈:")
            first_issue = data[0] if isinstance(data, list) else data
            print(json.dumps(first_issue, indent=2, ensure_ascii=False))
        else:
            print("\n이슈 데이터가 비어있습니다.")
            print("전체 응답:", response.text[:500])
    else:
        print(f"오류 응답: {response.text}")
        
except Exception as e:
    print(f"요청 실패: {e}")

# 2. 대시보드 요약 조회
print("\n\n=== 대시보드 요약 조회 ===\n")
url = "http://localhost:5000/api/issues/dashboard-summary"
print(f"요청 URL: {url}\n")

try:
    response = requests.get(url)
    print(f"응답 상태: {response.status_code}")
    
    if response.status_code == 200:
        data = response.json()
        print("응답 데이터:")
        print(json.dumps(data, indent=2, ensure_ascii=False))
    else:
        print(f"오류 응답: {response.text}")
        
except Exception as e:
    print(f"요청 실패: {e}")