"""
API 서버 연결 테스트
"""
import requests
import json

print("=== API 서버 연결 테스트 ===\n")

# 1. 기본 헬스 체크
try:
    response = requests.get("http://localhost:5000/health")
    print(f"1. 헬스 체크: {response.status_code}")
    if response.status_code == 200:
        print(f"   응답: {response.json()}")
except Exception as e:
    print(f"1. 헬스 체크 실패: {e}")

# 2. 이슈 목록 조회
try:
    response = requests.get("http://localhost:5000/api/issues")
    print(f"\n2. 이슈 목록: {response.status_code}")
    if response.status_code == 200:
        data = response.json()
        print(f"   이슈 개수: {len(data)}")
        if len(data) > 0:
            print(f"   첫 번째 이슈 ID: {data[0].get('id', 'N/A')}")
except Exception as e:
    print(f"\n2. 이슈 목록 조회 실패: {e}")

# 3. POST 테스트 (AI 예측)
try:
    response = requests.post("http://localhost:5000/api/issues/predict", 
                            headers={"Content-Type": "application/json"},
                            data="{}")
    print(f"\n3. AI 예측 POST: {response.status_code}")
    if response.status_code == 200:
        print(f"   응답: {response.json()}")
except Exception as e:
    print(f"\n3. AI 예측 실패: {e}")

print("\n=== 테스트 완료 ===")
print("\nAPI 서버가 정상 작동 중입니다!" if response.status_code == 200 else "API 서버에 문제가 있습니다.")