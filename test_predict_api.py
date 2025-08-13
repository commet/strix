"""
AI 예측 API 테스트
"""
import requests
import json

print("=== AI 예측 API 테스트 ===\n")

url = "http://localhost:5000/api/issues/predict"
print(f"POST 요청: {url}\n")

try:
    # POST 요청
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json; charset=utf-8"
    }
    
    response = requests.post(url, headers=headers, json={}, timeout=30)
    
    print(f"응답 상태: {response.status_code}")
    print(f"응답 헤더: {response.headers.get('Content-Type', 'N/A')}")
    
    if response.status_code == 200:
        data = response.json()
        print(f"\n응답 데이터:")
        print(json.dumps(data, indent=2, ensure_ascii=False))
    else:
        print(f"\n오류 응답:")
        print(response.text)
        
except requests.exceptions.Timeout:
    print("요청 시간 초과 (30초)")
except Exception as e:
    print(f"오류 발생: {e}")
    
print("\n=== 테스트 완료 ===")