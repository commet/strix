"""
Debug VBA API calls
"""
from flask import Flask, request, jsonify
from flask_cors import CORS
import json
from datetime import datetime

app = Flask(__name__)
CORS(app)

# 요청 로그 저장
request_log = []

@app.route('/api/query', methods=['POST'])
def query():
    # 요청 데이터 로그
    data = request.get_json()
    
    log_entry = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "question": data.get('question', ''),
        "doc_type": data.get('doc_type', 'both'),
        "headers": dict(request.headers),
        "remote_addr": request.remote_addr
    }
    
    request_log.append(log_entry)
    
    # 콘솔에 출력
    print(f"\n{'='*50}")
    print(f"[{log_entry['timestamp']}] New Request from {log_entry['remote_addr']}")
    print(f"Question: {log_entry['question']}")
    print(f"Doc Type: {log_entry['doc_type']}")
    print(f"{'='*50}\n")
    
    # 요청별로 다른 응답 생성
    question = data.get('question', '').lower()
    
    if '전고체' in question:
        answer = "전고체 배터리 관련 최신 정보입니다..."
    elif 'sk온' in question:
        answer = "SK온의 최근 동향과 전략입니다..."
    elif 'catl' in question.lower():
        answer = "CATL의 시장 점유율과 기술 개발 현황입니다..."
    else:
        # 기본 응답 - SK온 합병 내용
        answer = """SK온과 SK엔무브의 합병은 2025년 배터리 업계의 가장 중요한 이벤트 중 하나입니다.

【합병 개요】
합병 예정일: 2025년 11월 1일
통합법인명: SK이노베이션 배터리 사업부문 (가칭)
자본확충 규모: 약 5조원 (유상증자 방식)"""
    
    response = {
        "answer": answer,
        "total_sources": 5,
        "internal_docs": 3,
        "external_docs": 2,
        "sources": [
            {
                "title": f"Test Document for: {question[:30]}",
                "organization": "Test Org",
                "date": "2025-08-28",
                "type": "internal",
                "content": "Test content",
                "relevance_score": 0.85
            }
        ]
    }
    
    return jsonify(response)

@app.route('/api/log', methods=['GET'])
def get_log():
    """최근 요청 로그 확인"""
    return jsonify(request_log[-10:])  # 최근 10개만

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

if __name__ == '__main__':
    print("VBA Debug Server running on http://localhost:5000")
    print("Check logs at http://localhost:5000/api/log")
    app.run(host='0.0.0.0', port=5000, debug=True)