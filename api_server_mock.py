"""
Mock API Server for STRIX VBA Integration Testing
"""
from flask import Flask, request, jsonify
from flask_cors import CORS
from datetime import datetime

app = Flask(__name__)
CORS(app)

# Mock responses
MOCK_RESPONSES = {
    "default": "안녕하세요! STRIX 시스템이 정상적으로 작동하고 있습니다. VBA와의 연동이 성공적으로 이루어졌습니다.",
    "배터리": "전고체 배터리는 차세대 배터리 기술로, 기존 리튬이온 배터리보다 안전성과 에너지 밀도가 높습니다. 현재 당사는 R&D 부서에서 활발한 연구를 진행 중입니다.",
    "시장": "최근 배터리 시장은 전기차 수요 증가로 급성장하고 있습니다. 특히 중국과 미국 시장에서의 경쟁이 치열해지고 있습니다.",
    "리스크": "주요 리스크 요인으로는 원자재 가격 변동, 환율 변화, 그리고 기술 경쟁 심화가 있습니다. 당사는 이에 대한 대응 전략을 수립하고 있습니다."
}

@app.route('/api/query', methods=['GET', 'POST'])
def query():
    try:
        if request.method == 'GET':
            question = request.args.get('question', '')
            doc_type = request.args.get('doc_type', 'both')
        else:
            data = request.get_json()
            question = data.get('question', '')
            doc_type = data.get('doc_type', 'both')
        
        if not question:
            return jsonify({"error": "No question provided"}), 400
        
        # Find matching mock response
        answer = MOCK_RESPONSES["default"]
        for keyword, response in MOCK_RESPONSES.items():
            if keyword in question:
                answer = response
                break
        
        # Add timestamp
        answer += f"\n\n[생성 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]"
        
        # Prepare response
        response = {
            "answer": answer,
            "internal_docs": 2,
            "external_docs": 3,
            "sources": {
                "internal": [
                    {"title": "2024 Q1 배터리사업 중장기전략", "organization": "전략기획"},
                    {"title": "전고체배터리 개발현황", "organization": "R&D"}
                ],
                "external": [
                    {"title": "글로벌 배터리 시장 동향", "source": "PR팀"},
                    {"title": "EU 배터리 규제 현황", "source": "Google Alert"}
                ]
            }
        }
        
        return jsonify(response)
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "message": "Mock STRIX API Server is running"})

if __name__ == '__main__':
    print("Mock STRIX API Server running on http://localhost:5000")
    print("This is a test server with predefined responses")
    app.run(host='0.0.0.0', port=5000, debug=True)