"""
Simple API Server with Reference Support
"""
from flask import Flask, request, jsonify, Response
from flask_cors import CORS
import os
import sys
import asyncio
import json

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from rag.strix_chain import STRIXChain
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)

# Initialize STRIX chain
chain = STRIXChain()

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
        
        # Add doc type filter to question
        if doc_type == "internal":
            question += " (내부 문서에서만 검색)"
        elif doc_type == "external":
            question += " (외부 뉴스에서만 검색)"
        
        # Get answer from STRIX
        result = asyncio.run(chain.ainvoke({"question": question}))
        
        # 간단한 레퍼런스 추가
        sources = []
        
        # 내부 문서 레퍼런스
        if 'internal_docs' in result and result['internal_docs']:
            for i, doc in enumerate(result['internal_docs'][:3], 1):
                sources.append({
                    "number": i,
                    "type": "internal",
                    "title": doc.metadata.get('title', '제목 없음'),
                    "organization": doc.metadata.get('organization', '조직 미상'),
                    "date": doc.metadata.get('created_at', '')[:10] if doc.metadata.get('created_at') else '',
                    "snippet": doc.page_content[:200] + "..."
                })
        
        # 외부 문서 레퍼런스
        start_num = len(sources) + 1
        if 'external_docs' in result and result['external_docs']:
            for i, doc in enumerate(result['external_docs'][:3], start_num):
                sources.append({
                    "number": i,
                    "type": "external", 
                    "title": doc.metadata.get('title', '제목 없음'),
                    "organization": doc.metadata.get('organization', '출처 미상'),
                    "date": doc.metadata.get('created_at', '')[:10] if doc.metadata.get('created_at') else '',
                    "snippet": doc.page_content[:200] + "..."
                })
        
        # 답변에 레퍼런스 번호 추가
        answer = result['answer']
        if sources:
            answer += "\n\n📚 참고 문서:\n"
            for src in sources:
                answer += f"[{src['number']}] {src['title']} ({src['organization']})\n"
        
        # Prepare response
        response = {
            "answer": answer,
            "sources": sources,
            "total_sources": len(sources),
            "internal_docs": len(result.get('internal_docs', [])),
            "external_docs": len(result.get('external_docs', [])),
            "created_at": result.get('created_at', '')
        }
        
        # Return JSON with ensure_ascii=False to keep Korean characters
        return Response(
            json.dumps(response, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
        
    except Exception as e:
        import traceback
        print(f"Error: {str(e)}")
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

if __name__ == '__main__':
    print("STRIX Simple Reference API Server running on http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)