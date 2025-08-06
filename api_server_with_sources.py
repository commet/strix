"""
Enhanced API Server with Source Document Tracking
"""
from flask import Flask, request, jsonify, Response
from flask_cors import CORS
import os
import sys
import asyncio
import json

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.rag.strix_chain_with_sources import STRIXChainWithSources
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)

# Initialize enhanced STRIX chain
chain = STRIXChainWithSources()

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
        
        # Get answer with sources
        result = asyncio.run(chain.ainvoke_with_sources({"question": question}))
        
        # Prepare response
        response = {
            "answer": result['answer'],
            "sources": result['sources'],
            "total_sources": result['total_sources'],
            "internal_docs": len([s for s in result['sources'] if s['type'] == 'internal']),
            "external_docs": len([s for s in result['sources'] if s['type'] == 'external']),
            "created_at": result['timestamp']
        }
        
        # Return JSON with ensure_ascii=False to keep Korean characters
        return Response(
            json.dumps(response, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
        
    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        print(f"ERROR: {str(e)}")
        print(f"TRACEBACK:\n{error_detail}")
        return jsonify({"error": str(e), "detail": error_detail}), 500

@app.route('/api/query/simple', methods=['POST'])
def query_simple():
    """기존 호환성을 위한 간단한 응답"""
    try:
        data = request.get_json()
        question = data.get('question', '')
        
        result = asyncio.run(chain.ainvoke_with_sources({"question": question}))
        
        # 간단한 응답 (기존 형식)
        response = {
            "answer": result['answer'],
            "internal_docs": result['internal_docs'],
            "external_docs": result['external_docs']
        }
        
        return Response(
            json.dumps(response, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
        
    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        print(f"ERROR: {str(e)}")
        print(f"TRACEBACK:\n{error_detail}")
        return jsonify({"error": str(e), "detail": error_detail}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "version": "2.0-with-sources"})

if __name__ == '__main__':
    print("STRIX API Server with Sources running on http://localhost:5000")
    print("Now includes source document references!")
    app.run(host='0.0.0.0', port=5000, debug=True)