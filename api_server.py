"""
Simple API Server for STRIX VBA Integration
"""
from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import sys
import asyncio

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
        
        # Prepare response
        response = {
            "answer": result['answer'],
            "internal_docs": len(result.get('internal_docs', [])),
            "external_docs": len(result.get('external_docs', [])),
            "sources": {
                "internal": [{"title": doc.metadata.get('title', 'Untitled'), 
                            "organization": doc.metadata.get('organization', 'Unknown')} 
                           for doc in result.get('internal_docs', [])[:3]],
                "external": [{"title": doc.metadata.get('title', 'Untitled'), 
                            "source": doc.metadata.get('source', 'Unknown')} 
                           for doc in result.get('external_docs', [])[:3]]
            }
        }
        
        return jsonify(response)
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

if __name__ == '__main__':
    print("STRIX API Server running on http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)