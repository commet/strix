"""
Debug API Server for STRIX VBA Integration
"""
from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import sys
import asyncio

# UTF-8 encoding
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from rag.strix_chain import STRIXChain
from database.supabase_client import SupabaseClient
from langchain_openai import OpenAIEmbeddings
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)

# Initialize
chain = STRIXChain()
client = SupabaseClient()
embeddings = OpenAIEmbeddings()

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
        
        print(f"\n=== DEBUG: Received question: {question}")
        
        if not question:
            return jsonify({"error": "No question provided"}), 400
        
        # Direct search test
        print("=== DEBUG: Testing direct search first...")
        query_embedding = embeddings.embed_query(question)
        search_results = client.search_similar_chunks(query_embedding, limit=3)
        print(f"=== DEBUG: Direct search found {len(search_results)} results")
        
        for i, res in enumerate(search_results):
            print(f"  Result {i+1}: {res.get('doc_title', 'N/A')} (similarity: {res.get('similarity', 0):.3f})")
        
        # Use RAG chain
        print("=== DEBUG: Now using RAG chain...")
        result = asyncio.run(chain.ainvoke({"question": question}))
        
        print(f"=== DEBUG: RAG chain found {len(result.get('internal_docs', []))} internal, {len(result.get('external_docs', []))} external docs")
        
        # Prepare response
        response = {
            "answer": result['answer'],
            "internal_docs": len(result.get('internal_docs', [])),
            "external_docs": len(result.get('external_docs', [])),
            "debug_info": {
                "direct_search_count": len(search_results),
                "rag_internal_count": len(result.get('internal_docs', [])),
                "rag_external_count": len(result.get('external_docs', []))
            }
        }
        
        return jsonify(response)
        
    except Exception as e:
        print(f"=== DEBUG: Error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

if __name__ == '__main__':
    print("STRIX Debug API Server running on http://localhost:5000")
    print("Encoding: UTF-8")
    app.run(host='0.0.0.0', port=5000, debug=True)