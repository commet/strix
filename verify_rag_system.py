"""
RAG System Verification Script
실제 데이터와 Mock 데이터를 구분하여 테스트
"""
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from rag.strix_chain import STRIXChain
from langchain_openai import OpenAIEmbeddings
from dotenv import load_dotenv
import json
from datetime import datetime

load_dotenv()

def check_database_contents():
    """데이터베이스 내용 확인"""
    print("\n" + "="*70)
    print("1. DATABASE CONTENTS CHECK")
    print("="*70)
    
    client = SupabaseClient()
    
    # documents 테이블 확인
    try:
        docs = client.client.table('documents').select("*").execute()
        
        if docs.data:
            print(f"\n[OK] Found {len(docs.data)} documents in database")
            
            # 문서 타입별 분류
            internal_docs = [d for d in docs.data if d.get('type') == 'internal']
            external_docs = [d for d in docs.data if d.get('type') == 'external']
            
            print(f"  - Internal documents: {len(internal_docs)}")
            print(f"  - External documents: {len(external_docs)}")
            
            # 샘플 문서 출력
            print("\n  Sample documents:")
            for doc in docs.data[:3]:
                print(f"    - [{doc.get('type')}] {doc.get('title', 'No title')[:60]}...")
                if doc.get('source'):
                    print(f"      Source: {doc.get('source')}")
                    
            # Mock 데이터 여부 확인
            mock_indicators = ['Mock', 'mock', 'test', 'Test', '시뮬레이션']
            mock_docs = [d for d in docs.data if any(ind in str(d.get('title', '')) for ind in mock_indicators)]
            
            if mock_docs:
                print(f"\n[WARNING] Found {len(mock_docs)} mock/test documents")
            else:
                print("\n[OK] No obvious mock data detected")
                
        else:
            print("\n[ERROR] No documents found in database")
            return False
            
    except Exception as e:
        print(f"\n[ERROR] Error accessing documents: {str(e)}")
        return False
    
    # chunks 테이블 확인
    try:
        chunks = client.client.table('chunks').select("id").limit(10).execute()
        print(f"\n[OK] Found {len(chunks.data) if chunks.data else 0} chunks")
    except Exception as e:
        print(f"[ERROR] Error accessing chunks: {str(e)}")
    
    # embeddings 테이블 확인
    try:
        embeddings = client.client.table('embeddings').select("id").limit(10).execute()
        print(f"[OK] Found {len(embeddings.data) if embeddings.data else 0} embeddings")
    except Exception as e:
        print(f"[ERROR] Error accessing embeddings: {str(e)}")
    
    return True

def test_rag_search():
    """RAG 검색 테스트"""
    print("\n" + "="*70)
    print("2. RAG SEARCH TEST")
    print("="*70)
    
    try:
        chain = STRIXChain()
        
        # 테스트 질문들
        test_queries = [
            "배터리 시장 현황은 어떻게 되나요?",
            "전기차 산업 동향을 알려주세요",
            "최근 기술 혁신 사례는?",
        ]
        
        for i, query in enumerate(test_queries, 1):
            print(f"\nTest {i}: {query}")
            print("-" * 50)
            
            try:
                result = chain.invoke(query)
                
                # 결과 분석
                internal_count = len(result.get('internal_docs', []))
                external_count = len(result.get('external_docs', []))
                
                print(f"[OK] Response generated")
                print(f"  - Internal docs used: {internal_count}")
                print(f"  - External docs used: {external_count}")
                
                # 답변 일부 출력
                answer = result.get('answer', '')
                if answer:
                    print(f"  - Answer preview: {answer[:200]}...")
                    
                # 사용된 문서 확인
                if internal_count > 0:
                    print("\n  Internal sources:")
                    for doc in result['internal_docs'][:2]:
                        content_preview = doc.page_content[:100] if hasattr(doc, 'page_content') else 'N/A'
                        print(f"    - {content_preview}...")
                        
                if external_count > 0:
                    print("\n  External sources:")
                    for doc in result['external_docs'][:2]:
                        content_preview = doc.page_content[:100] if hasattr(doc, 'page_content') else 'N/A'
                        print(f"    - {content_preview}...")
                        
            except Exception as e:
                print(f"[ERROR] Search failed: {str(e)}")
                
    except Exception as e:
        print(f"\n[ERROR] Failed to initialize RAG chain: {str(e)}")
        return False
    
    return True

def test_vector_search():
    """벡터 검색 직접 테스트"""
    print("\n" + "="*70)
    print("3. DIRECT VECTOR SEARCH TEST")
    print("="*70)
    
    try:
        client = SupabaseClient()
        embeddings_model = OpenAIEmbeddings()
        
        # 테스트 쿼리 임베딩 생성
        test_query = "배터리 기술 혁신"
        print(f"\nGenerating embedding for: '{test_query}'")
        query_embedding = embeddings_model.embed_query(test_query)
        print(f"[OK] Embedding generated (dimension: {len(query_embedding)})")
        
        # 벡터 검색 수행
        if hasattr(client, 'search_similar_chunks'):
            print("\nPerforming vector search...")
            results = client.search_similar_chunks(query_embedding, limit=5)
            
            if results:
                print(f"[OK] Found {len(results)} similar chunks")
                for i, result in enumerate(results[:3], 1):
                    print(f"\n  Result {i}:")
                    print(f"    Content: {result.get('content', 'N/A')[:150]}...")
                    if result.get('metadata'):
                        print(f"    Metadata: {json.dumps(result['metadata'], ensure_ascii=False, indent=6)[:200]}")
            else:
                print("[ERROR] No results found")
        else:
            print("[ERROR] search_similar_chunks method not implemented")
            
    except Exception as e:
        print(f"\n[ERROR] Vector search failed: {str(e)}")
        return False
    
    return True

def check_api_server():
    """API 서버 상태 확인"""
    print("\n" + "="*70)
    print("4. API SERVER CHECK")
    print("="*70)
    
    import requests
    
    try:
        response = requests.get("http://localhost:5000/health", timeout=2)
        if response.status_code == 200:
            print("[OK] API server is running")
            
            # 테스트 쿼리
            test_response = requests.post(
                "http://localhost:5000/api/query",
                json={"question": "테스트 질문", "doc_type": "both"}
            )
            
            if test_response.status_code == 200:
                data = test_response.json()
                print(f"[OK] API query successful")
                print(f"  - Answer length: {len(data.get('answer', ''))} chars")
                print(f"  - Sources: {data.get('internal_docs', 0)} internal, {data.get('external_docs', 0)} external")
            else:
                print(f"[ERROR] API query failed: {test_response.status_code}")
                
    except requests.exceptions.RequestException:
        print("[ERROR] API server is not running")
        print("  Run: python api_server.py")
    
    return True

def main():
    print("\n" + "="*70)
    print("STRIX RAG SYSTEM VERIFICATION")
    print("="*70)
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. 데이터베이스 확인
    db_ok = check_database_contents()
    
    # 2. RAG 검색 테스트
    if db_ok:
        rag_ok = test_rag_search()
    else:
        print("\n[WARNING] Skipping RAG search test due to database issues")
        rag_ok = False
    
    # 3. 벡터 검색 테스트
    if db_ok:
        vector_ok = test_vector_search()
    else:
        vector_ok = False
    
    # 4. API 서버 확인
    check_api_server()
    
    # 최종 진단
    print("\n" + "="*70)
    print("DIAGNOSIS")
    print("="*70)
    
    if db_ok and rag_ok:
        print("[OK] RAG system is working properly")
    elif db_ok and not rag_ok:
        print("[WARNING] Database has data but RAG search is failing")
        print("  Possible issues:")
        print("  - OpenAI API key issues")
        print("  - Embedding dimension mismatch")
        print("  - search_similar_chunks implementation")
    elif not db_ok:
        print("[ERROR] Database is empty or inaccessible")
        print("\n  To load data:")
        print("  1. Load mock data: python generate_mock_data_with_metadata.py")
        print("  2. Or load real documents using the data ingestion scripts")
    
    print("\n" + "="*70)

if __name__ == "__main__":
    main()