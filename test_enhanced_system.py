"""
Test enhanced RAG system with more documents and longer answers
"""
import requests
import json

# API endpoint
url = "http://localhost:5000/api/query"

# 경쟁사 관련 테스트 질문
test_query = "경쟁사 대비 우리의 강점과 약점을 분석해주세요"

print("=== Testing Enhanced RAG System ===\n")
print(f"Question: {test_query}")
print("-" * 80)

try:
    response = requests.post(url, json={"question": test_query})
    
    if response.status_code == 200:
        data = response.json()
        answer = data.get('answer', 'No answer')
        sources = data.get('sources', [])
        
        print("\n【Answer】")
        print(answer)
        
        print(f"\n【Sources】 Total: {len(sources)} documents")
        print("-" * 80)
        
        # 내부 문서
        internal_sources = [s for s in sources if s.get('type') == 'internal']
        print(f"\n내부 문서 ({len(internal_sources)}개):")
        for i, source in enumerate(internal_sources, 1):
            print(f"  [{i}] {source.get('title', 'No title')}")
            print(f"      - {source.get('organization', 'N/A')} | {source.get('date', 'N/A')}")
        
        # 외부 문서  
        external_sources = [s for s in sources if s.get('type') == 'external']
        print(f"\n외부 문서 ({len(external_sources)}개):")
        for i, source in enumerate(external_sources, 1):
            print(f"  [{len(internal_sources)+i}] {source.get('title', 'No title')}")
            print(f"      - {source.get('organization', 'N/A')} | {source.get('date', 'N/A')}")
        
        # 통계
        print("\n【Statistics】")
        print(f"- Answer length: {len(answer)} characters")
        print(f"- Total sources: {data.get('total_sources', 0)}")
        print(f"- Internal docs: {data.get('internal_docs', 0)}")
        print(f"- External docs: {data.get('external_docs', 0)}")
        
    else:
        print(f"Error: Status {response.status_code}")
        
except Exception as e:
    print(f"Error: {e}")