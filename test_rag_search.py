"""
Test RAG Search to Debug Reference Issues
"""
import asyncio
import sys
import os
import json

sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.rag.strix_chain_with_sources import STRIXChainWithSources
from dotenv import load_dotenv

load_dotenv()

async def test_search():
    print("=== STRIX RAG Search Test ===\n")
    
    # Initialize chain
    chain = STRIXChainWithSources()
    
    # Test question
    test_question = "전고체 배터리 개발 현황은?"
    print(f"Question: {test_question}\n")
    
    # Execute search and formatting
    print("1. Searching documents...")
    search_result = await chain.search_and_format_with_sources({"question": test_question})
    
    print("\n2. Search results:")
    print(f"- Internal context length: {len(search_result.get('internal_context', ''))}")
    print(f"- External context length: {len(search_result.get('external_context', ''))}")
    
    print("\n3. Stored source documents:")
    if hasattr(chain, 'source_documents'):
        internal_docs = chain.source_documents.get('internal', [])
        external_docs = chain.source_documents.get('external', [])
        
        print(f"- Internal docs: {len(internal_docs)}")
        for i, doc in enumerate(internal_docs, 1):
            print(f"  [{i}] {doc.get('metadata', {}).get('title', 'N/A')}")
            
        print(f"- External docs: {len(external_docs)}")
        for i, doc in enumerate(external_docs, len(internal_docs) + 1):
            print(f"  [{i}] {doc.get('metadata', {}).get('title', 'N/A')}")
    
    print("\n4. Generating full response...")
    result = await chain.ainvoke_with_sources({"question": test_question})
    
    print("\n5. Final results:")
    print(f"- Answer length: {len(result['answer'])}")
    print(f"- Source count: {result['total_sources']}")
    print("\nSource details:")
    for src in result['sources']:
        print(f"  [{src['number']}] {src['title']} ({src['type']}) - {src['organization']}")
    
    # Check JSON response
    print("\n6. JSON response sample:")
    json_response = json.dumps(result, ensure_ascii=False, indent=2)
    print(json_response[:1000] + "..." if len(json_response) > 1000 else json_response)

if __name__ == "__main__":
    asyncio.run(test_search())