"""
Test RAG Search and Response Generation
"""
import os
import sys
import asyncio
from dotenv import load_dotenv

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from rag.strix_chain import STRIXChain
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

async def test_rag_search():
    """Test RAG search functionality"""
    load_dotenv()
    
    print("STRIX RAG Search Test")
    print("="*50)
    
    # Initialize chain
    chain = STRIXChain()
    
    # Test queries
    test_queries = [
        "전고체 배터리 개발 현황은 어떻게 되나요?",
        "최근 배터리 시장 동향을 알려주세요",
        "리튬 가격 변동과 리스크 관리 방안은?",
        "EU 배터리 규제 내용을 요약해주세요",
        "우리 회사의 전고체 배터리 투자 계획은?"
    ]
    
    for i, query in enumerate(test_queries, 1):
        print(f"\n{i}. Query: {query}")
        print("-" * 50)
        
        try:
            # Invoke the chain
            result = await chain.ainvoke({"question": query})
            
            print(f"Answer: {result['answer']}")
            print(f"\nFound {len(result['internal_docs'])} internal documents")
            print(f"Found {len(result['external_docs'])} external documents")
            
            # Show source documents
            if result['internal_docs']:
                print("\nInternal Sources:")
                for doc in result['internal_docs'][:2]:  # Show first 2
                    print(f"  - {doc.metadata.get('title', 'Untitled')} ({doc.metadata.get('organization', 'Unknown')})")
            
            if result['external_docs']:
                print("\nExternal Sources:")
                for doc in result['external_docs'][:2]:  # Show first 2
                    print(f"  - {doc.metadata.get('title', 'Untitled')} ({doc.metadata.get('source', 'Unknown')})")
                    
        except Exception as e:
            logger.error(f"Error processing query: {e}")
            print(f"[ERROR] Failed to process query: {e}")
    
    # Skip interactive mode to avoid EOF errors

if __name__ == "__main__":
    asyncio.run(test_rag_search())