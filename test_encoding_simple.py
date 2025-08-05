"""
Simple test to verify RAG is returning actual content
"""
import os
import sys
import asyncio

# Set encoding
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from rag.strix_chain import STRIXChain
from dotenv import load_dotenv

load_dotenv()

async def test_simple():
    chain = STRIXChain()
    
    # Test query
    question = "우리 회사의 전고체 배터리 투자 계획은?"
    print(f"Question: {question}")
    print("-" * 50)
    
    result = await chain.ainvoke({"question": question})
    
    # Check if we got real content
    answer = result['answer']
    print(f"Answer preview (first 100 chars): {answer[:100]}...")
    print(f"Answer contains '5조원': {'5조원' in answer}")
    print(f"Answer contains '1.5조원': {'1.5조원' in answer}")
    print(f"Internal docs found: {len(result.get('internal_docs', []))}")
    print(f"External docs found: {len(result.get('external_docs', []))}")
    
    # Show source titles if available
    if result.get('internal_docs'):
        print("\nInternal document titles:")
        for doc in result['internal_docs'][:3]:
            print(f"  - {doc.get('metadata', {}).get('doc_title', 'N/A')}")

if __name__ == "__main__":
    asyncio.run(test_simple())