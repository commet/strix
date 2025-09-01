"""
Test answer generation to debug why answers are always the same
"""
import asyncio
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from rag.strix_chain_with_sources import STRIXChainWithSources

async def test_different_queries():
    """Test with different queries to see if answers change"""
    
    chain = STRIXChainWithSources()
    
    test_queries = [
        "전고체 배터리 현황은?",
        "CATL의 최근 동향은?",
        "리튬 가격 전망은?",
        "SK온의 최근 소식은?",
        "배터리 재활용 시장은?"
    ]
    
    print("=== TESTING ANSWER GENERATION ===\n")
    
    for query in test_queries:
        print(f"Query: {query}")
        print("-" * 50)
        
        try:
            # Get the search results and context
            formatted_inputs = await chain.search_and_format_with_sources({"question": query})
            
            print(f"Internal context length: {len(formatted_inputs['internal_context'])}")
            print(f"External context length: {len(formatted_inputs['external_context'])}")
            
            # Show first 200 chars of each context
            print(f"Internal preview: {formatted_inputs['internal_context'][:200]}...")
            print(f"External preview: {formatted_inputs['external_context'][:200]}...")
            
            # Generate answer
            answer = await chain.chain.ainvoke({"question": query})
            
            # Show first 300 chars of answer
            print(f"Answer preview: {answer[:300]}...")
            print("\n" + "="*50 + "\n")
            
        except Exception as e:
            print(f"Error: {e}")
            print("\n" + "="*50 + "\n")

if __name__ == "__main__":
    asyncio.run(test_different_queries())