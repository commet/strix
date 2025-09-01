"""
Test API with different queries
"""
import requests
import json

# API endpoint
url = "http://localhost:5000/api/query"

# Test queries
test_queries = [
    "전고체 배터리 현황",
    "CATL 최근 동향",
    "리튬 가격",
    "배터리 재활용",
    "테슬라 소식"
]

print("=== TESTING API RESPONSES ===\n")

for query in test_queries:
    print(f"Query: {query}")
    print("-" * 50)
    
    try:
        response = requests.post(url, json={"question": query})
        
        if response.status_code == 200:
            data = response.json()
            answer = data.get('answer', 'No answer')
            
            # Show first 400 chars of answer
            print(f"Answer preview: {answer[:400]}...")
            
            # Check if sources are returned
            sources = data.get('sources', [])
            if sources:
                print(f"Sources: {len(sources)} documents")
                for i, source in enumerate(sources[:3], 1):
                    print(f"  {i}. {source.get('title', 'No title')}")
        else:
            print(f"Error: Status {response.status_code}")
            
    except Exception as e:
        print(f"Error: {e}")
    
    print("\n" + "="*50 + "\n")