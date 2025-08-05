"""
STRIX Streamlit Demo App with API Endpoint
"""
import streamlit as st
import os
import sys
import asyncio
from datetime import datetime
import json

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from rag.strix_chain import STRIXChain
from database.supabase_client import SupabaseClient

# API endpoint handling
from streamlit.components.v1 import html
import streamlit.components.v1 as components

# Page config
st.set_page_config(
    page_title="STRIX - Strategic Intelligence System",
    page_icon="🦉",
    layout="wide"
)

# Check if this is an API request
query_params = st.query_params
if "api" in query_params and query_params["api"] == "true":
    # API mode - handle JSON request/response
    if "question" in query_params:
        chain = STRIXChain()
        question = query_params["question"]
        doc_type = query_params.get("doc_type", "both")
        
        # Map doc_type
        if doc_type == "internal":
            question += " (내부 문서에서만 검색)"
        elif doc_type == "external":
            question += " (외부 뉴스에서만 검색)"
        
        # Get answer
        result = asyncio.run(chain.ainvoke({"question": question}))
        
        # Return JSON response
        response = {
            "answer": result['answer'],
            "internal_docs": len(result.get('internal_docs', [])),
            "external_docs": len(result.get('external_docs', [])),
            "sources": {
                "internal": [{"title": doc.metadata.get('title'), 
                            "organization": doc.metadata.get('organization')} 
                           for doc in result.get('internal_docs', [])[:3]],
                "external": [{"title": doc.metadata.get('title'), 
                            "source": doc.metadata.get('source')} 
                           for doc in result.get('external_docs', [])[:3]]
            }
        }
        
        st.json(response)
        st.stop()

# Regular Streamlit UI mode
# Initialize session state
if 'chain' not in st.session_state:
    st.session_state.chain = STRIXChain()
if 'supabase' not in st.session_state:
    st.session_state.supabase = SupabaseClient()
if 'messages' not in st.session_state:
    st.session_state.messages = []

# Title and description
st.title("🦉 STRIX - Strategic Intelligence System")
st.markdown("내부 문서와 외부 뉴스를 통합하여 전략적 인사이트를 제공합니다.")

# Add API usage info in sidebar
with st.sidebar:
    st.header("📊 시스템 상태")
    
    # Get database stats
    try:
        stats = st.session_state.supabase.client.rpc('get_document_stats').execute()
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("전체 문서", stats.data[0]['stat_value'])
            st.metric("내부 문서", stats.data[1]['stat_value'])
        with col2:
            st.metric("외부 뉴스", stats.data[2]['stat_value'])
            st.metric("청크 수", stats.data[3]['stat_value'])
    except:
        st.info("통계를 불러올 수 없습니다.")
    
    st.divider()
    
    # API 사용법
    with st.expander("🔌 API 사용법"):
        st.code("""
# VBA에서 사용:
http://localhost:8501/?api=true&question=YOUR_QUESTION&doc_type=both

# Parameters:
- question: 질문 (URL encoded)
- doc_type: both|internal|external
        """)
    
    st.divider()
    
    # Search settings
    st.header("🔍 검색 설정")
    search_type = st.selectbox("문서 유형", ["전체", "내부 문서만", "외부 뉴스만"])
    
    # Categories
    categories = st.multiselect(
        "카테고리 필터",
        ["Macro", "산업", "기술", "리스크", "경쟁사", "정책"]
    )
    
    # Organizations
    organizations = st.multiselect(
        "조직 필터",
        ["전략기획", "R&D", "경영지원", "생산", "영업마케팅"]
    )

# Main chat interface
st.header("💬 질문하기")

# Display chat messages
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.write(message["content"])
        
        # Show sources if available
        if "sources" in message:
            with st.expander("📚 참고 문서"):
                if message["sources"]["internal"]:
                    st.subheader("내부 문서")
                    for doc in message["sources"]["internal"]:
                        st.write(f"- {doc['title']} ({doc['organization']})")
                
                if message["sources"]["external"]:
                    st.subheader("외부 뉴스")
                    for doc in message["sources"]["external"]:
                        st.write(f"- {doc['title']} ({doc['source']})")

# Chat input
if prompt := st.chat_input("질문을 입력하세요..."):
    # Add user message
    st.session_state.messages.append({"role": "user", "content": prompt})
    
    with st.chat_message("user"):
        st.write(prompt)
    
    # Generate response
    with st.chat_message("assistant"):
        with st.spinner("답변을 생성하고 있습니다..."):
            try:
                # Prepare query with filters
                query = prompt
                if search_type == "내부 문서만":
                    query += " (내부 문서에서만 검색)"
                elif search_type == "외부 뉴스만":
                    query += " (외부 뉴스에서만 검색)"
                
                # Run the chain
                result = asyncio.run(st.session_state.chain.ainvoke({"question": query}))
                
                # Display answer
                st.write(result['answer'])
                
                # Prepare sources
                sources = {
                    "internal": [],
                    "external": []
                }
                
                for doc in result.get('internal_docs', []):
                    sources["internal"].append({
                        "title": doc.metadata.get('title', 'Untitled'),
                        "organization": doc.metadata.get('organization', 'Unknown')
                    })
                
                for doc in result.get('external_docs', []):
                    sources["external"].append({
                        "title": doc.metadata.get('title', 'Untitled'),
                        "source": doc.metadata.get('source', 'Unknown')
                    })
                
                # Add assistant message with sources
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": result['answer'],
                    "sources": sources
                })
                
                # Show sources in expander
                if sources["internal"] or sources["external"]:
                    with st.expander("📚 참고 문서"):
                        if sources["internal"]:
                            st.subheader("내부 문서")
                            for doc in sources["internal"]:
                                st.write(f"- {doc['title']} ({doc['organization']})")
                        
                        if sources["external"]:
                            st.subheader("외부 뉴스")
                            for doc in sources["external"]:
                                st.write(f"- {doc['title']} ({doc['source']})")
                
            except Exception as e:
                st.error(f"오류가 발생했습니다: {str(e)}")
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": f"죄송합니다. 오류가 발생했습니다: {str(e)}"
                })

# Footer
st.divider()
st.caption("STRIX - Strategic Intelligence System | Powered by LangChain & Supabase")