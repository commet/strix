"""
Enhanced STRIX Chain with Source Document Tracking
"""
import asyncio
from typing import Dict, List, Any
from datetime import datetime
from langchain_openai import ChatOpenAI
from langchain_openai import OpenAIEmbeddings
from langchain.prompts import ChatPromptTemplate
from langchain.schema.runnable import RunnablePassthrough, RunnableLambda
from langchain.schema.output_parser import StrOutputParser
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from database.supabase_client import SupabaseClient


class STRIXChainWithSources:
    def __init__(self):
        self.embeddings = OpenAIEmbeddings()
        self.llm = ChatOpenAI(temperature=0, model="gpt-4o-mini")
        self.supabase = SupabaseClient()
        self.chain = self._build_chain()
        
    def _build_chain(self):
        prompt = ChatPromptTemplate.from_messages([
            ("system", """당신은 회사의 전략 정보를 분석하는 AI 어시스턴트입니다.
            제공된 문서들을 바탕으로 정확하고 통찰력 있는 답변을 제공하세요.
            
            답변할 때 다음 형식을 따라주세요:
            1. 핵심 답변을 먼저 제시
            2. 근거가 되는 내용을 인용할 때는 [1], [2] 등의 번호를 사용
            3. 내부 문서와 외부 뉴스를 구분하여 설명
            
            내부 문서:
            {internal_context}
            
            외부 뉴스:
            {external_context}
            """),
            ("user", "{question}")
        ])
        
        chain = (
            RunnableLambda(self.search_and_format_with_sources)
            | prompt
            | self.llm
            | StrOutputParser()
        )
        
        return chain
    
    async def search_and_format_with_sources(self, inputs: Dict[str, Any]) -> Dict[str, Any]:
        """검색하고 소스 문서 정보와 함께 포맷팅"""
        # inputs가 dict인 경우와 string인 경우 모두 처리
        if isinstance(inputs, dict):
            question = inputs.get("question", "")
        else:
            question = str(inputs)
        
        # 벡터 검색 수행
        internal_docs, external_docs = await self._search_documents(question)
        
        # 컨텍스트 생성 (번호 포함)
        internal_context = self._format_documents_with_numbers(internal_docs, start_num=1)
        external_context = self._format_documents_with_numbers(
            external_docs, 
            start_num=len(internal_docs) + 1
        )
        
        # 소스 문서 정보 저장
        self.source_documents = {
            "internal": internal_docs,
            "external": external_docs
        }
        
        return {
            "question": question,
            "internal_context": internal_context,
            "external_context": external_context
        }
    
    async def _search_documents(self, query: Any) -> tuple:
        """문서 검색 및 메타데이터 포함"""
        # 쿼리 임베딩 생성
        query_embedding = await self.embeddings.aembed_query(query)
        
        # 내부 문서 검색
        all_results = self.supabase.search_similar_chunks(
            query_embedding=query_embedding,
            limit=6  # 내부 3개 + 외부 3개
        )
        
        # 타입별로 분리
        internal_results = []
        external_results = []
        
        for result in all_results:
            # doc_type 필드명 수정
            doc_type = result.get('doc_type', 'unknown')
            
            # 결과에 type 필드 추가 (하위 호환성)
            result['type'] = doc_type
            result['metadata'] = result.get('metadata', {})
            if 'doc_title' in result:
                result['metadata']['title'] = result['doc_title']
            if 'doc_organization' in result:
                result['metadata']['organization'] = result['doc_organization']
            if 'doc_category' in result:
                result['metadata']['category'] = result['doc_category']
            # created_at 필드도 메타데이터에 추가
            if 'created_at' in result and result['created_at']:
                result['metadata']['created_at'] = result['created_at']
            
            if doc_type == 'internal':
                internal_results.append(result)
            else:
                external_results.append(result)
        
        # 각각 최대 3개까지만
        internal_results = internal_results[:3]
        external_results = external_results[:3]
        
        return internal_results, external_results
    
    def _format_documents_with_numbers(self, docs: List[Dict], start_num: int) -> str:
        """문서를 번호와 함께 포맷팅"""
        if not docs:
            return "관련 문서가 없습니다."
        
        formatted = []
        for i, doc in enumerate(docs, start=start_num):
            # 메타데이터 추출
            title = doc.get('metadata', {}).get('title', '제목 없음')
            org = doc.get('metadata', {}).get('organization', '조직 미상')
            date = doc.get('metadata', {}).get('created_at', '')
            content = doc.get('content', '')
            
            formatted.append(
                f"[{i}] {title} ({org}, {date[:10] if date else '날짜 미상'})\n"
                f"내용: {content[:500]}..."
            )
        
        return "\n\n".join(formatted)
    
    async def ainvoke_with_sources(self, inputs: Dict[str, Any]) -> Dict[str, Any]:
        """답변과 함께 소스 문서 정보 반환"""
        # 답변 생성
        answer = await self.chain.ainvoke(inputs)
        
        # 소스 문서 정보 포맷팅
        sources = []
        
        # 내부 문서 소스
        for i, doc in enumerate(self.source_documents.get("internal", []), start=1):
            sources.append({
                "number": i,
                "type": "internal",
                "title": doc.get('metadata', {}).get('title', '제목 없음'),
                "organization": doc.get('metadata', {}).get('organization') or doc.get('doc_organization') or '조직 미상',
                "date": doc.get('metadata', {}).get('created_at', '')[:10] if doc.get('metadata', {}).get('created_at') else doc.get('created_at', '')[:10] if doc.get('created_at') else '',
                "file_path": doc.get('metadata', {}).get('file_path', ''),
                "snippet": doc.get('content', '')[:200] + "...",
                "relevance_score": doc.get('similarity', 0)
            })
        
        # 외부 문서 소스
        start_num = len(self.source_documents.get("internal", [])) + 1
        for i, doc in enumerate(self.source_documents.get("external", []), start=start_num):
            sources.append({
                "number": i,
                "type": "external",
                "title": doc.get('metadata', {}).get('title', '제목 없음'),
                "organization": doc.get('metadata', {}).get('organization') or doc.get('doc_organization') or '출처 미상',
                "date": doc.get('metadata', {}).get('created_at', '')[:10] if doc.get('metadata', {}).get('created_at') else doc.get('created_at', '')[:10] if doc.get('created_at') else '',
                "url": doc.get('metadata', {}).get('url', ''),
                "snippet": doc.get('content', '')[:200] + "...",
                "relevance_score": doc.get('similarity', 0)
            })
        
        return {
            "answer": answer,
            "sources": sources,
            "total_sources": len(sources),
            "timestamp": datetime.now().isoformat()
        }