"""
STRIX RAG Chain using LangGraph
"""
from typing import List, Dict, Any, TypedDict, Annotated, Literal
from typing_extensions import TypedDict
from langchain_core.documents import Document
from langchain_openai import ChatOpenAI, OpenAIEmbeddings
from langchain_core.prompts import ChatPromptTemplate
from langgraph.graph import START, StateGraph
import logging
from config import DEFAULT_MODEL, TEMPERATURE, EMBEDDING_MODEL
from database.supabase_client import SupabaseClient

logger = logging.getLogger(__name__)

# Define State
class STRIXState(TypedDict):
    question: str
    query: Dict[str, Any]  # Structured query
    internal_docs: List[Document]
    external_docs: List[Document]
    context: str
    answer: str
    metadata: Dict[str, Any]

class QueryStructure(TypedDict):
    """Structured query for better retrieval"""
    query: str
    time_range: Literal["1week", "1month", "3months", "6months", "1year", "all"]
    doc_type: Literal["internal", "external", "both"]
    categories: List[str]
    organizations: List[str]

class STRIXChain:
    """Main RAG chain for STRIX"""
    
    def __init__(self):
        self.supabase = SupabaseClient()
        self.llm = ChatOpenAI(model=DEFAULT_MODEL, temperature=TEMPERATURE)
        self.embeddings = OpenAIEmbeddings(model=EMBEDDING_MODEL)
        self.graph = self._build_graph()
    
    def _build_graph(self):
        """Build LangGraph chain"""
        
        # Define nodes
        def analyze_query(state: STRIXState):
            """Analyze and structure the query"""
            prompt = ChatPromptTemplate.from_messages([
                ("system", """당신은 질의 분석 전문가입니다. 사용자의 질문을 분석하여 구조화된 검색 쿼리로 변환하세요.
                
                다음 정보를 추출하세요:
                - query: 핵심 검색어
                - time_range: 시간 범위 (1week/1month/3months/6months/1year/all)
                - doc_type: 문서 유형 (internal/external/both)
                - categories: 관련 카테고리 리스트 (Macro/산업/기술/리스크/경쟁사/정책)
                - organizations: 관련 조직 리스트 (전략기획/R&D/경영지원/생산/영업마케팅)
                """),
                ("user", "{question}")
            ])
            
            structured_llm = self.llm.with_structured_output(QueryStructure, method="function_calling")
            messages = prompt.format_messages(question=state["question"])
            query = structured_llm.invoke(messages)
            
            return {"query": dict(query)}
        
        def retrieve_documents(state: STRIXState):
            """Retrieve relevant documents"""
            query = state["query"]
            
            # Generate embedding for query
            query_embedding = self.embeddings.embed_query(query["query"])
            
            # Prepare filters
            filters = {}
            if query["time_range"] != "all":
                # Add time filter logic
                pass
            
            # Search internal documents
            internal_docs = []
            if query["doc_type"] in ["internal", "both"]:
                internal_results = self.supabase.search_similar_chunks(
                    query_embedding, 
                    limit=5,
                    filters={**filters, "type": "internal"}
                )
                internal_docs = [
                    Document(
                        page_content=r["content"],
                        metadata=r["metadata"]
                    ) for r in internal_results
                ]
            
            # Search external documents
            external_docs = []
            if query["doc_type"] in ["external", "both"]:
                external_results = self.supabase.search_similar_chunks(
                    query_embedding,
                    limit=5,
                    filters={**filters, "type": "external"}
                )
                external_docs = [
                    Document(
                        page_content=r["content"],
                        metadata=r["metadata"]
                    ) for r in external_results
                ]
            
            return {
                "internal_docs": internal_docs,
                "external_docs": external_docs
            }
        
        def generate_answer(state: STRIXState):
            """Generate answer using retrieved documents"""
            
            # Prepare context
            context_parts = []
            
            if state["internal_docs"]:
                context_parts.append("=== 내부 문서 ===")
                for doc in state["internal_docs"]:
                    context_parts.append(f"[{doc.metadata.get('organization', 'N/A')}] {doc.page_content}")
            
            if state["external_docs"]:
                context_parts.append("\n=== 외부 뉴스 ===")
                for doc in state["external_docs"]:
                    context_parts.append(f"[{doc.metadata.get('category', 'N/A')}] {doc.page_content}")
            
            context = "\n\n".join(context_parts)
            
            # Generate answer
            prompt = ChatPromptTemplate.from_messages([
                ("system", """당신은 기업 전략 분석 전문가입니다. 
                제공된 내부 문서와 외부 뉴스를 바탕으로 통합적인 인사이트를 제공하세요.
                
                답변 시 다음을 포함하세요:
                1. 핵심 요약
                2. 내부 현황과 외부 환경의 연관성
                3. 주요 시사점 및 제언
                
                답변은 명확하고 실행 가능한 인사이트를 포함해야 합니다."""),
                ("user", "질문: {question}\n\n컨텍스트:\n{context}")
            ])
            
            response = self.llm.invoke(
                prompt.format_messages(question=state["question"], context=context)
            )
            
            return {
                "context": context,
                "answer": response.content,
                "metadata": {
                    "internal_doc_count": len(state["internal_docs"]),
                    "external_doc_count": len(state["external_docs"]),
                    "query_structure": state["query"]
                }
            }
        
        # Build graph
        graph_builder = StateGraph(STRIXState)
        graph_builder.add_node("analyze_query", analyze_query)
        graph_builder.add_node("retrieve_documents", retrieve_documents)
        graph_builder.add_node("generate_answer", generate_answer)
        
        # Add edges
        graph_builder.add_edge(START, "analyze_query")
        graph_builder.add_edge("analyze_query", "retrieve_documents")
        graph_builder.add_edge("retrieve_documents", "generate_answer")
        
        return graph_builder.compile()
    
    def invoke(self, question: str) -> Dict[str, Any]:
        """Process a question through the RAG chain"""
        try:
            result = self.graph.invoke({"question": question})
            
            # Log search for learning
            self.supabase.log_search(
                query=question,
                results={
                    "answer": result["answer"],
                    "internal_docs": len(result.get("internal_docs", [])),
                    "external_docs": len(result.get("external_docs", []))
                },
                metadata=result.get("metadata", {})
            )
            
            return result
        except Exception as e:
            logger.error(f"Error in RAG chain: {e}")
            raise
    
    async def ainvoke(self, inputs: Dict[str, Any]) -> Dict[str, Any]:
        """Async process a question through the RAG chain"""
        return self.invoke(inputs.get("question", ""))
    
    def stream(self, question: str):
        """Stream the RAG chain execution"""
        try:
            for step in self.graph.stream({"question": question}, stream_mode="updates"):
                yield step
        except Exception as e:
            logger.error(f"Error in RAG stream: {e}")
            raise