# STRIX RAG Architecture

## 개요
STRIX를 LangChain 기반 RAG(Retrieval Augmented Generation) 시스템으로 재설계합니다. 
내부 보고서와 외부 뉴스를 통합하여 지능적인 인사이트를 제공하는 시스템입니다.

## 기술 스택
- **Framework**: LangChain
- **Database**: Supabase (Vector Store + Metadata)
- **Embedding**: OpenAI Embeddings
- **LLM**: GPT-4 (또는 사내 LLM)
- **Frontend**: Streamlit (프로토타입)

## 시스템 아키텍처

```
┌─────────────────────────────────────────────────────────────┐
│                      STRIX RAG System                        │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌─────────────┐        ┌─────────────┐                    │
│  │  Document   │        │   News      │                    │
│  │  Ingestion  │        │  Ingestion  │                    │
│  └──────┬──────┘        └──────┬──────┘                    │
│         │                       │                            │
│         ▼                       ▼                            │
│  ┌─────────────────────────────────────┐                   │
│  │        Text Splitter                 │                   │
│  │  (Chunk documents intelligently)     │                   │
│  └──────────────┬──────────────────────┘                   │
│                 │                                            │
│                 ▼                                            │
│  ┌─────────────────────────────────────┐                   │
│  │        Embeddings                    │                   │
│  │  (Convert to vectors)                │                   │
│  └──────────────┬──────────────────────┘                   │
│                 │                                            │
│                 ▼                                            │
│  ┌─────────────────────────────────────┐                   │
│  │     Supabase Vector Store           │                   │
│  │  - pgvector for embeddings          │                   │
│  │  - Metadata tables                  │                   │
│  └─────────────────────────────────────┘                   │
│                 │                                            │
│                 ▼                                            │
│  ┌─────────────────────────────────────┐                   │
│  │      RAG Chain (LangGraph)          │                   │
│  │  1. Query Analysis                  │                   │
│  │  2. Retrieval                       │                   │
│  │  3. Generation                      │                   │
│  └─────────────────────────────────────┘                   │
│                                                              │
└─────────────────────────────────────────────────────────────┘
```

## 주요 컴포넌트

### 1. Document Loaders
```python
# 내부 문서 로더
- PDFLoader
- DocxLoader  
- UnstructuredPowerPointLoader
- Custom EmailLoader (Outlook .msg)

# 외부 뉴스 로더
- WebBaseLoader
- Custom NewsAPILoader
```

### 2. Text Processing
```python
# 한국어 최적화 Splitter
- RecursiveCharacterTextSplitter
- Custom metadata extraction
- 문서 유형별 chunk 전략
```

### 3. Vector Store (Supabase)
```python
# pgvector 테이블 구조
- documents (원본 문서 메타데이터)
- chunks (분할된 텍스트 청크)
- embeddings (벡터 임베딩)
- correlations (내부-외부 연관도)
```

### 4. RAG Chain Components
```python
# LangGraph State
class STRIXState(TypedDict):
    question: str
    query: dict  # 구조화된 쿼리
    internal_docs: List[Document]
    external_news: List[Document]
    correlation_score: float
    answer: str
    metadata: dict
```

## 핵심 기능 구현

### 1. 통합 검색 (Hybrid Search)
- 내부 문서와 외부 뉴스 동시 검색
- 메타데이터 필터링 (조직, 기간, 카테고리)
- 의미적 유사도 + 키워드 매칭

### 2. 지능형 연관도 분석
- 내부 이슈와 외부 뉴스 자동 매핑
- LLM 기반 연관도 점수 계산
- 시계열 분석

### 3. 인사이트 생성
- 맥락 기반 요약
- 트렌드 분석
- 리스크 조기 감지

## Supabase 스키마

```sql
-- 문서 메타데이터
CREATE TABLE documents (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    type VARCHAR(50), -- 'internal' or 'external'
    source VARCHAR(255),
    title TEXT,
    organization VARCHAR(100),
    category VARCHAR(100),
    created_at TIMESTAMP,
    metadata JSONB
);

-- 텍스트 청크
CREATE TABLE chunks (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    document_id UUID REFERENCES documents(id),
    content TEXT,
    chunk_index INTEGER,
    metadata JSONB
);

-- 벡터 임베딩
CREATE TABLE embeddings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    chunk_id UUID REFERENCES chunks(id),
    embedding vector(1536)
);

-- 연관도
CREATE TABLE correlations (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    internal_doc_id UUID REFERENCES documents(id),
    external_doc_id UUID REFERENCES documents(id),
    score FLOAT,
    reasoning TEXT,
    created_at TIMESTAMP DEFAULT NOW()
);
```

## 구현 로드맵

### Phase 1: 기본 인프라 (1주)
- Supabase 설정 및 스키마 생성
- 기본 Document Loader 구현
- Embedding 파이프라인 구축

### Phase 2: RAG Chain 구현 (1주)
- LangGraph 기반 검색/생성 체인
- Query Analysis 구현
- Hybrid Retrieval 구현

### Phase 3: 고급 기능 (1주)
- 연관도 분석 알고리즘
- 시계열 대시보드
- Newsletter 자동 생성

### Phase 4: 통합 및 최적화 (1주)
- API 서버 구축
- 프론트엔드 연동
- 성능 최적화

## 장점

1. **확장성**: 클라우드 기반으로 대용량 처리 가능
2. **정확도**: LLM 기반 의미 이해로 정확한 검색
3. **자동화**: 수동 태깅 최소화
4. **실시간성**: 스트리밍 업데이트 지원
5. **유연성**: 다양한 문서 형식 지원

## 사용 예시

```python
# 질의 예시
"지난 달 배터리 관련 내부 보고서와 외부 규제 뉴스 연관성 분석"
"전고체 배터리 개발 현황과 경쟁사 동향 비교"
"ESG 관련 경영진 관심사항과 최근 정책 변화"
```

이제 이 아키텍처를 기반으로 실제 코드를 구현해나가면 됩니다.