"""
ESS 사업 관련 시연용 문서 업로드 (간단 버전)
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from datetime import datetime, timedelta
import uuid
from database.supabase_client import SupabaseClient
from langchain_openai import OpenAIEmbeddings
from langchain.text_splitter import RecursiveCharacterTextSplitter
import asyncio

# 텍스트 분할기
text_splitter = RecursiveCharacterTextSplitter(
    chunk_size=1000,
    chunk_overlap=200,
    separators=["\n\n", "\n", ". ", " ", ""]
)

# 임베딩 모델
embeddings = OpenAIEmbeddings()

# ESS 관련 문서들 (간소화)
ess_documents = [
    # 내부 문서
    {
        "type": "internal",
        "title": "ESS 사업 진출 전략 종합 보고서",
        "organization": "전략기획",
        "category": "전략기획",
        "content": """ESS 사업 진출 전략 종합 보고서

I. Executive Summary
- 글로벌 ESS 시장 연평균 35% 성장 전망 (2024-2030)
- 당사 배터리 기술력 활용한 시장 진입 기회
- 2027년까지 글로벌 Top 5 진입 목표

II. 시장 분석
1. 시장 규모: 2024년 200억 달러 → 2030년 1,200억 달러
2. 주요 경쟁사: Tesla Energy, CATL, LG에너지솔루션, 삼성SDI
3. 성장 동력: 재생에너지 확산, 전력망 안정화 수요

III. 진출 전략
- Phase 1 (2025): 국내 시장 기반 구축
- Phase 2 (2026): 아시아 시장 확대  
- Phase 3 (2027): 글로벌 시장 본격 진출
- 핵심 차별화: 고안전성 LFP 기술, AI 에너지 관리""",
        "created_at": datetime.now() - timedelta(days=15)
    },
    {
        "type": "internal",
        "title": "ESS 기술개발 로드맵 회의록",
        "organization": "R&D센터",
        "category": "R&D",
        "content": """ESS 기술개발 전략 회의록

I. LFP 배터리 고도화
- 에너지밀도 200Wh/kg 목표
- 수명 10,000 사이클 달성
- -30°C 저온 성능 개선

II. 안전성 기술
- 3중 안전 시스템 개발
- UL9540A 인증 획득 추진
- AI 기반 이상 감지 시스템

III. 차세대 기술
- 나트륨이온 배터리 ESS 개발
- 전고체 배터리 적용 검토
- 2025년 R&D 예산: 3,000억원""",
        "created_at": datetime.now() - timedelta(days=10)
    },
    {
        "type": "internal", 
        "title": "ESS 화재 리스크 관리 방안",
        "organization": "경영지원",
        "category": "경영지원",
        "content": """ESS 화재 리스크 종합 관리 방안

I. 화재 원인 분석
- 배터리 자체 요인 40%
- 시스템 요인 35%
- 운영 요인 25%

II. 대응 방안
1. 기술적: LFP 100% 적용, 4단계 안전 시스템
2. 품질: 전수 검사, AI 불량 예측
3. 운영: 24/7 모니터링, 긴급 대응팀
4. 보험: 제조물 책임보험, 10년 성능 보증

III. 투자 계획
- 안전성 R&D: 연 500억원
- 화재 사고율 목표: 0.001% 이하""",
        "created_at": datetime.now() - timedelta(days=7)
    },
    {
        "type": "internal",
        "title": "ESS 글로벌 영업 전략",
        "organization": "영업마케팅", 
        "category": "영업마케팅",
        "content": """ESS 글로벌 영업 전략

I. 목표 시장
- Tier 1: 미국(IRA), 유럽(RE100), 호주
- Tier 2: 일본, 인도, 중동

II. 고객 세그먼트
- Utility: 100MWh 이상 대형
- C&I: 1-10MWh 중형
- Residential: 10-30kWh 소형

III. 2025년 목표
- 수주: 5GWh
- 매출: 3조원
- 신규 고객: 50개사
- 시장점유율: 5%""",
        "created_at": datetime.now() - timedelta(days=5)
    },
    # 외부 문서
    {
        "type": "external",
        "title": "[속보] 미국, ESS 투자세액공제 40%로 확대",
        "organization": "에너지경제",
        "category": "정책",
        "content": """미국 정부가 ESS 투자 세액공제를 30%에서 40%로 확대 발표.

주요 내용:
- 기본 세액공제: 40%
- 미국산 배터리: +10%
- 최대 세액공제: 60%

시장 전망:
- 미국 ESS 시장 연 50% 성장
- 2030년까지 200GWh 수요
- 한국 기업들 수혜 예상""",
        "created_at": datetime.now() - timedelta(days=2)
    },
    {
        "type": "external",
        "title": "[글로벌] CATL, 사우디 10GWh ESS 프로젝트 수주",
        "organization": "배터리인사이트",
        "category": "경쟁사",
        "content": """CATL이 사우디에서 세계 최대 10GWh ESS 프로젝트 수주.

프로젝트 개요:
- 규모: 10GWh
- 금액: 50억 달러
- 기간: 2025-2027년

경쟁사 동향:
- BYD: 중동 5GWh 추진
- 테슬라: Megapack 확대
- LG엔솔: 미국 집중
- 삼성SDI: 유럽 공략""",
        "created_at": datetime.now() - timedelta(days=4)
    },
    {
        "type": "external",
        "title": "[리포트] 글로벌 ESS 시장 2030년 1,200억 달러",
        "organization": "마켓리서치",
        "category": "산업",
        "content": """글로벌 ESS 시장 급성장 전망.

시장 규모:
- 2024년: 200억 달러
- 2030년: 1,200억 달러
- CAGR: 35%

성장 동력:
- 재생에너지 확산
- 전력망 현대화
- 전기요금 상승

지역별:
- 북미 35%
- 유럽 25%
- 중국 20%""",
        "created_at": datetime.now() - timedelta(days=6)
    }
]

async def upload_documents():
    """문서를 Supabase에 업로드"""
    client = SupabaseClient()
    
    print(f"총 {len(ess_documents)}개 ESS 문서 업로드 시작...")
    
    for i, doc in enumerate(ess_documents, 1):
        try:
            # 문서 삽입
            doc_data = {
                "type": doc["type"],
                "title": doc["title"],
                "organization": doc["organization"],
                "category": doc["category"],
                "created_at": doc["created_at"].isoformat(),
                "source": "ESS Demo",
                "file_path": f"demo/ess_{i}.txt",
                "metadata": {"topic": "ESS"}
            }
            
            doc_response = client.client.table("documents").insert(doc_data).execute()
            
            if doc_response.data:
                document_id = doc_response.data[0]["id"]
                print(f"[{i}] {doc['title'][:30]}... 업로드 완료")
                
                # 청크 생성
                chunks = text_splitter.split_text(doc["content"])
                
                for j, chunk_text in enumerate(chunks):
                    # 청크 삽입
                    chunk_data = {
                        "document_id": document_id,
                        "content": chunk_text,
                        "chunk_index": j,
                        "metadata": {"chunk_num": j+1}
                    }
                    
                    chunk_response = client.client.table("chunks").insert(chunk_data).execute()
                    
                    if chunk_response.data:
                        chunk_id = chunk_response.data[0]["id"]
                        
                        # 임베딩 생성
                        embedding = await embeddings.aembed_query(chunk_text)
                        
                        # 임베딩 저장
                        embedding_data = {
                            "chunk_id": chunk_id,
                            "embedding": embedding
                        }
                        
                        client.client.table("embeddings").insert(embedding_data).execute()
                
                print(f"  → {len(chunks)}개 청크 및 임베딩 생성 완료")
                
        except Exception as e:
            print(f"오류: {doc['title'][:30]}... - {str(e)}")
            continue
    
    print("\n✅ ESS 시연 데이터 업로드 완료!")

if __name__ == "__main__":
    asyncio.run(upload_documents())