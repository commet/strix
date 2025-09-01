"""
Update document dates to be more diverse
문서 날짜를 다양하게 업데이트
"""
import sys
import os
from datetime import datetime, timedelta
import random
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient

def update_document_dates():
    """문서 날짜를 다양하게 업데이트"""
    client = SupabaseClient()
    
    print("=== Updating Document Dates ===\n")
    
    # 모든 문서 가져오기
    docs = client.client.table('documents').select('id, title, created_at').execute()
    
    print(f"Found {len(docs.data)} documents to update\n")
    
    # 날짜 생성 - 2024년 1월부터 2025년 1월까지
    base_date = datetime(2024, 1, 1)
    date_range = 365  # 1년 범위
    
    # 각 문서마다 다른 날짜 할당
    for i, doc in enumerate(docs.data):
        # 문서 제목에 따라 날짜 패턴 설정
        title = doc['title']
        
        if '2025' in title:
            # 2025년 관련 문서는 최근 날짜
            days_offset = random.randint(330, 365)
        elif '2024' in title:
            # 2024년 관련 문서는 중간 날짜
            days_offset = random.randint(180, 330)
        elif 'Q1' in title or '1분기' in title:
            # 1분기 문서
            days_offset = random.randint(0, 90)
        elif 'Q2' in title or '2분기' in title:
            # 2분기 문서
            days_offset = random.randint(90, 180)
        elif 'Q3' in title or '3분기' in title:
            # 3분기 문서
            days_offset = random.randint(180, 270)
        elif 'Q4' in title or '4분기' in title:
            # 4분기 문서
            days_offset = random.randint(270, 365)
        else:
            # 랜덤 날짜
            days_offset = random.randint(0, date_range)
        
        new_date = base_date + timedelta(days=days_offset)
        
        # 주말이면 평일로 조정
        if new_date.weekday() >= 5:  # 토요일(5) 또는 일요일(6)
            if new_date.weekday() == 5:  # 토요일
                new_date -= timedelta(days=1)  # 금요일로
            else:  # 일요일
                new_date += timedelta(days=1)  # 월요일로
        
        # 업데이트
        try:
            client.client.table('documents').update({
                'created_at': new_date.strftime('%Y-%m-%d %H:%M:%S')
            }).eq('id', doc['id']).execute()
            
            print(f"Updated: {doc['title'][:50]}...")
            print(f"  New date: {new_date.strftime('%Y-%m-%d')}")
            
        except Exception as e:
            print(f"Error updating {doc['title']}: {e}")
    
    print("\n=== Date Update Complete ===")
    
    # 업데이트 후 날짜 분포 확인
    check_date_distribution(client)

def check_date_distribution(client):
    """날짜 분포 확인"""
    docs = client.client.table('documents').select('created_at').execute()
    
    date_counts = {}
    for doc in docs.data:
        date_str = doc['created_at'][:7]  # YYYY-MM 형식
        date_counts[date_str] = date_counts.get(date_str, 0) + 1
    
    print("\n날짜 분포:")
    for date_str in sorted(date_counts.keys()):
        print(f"  {date_str}: {date_counts[date_str]} documents")

def add_more_recent_documents():
    """최근 날짜 문서 추가"""
    from langchain_openai import OpenAIEmbeddings
    from langchain.text_splitter import RecursiveCharacterTextSplitter
    
    client = SupabaseClient()
    embeddings = OpenAIEmbeddings()
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=200
    )
    
    # 최신 뉴스/보고서 형태의 문서들
    recent_docs = [
        {
            "title": "[긴급] 2025년 1월 리튬 가격 급등",
            "organization": "원자재분석팀",
            "category": "Macro",
            "type": "external",
            "date": datetime(2025, 1, 15),
            "content": """
[긴급속보] 리튬 가격 단기 급등 예상

□ 주요 내용
ㅇ 칠레 폭우로 주요 리튬 광산 생산 중단
ㅇ 탄산리튬 현물가 $20,000/톤 돌파
ㅇ 향후 3개월간 공급 부족 예상

□ 영향 분석
- 배터리 제조원가 10-15% 상승 압박
- 장기계약 재협상 불가피
- 재고 확보 경쟁 심화

□ 대응 방안
- 즉시 재고 확보 (3개월분)
- 대체 공급처 긴급 확보
- 가격 헤징 포지션 구축
            """
        },
        {
            "title": "2024년 12월 월간 경영실적",
            "organization": "재무팀",
            "category": "경영지원",
            "type": "internal",
            "date": datetime(2025, 1, 5),
            "content": """
2024년 12월 경영실적 보고

I. 매출 실적
- 월 매출: 1.5조원 (YoY +25%)
- 누적 매출: 15.2조원 (목표 대비 101%)

II. 수익성
- 영업이익: 800억원
- 영업이익률: 5.3%
- EBITDA: 1,200억원

III. 주요 성과
- 미국 공장 가동률 85% 달성
- 신규 고객 2개사 확보
- 품질 불량률 0.5% 이하 유지
            """
        },
        {
            "title": "[주간리포트] 2025년 1월 2주차 시장동향",
            "organization": "시장분석팀",
            "category": "산업",
            "type": "external",
            "date": datetime(2025, 1, 10),
            "content": """
주간 배터리 시장 동향 (2025.1.6-1.10)

□ 주요 이슈
1. Tesla, 신형 Model 3 배터리 공급사 변경 검토
2. EU, 배터리 규제 추가 강화 발표
3. 인도 정부, 배터리 현지화 의무화

□ 경쟁사 동향
- CATL: 유럽 3공장 착공
- BYD: 인도 진출 본격화
- LG에너지: 북미 증설 발표

□ 시장 전망
- 단기: 수요 회복세 지속
- 중기: 공급과잉 우려 상존
            """
        },
        {
            "title": "2024년 11월 R&D 진척 보고",
            "organization": "R&D센터",
            "category": "R&D",
            "type": "internal",
            "date": datetime(2024, 12, 1),
            "content": """
2024년 11월 R&D 주요 진척사항

I. 전고체 배터리
- 시제품 성능 테스트 완료
- 에너지밀도 380Wh/kg 달성
- 안정성 테스트 통과

II. 차세대 양극재
- NCM9.5.5 조성 최적화
- 수명 3,000사이클 확인
- 원가 10% 절감 달성

III. 특허 출원
- 국내 특허 5건
- PCT 출원 3건
- 핵심특허 등록 2건
            """
        },
        {
            "title": "[속보] 미국 IRA 세부규정 변경",
            "organization": "정책분석실",
            "category": "정책",
            "type": "external",
            "date": datetime(2024, 12, 20),
            "content": """
미국 IRA 세부 규정 변경 발표

□ 변경 내용
ㅇ 배터리 부품 요건 완화 (70% → 65%)
ㅇ 적용 시기 6개월 연기
ㅇ 한국 기업 추가 혜택

□ 예상 영향
- 한국 배터리 3사 수혜
- 투자 계획 재조정 필요
- 중국 견제 지속

□ 대응 방향
- 미국 생산 계획 유지
- 현지 파트너십 강화
- 정책 모니터링 지속
            """
        }
    ]
    
    print("\n=== Adding Recent Documents ===")
    
    for doc_data in recent_docs:
        try:
            print(f"Adding: {doc_data['title']}")
            
            # 문서 삽입
            doc_record = {
                'type': doc_data['type'],
                'source': doc_data['organization'],
                'title': doc_data['title'],
                'organization': doc_data['organization'],
                'category': doc_data['category'],
                'created_at': doc_data['date'].strftime('%Y-%m-%d %H:%M:%S'),
                'file_path': f"recent/{doc_data['title']}.txt",
                'metadata': {'recent': True}
            }
            
            doc_result = client.client.table('documents').insert(doc_record).execute()
            doc_id = doc_result.data[0]['id']
            
            # 청크 생성 및 임베딩
            chunks = text_splitter.split_text(doc_data['content'])
            
            for j, chunk_text in enumerate(chunks):
                chunk_record = {
                    'document_id': doc_id,
                    'content': chunk_text,
                    'chunk_index': j,
                    'metadata': {'title': doc_data['title']}
                }
                
                chunk_result = client.client.table('chunks').insert(chunk_record).execute()
                chunk_id = chunk_result.data[0]['id']
                
                embedding = embeddings.embed_query(chunk_text)
                
                embedding_record = {
                    'chunk_id': chunk_id,
                    'embedding': embedding,
                    'model': 'text-embedding-ada-002'
                }
                
                client.client.table('embeddings').insert(embedding_record).execute()
            
            print(f"  Added with {len(chunks)} chunks")
            
        except Exception as e:
            print(f"  Error: {e}")

if __name__ == "__main__":
    # 기존 문서 날짜 업데이트
    update_document_dates()
    
    # 최신 문서 추가
    print("\n" + "="*50)
    add_more_recent_documents()