"""
Create more competitive analysis related documents
경쟁사 분석 관련 추가 문서 생성
"""
import sys
import os
from datetime import datetime, timedelta
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from langchain_openai import OpenAIEmbeddings
from langchain.text_splitter import RecursiveCharacterTextSplitter

# 추가 경쟁사 관련 문서
competitive_documents = [
    {
        "title": "글로벌 배터리 경쟁 구도 심층 분석",
        "organization": "전략기획",
        "category": "전략기획",
        "type": "internal",
        "content": """
글로벌 배터리 경쟁 구도 심층 분석 (2025.1)

I. 경쟁 구도 변화
1. 시장 집중도
   - CR3: 67% (CATL 37%, BYD 16%, LG에너지 14%)
   - CR5: 78% (+ 파나소닉 6%, SK온 5%)
   - HHI: 2,100 (과점시장 심화)

2. 지역별 경쟁 양상
   - 중국: CATL/BYD 양강 구도
   - 한국: LG/삼성/SK 3사 경쟁
   - 일본: 파나소닉 독주, 도요타 진입
   - 미국: 테슬라 자체생산 확대

II. CATL 대응 전략
1. 강점 분석
   - 시장점유율 37% 압도적 1위
   - 원가경쟁력: 당사 대비 25% 저렴
   - 중국 내수시장 80% 장악
   - LFP 기술 선도

2. 약점 분석
   - 지정학적 리스크 (미중 갈등)
   - 품질 안정성 이슈
   - 프리미엄 시장 열세
   - 해외 생산기지 부족

3. 대응 방안
   - 기술 차별화: 안전성, 수명
   - 프리미엄 시장 집중
   - 미국/유럽 현지화
   - 차세대 기술 선점

III. BYD 대응 전략
1. 특징
   - 수직계열화 완성
   - 완성차 시너지
   - Blade Battery 차별화

2. 대응
   - 고객 다변화
   - 모듈화 기술
   - 안전성 강조

IV. LG에너지솔루션 대비
1. 당사 대비 우위
   - 규모: 3배
   - 글로벌 고객: 2배
   - 수익성: 영업이익률 7%

2. 추격 전략
   - 차별화 기술
   - 틈새시장
   - 원가 개선
        """
    },
    {
        "title": "주요 경쟁사 기술 로드맵 분석",
        "organization": "R&D",
        "category": "R&D",
        "type": "internal",
        "content": """
주요 경쟁사 기술 로드맵 분석 (2025.1)

I. CATL 기술 전략
1. 현재 기술
   - CTP 3.0: 셀투팩 직접 연결
   - Qilin: 에너지밀도 255Wh/kg
   - 급속충전: 10분 80%

2. 개발 중
   - CTC: 셀투샤시
   - M3P: 망간 기반 배터리
   - 응축물질 배터리

3. 2027 목표
   - 에너지밀도: 500Wh/kg
   - 충전: 5분 80%
   - 수명: 100만km

II. BYD 기술 전략
1. Blade Battery
   - LFP 기반
   - 안전성 극대화
   - CTB 기술

2. 차세대
   - 나트륨이온
   - 고망간
   - 전고체 2030

III. 테슬라 4680
1. 현재
   - 자체 생산 50%
   - 원가 50% 절감 목표
   - 구조용 배터리

2. 계획
   - 연 1TWh 생산
   - 건식 전극
   - 실리콘 음극

IV. 일본 전고체 연합
1. 도요타
   - 2027 양산
   - 10분 충전
   - 1,200km 주행

2. 파나소닉
   - 2028 목표
   - 에너지밀도 2배
   - 소형 EV 적용

V. 당사 차별화 포인트
1. 전고체
   - 2027 양산
   - 독자 기술
   - 파트너십

2. 리튬메탈
   - 2026 파일럿
   - 450Wh/kg
   - 안전성 확보

3. AI 기반 BMS
   - 수명 예측
   - 최적 제어
   - 고장 진단
        """
    },
    {
        "title": "경쟁사별 고객 포트폴리오 분석",
        "organization": "영업마케팅",
        "category": "영업마케팅",
        "type": "internal",
        "content": """
경쟁사별 고객 포트폴리오 분석 (2025.1)

I. CATL 고객
1. 주요 고객
   - Tesla: 30% (Model 3/Y LFP)
   - NIO/Xpeng/Li Auto: 25%
   - VW Group: 15%
   - BMW/Daimler: 10%
   - 기타: 20%

2. 특징
   - 중국 OEM 의존도 55%
   - Tesla 의존도 감소 추세
   - 유럽 OEM 확대

II. LG에너지솔루션 고객
1. 주요 고객
   - GM: 25%
   - Stellantis: 20%
   - VW: 15%
   - 현대기아: 15%
   - Ford: 10%

2. 특징
   - 미국 OEM 45%
   - 유럽 OEM 35%
   - 장기계약 중심

III. BYD 고객
1. 자체 브랜드: 70%
2. 외부 공급: 30%
   - Toyota
   - 중국 로컬

IV. 당사 고객 전략
1. 현재
   - Top 3: 70%
   - 장기계약: 60%

2. 목표
   - 고객 다변화
   - 신규 10개사
   - 의존도 50%

V. 고객 확보 전략
1. 차별화
   - 품질/안전성
   - 맞춤형 개발
   - 기술 지원

2. 서비스
   - JDP 강화
   - A/S 체계
   - 공급 안정성

3. 가격
   - TCO 어필
   - 장기계약 할인
   - 볼륨 인센티브
        """
    },
    {
        "title": "[업계동향] 2025 배터리 업계 M&A 동향",
        "organization": "증권사리서치",
        "category": "산업",
        "type": "external",
        "content": """
[업계동향] 2025년 글로벌 배터리 업계 M&A 전망

□ 주요 M&A 동향
ㅇ CATL, 볼리비아 리튬 기업 인수 (30억달러)
ㅇ Stellantis, 배터리 스타트업 인수 (15억달러)
ㅇ Ford, LFP 특허 기업 인수 (5억달러)
ㅇ 중국 자본, 인니 니켈 프로젝트 투자 (20억달러)

□ M&A 트렌드
1. 수직계열화
   - 원자재 확보 경쟁
   - 상류 부문 투자 급증
   - 리튬/니켈/코발트 기업 인수

2. 기술 확보
   - 차세대 기술 스타트업
   - 특허 포트폴리오
   - 인재 영입 목적

3. 시장 진입
   - 현지 기업 인수
   - JV → 인수 전환
   - 생산설비 확보

□ 2025년 예상 딜
ㅇ 한국 배터리사, 미국 현지기업 인수
ㅇ 일본 기업, 전고체 스타트업 투자
ㅇ 중국 기업, 유럽 생산기지 인수
ㅇ OEM, 배터리 기업 지분 투자

□ 밸류에이션
- EV/EBITDA: 8-12배
- 기술 기업: 15-20배
- 자원 기업: 10-15배
        """
    },
    {
        "title": "[산업리포트] 한중일 배터리 3국 경쟁",
        "organization": "산업연구원",
        "category": "산업",
        "type": "external",
        "content": """
[산업리포트] 한중일 배터리 3국 경쟁 구도

□ 국가별 점유율 (2025.1)
ㅇ 중국: 65% (CATL, BYD, CALB 등)
ㅇ 한국: 23% (LG, 삼성, SK)
ㅇ 일본: 7% (파나소닉, AESC)
ㅇ 기타: 5%

□ 중국 전략
1. 강점
   - 규모의 경제
   - 원가 경쟁력
   - 내수 시장
   - 정부 지원

2. 전략
   - 글로벌 확장
   - 기술 추격
   - 자원 확보

□ 한국 전략
1. 강점
   - 기술력
   - 품질
   - 글로벌 고객

2. 전략
   - 차별화
   - 미국/유럽
   - 차세대 기술

□ 일본 전략
1. 강점
   - 전고체 기술
   - 소재 기술
   - 품질 관리

2. 전략
   - 전고체 선점
   - 도요타 연합
   - 프리미엄

□ 경쟁 전망
- 중국 독주 지속
- 한국 차별화 생존
- 일본 전고체 승부
- 지역 블록화 가속
        """
    },
    {
        "title": "[경쟁사뉴스] 삼성SDI 전고체 배터리 돌파구",
        "organization": "테크미디어",
        "category": "경쟁사",
        "type": "external",
        "content": """
[경쟁사뉴스] 삼성SDI, 전고체 배터리 기술 돌파구 마련

□ 기술 혁신
ㅇ 황화물계 고체전해질 개발 성공
ㅇ 이온전도도 10mS/cm 달성
ㅇ 계면 저항 50% 감소
ㅇ 2027년 양산 목표 유지

□ 투자 계획
ㅇ 전고체 R&D: 연 5천억원
ㅇ 파일럿 라인: 2026년 완공
ㅇ 양산 라인: 2027년 10GWh

□ 고객사 반응
- 현대차: 2027년 제네시스 적용 검토
- BMW: 차세대 전기차 협력
- Stellantis: 공동개발 협의

□ 경쟁사 대응
- LG에너지: R&D 투자 확대
- SK온: 솔리드파워와 협력
- CATL: 2030년 목표 유지

□ 시장 영향
- 한국 기업 기술 리더십
- 전고체 상용화 가속
- 프리미엄 시장 재편
        """
    },
    {
        "title": "[특별분석] Northvolt 위기와 유럽 배터리 전략",
        "organization": "유럽리서치",
        "category": "경쟁사",
        "type": "external",
        "content": """
[특별분석] Northvolt 경영위기가 유럽 배터리 산업에 미치는 영향

□ Northvolt 현황
ㅇ 자금 부족: 50억유로 필요
ㅇ 생산 지연: 목표 대비 30%
ㅇ 품질 이슈: 수율 60% 수준
ㅇ 구조조정: 인력 20% 감축

□ 원인 분석
1. 과도한 확장
   - 동시 다발 투자
   - 기술력 부족
   - 경험 부재

2. 시장 환경
   - 수요 둔화
   - 가격 경쟁
   - 중국 기업 공세

□ 유럽 배터리 전략 재검토
ㅇ EU 자립도 목표 하향 (90% → 70%)
ㅇ 아시아 기업 협력 불가피
ㅇ 정부 지원 확대 필요

□ 한국 기업 기회
- 유럽 진출 호기
- 현지 파트너십
- 기술 이전/라이선싱
- M&A 기회

□ 시사점
- 기술력 없는 진입 한계
- 규모의 경제 필수
- 정부 지원 한계
- 아시아 의존 지속
        """
    }
]

def upload_competitive_docs():
    """경쟁사 관련 추가 문서 업로드"""
    client = SupabaseClient()
    embeddings = OpenAIEmbeddings()
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=200,
        separators=["\n\n", "\n", ".", "!", "?", ",", " ", ""],
        length_function=len
    )
    
    print("=== Uploading Additional Competitive Analysis Documents ===\n")
    
    for i, doc_data in enumerate(competitive_documents, 1):
        try:
            print(f"[{i}/{len(competitive_documents)}] Processing: {doc_data['title']}")
            
            # 문서 삽입
            doc_record = {
                'type': doc_data['type'],
                'source': doc_data.get('organization', ''),
                'title': doc_data['title'],
                'organization': doc_data.get('organization', ''),
                'category': doc_data.get('category', ''),
                'created_at': datetime.now().strftime("%Y-%m-%d"),
                'file_path': f"competitive_docs/{doc_data['title']}.txt",
                'metadata': {
                    'focus': 'competitive_analysis',
                    'importance': 'high'
                }
            }
            
            doc_result = client.client.table('documents').insert(doc_record).execute()
            doc_id = doc_result.data[0]['id']
            
            # 청크 생성 및 임베딩
            chunks = text_splitter.split_text(doc_data['content'])
            
            for j, chunk_text in enumerate(chunks):
                # 청크 삽입
                chunk_record = {
                    'document_id': doc_id,
                    'content': chunk_text,
                    'chunk_index': j,
                    'metadata': {
                        'title': doc_data['title'],
                        'type': doc_data['type'],
                        'focus': 'competitive'
                    }
                }
                
                chunk_result = client.client.table('chunks').insert(chunk_record).execute()
                chunk_id = chunk_result.data[0]['id']
                
                # 임베딩 생성 및 삽입
                embedding = embeddings.embed_query(chunk_text)
                
                embedding_record = {
                    'chunk_id': chunk_id,
                    'embedding': embedding,
                    'model': 'text-embedding-ada-002'
                }
                
                client.client.table('embeddings').insert(embedding_record).execute()
            
            print(f"  Created {len(chunks)} chunks with embeddings")
            
        except Exception as e:
            print(f"  Error: {e}")
    
    print(f"\n=== Complete ===")
    print(f"Uploaded {len(competitive_documents)} competitive analysis documents")

if __name__ == "__main__":
    upload_competitive_docs()