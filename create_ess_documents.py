"""
ESS 사업 관련 시연용 문서 생성 스크립트
경영진 시연을 위한 ESS(Energy Storage System) 중심 데이터
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from datetime import datetime, timedelta
import random
from database.supabase_client import SupabaseClient
from database.document_ingester import DocumentIngester
import asyncio

# ESS 관련 내부 문서들
internal_ess_documents = [
    {
        "type": "internal",
        "title": "ESS 사업 진출 전략 보고서",
        "organization": "전략기획",
        "category": "전략기획",
        "content": """
ESS 사업 진출 전략 종합 보고서

I. Executive Summary
- 글로벌 ESS 시장 연평균 35% 성장 전망 (2024-2030)
- 당사 배터리 기술력 활용한 시장 진입 기회
- 2027년까지 글로벌 Top 5 진입 목표

II. 시장 분석
1. 시장 규모 및 성장성
   - 2024년: 200억 달러 → 2030년: 1,200억 달러
   - 재생에너지 확산이 핵심 성장 동력
   - 전력망 안정화 수요 급증

2. 주요 경쟁사 현황
   - Tesla Energy: Megapack 중심 대형 ESS
   - CATL: 중국 내수 기반 급성장
   - LG에너지솔루션: 북미 시장 공략
   - 삼성SDI: 프리미엄 시장 타겟

III. 당사 진출 전략
1. 단계별 접근
   - Phase 1 (2025): 국내 시장 기반 구축
   - Phase 2 (2026): 아시아 시장 확대
   - Phase 3 (2027): 글로벌 시장 본격 진출

2. 핵심 차별화 요소
   - 고안전성 LFP 배터리 기술
   - AI 기반 에너지 관리 시스템
   - 모듈러 설계로 확장성 극대화

3. 투자 계획
   - 총 투자액: 5조원 (2025-2027)
   - ESS 전용 공장 2개소 건설
   - R&D 센터 설립

IV. 리스크 관리
- 화재 안전성 확보 최우선
- 원자재 가격 변동 헤징
- 정책 변화 대응 체계 구축

V. 예상 성과
- 2027년 매출 10조원 달성
- 영업이익률 15% 이상
- 시장점유율 8% 확보
        """,
        "created_at": datetime.now() - timedelta(days=15)
    },
    {
        "type": "internal",
        "title": "ESS 기술개발 로드맵 회의록",
        "organization": "R&D센터",
        "category": "R&D",
        "content": """
ESS 기술개발 전략 회의록
일시: 2024년 11월 20일
참석: CTO, R&D 센터장, ESS 개발팀

I. 회의 안건
- ESS용 차세대 배터리 기술 개발 방향
- 안전성 강화 방안
- 경쟁사 기술 대응 전략

II. 주요 논의사항

1. LFP 배터리 고도화
   - 에너지밀도 200Wh/kg 목표 (현재 180Wh/kg)
   - 수명 10,000 사이클 달성
   - -30°C 저온 성능 개선

2. 안전성 기술
   - 3중 안전 시스템 개발
     * 셀 레벨: 과충전 방지 회로
     * 모듈 레벨: 열확산 방지 구조
     * 시스템 레벨: AI 이상 감지
   - UL9540A 인증 획득 추진

3. BMS 고도화
   - AI 기반 SOH 예측 정확도 95% 이상
   - 실시간 열관리 최적화
   - 클라우드 연동 원격 모니터링

4. 차세대 기술
   - 나트륨이온 배터리 ESS 개발
   - 전고체 배터리 ESS 적용 검토
   - 리튬메탈 음극 연구

III. Action Items
1. Q1 2025: LFP Gen2 프로토타입 완성
2. Q2 2025: 안전성 인증 테스트
3. Q3 2025: 파일럿 프로젝트 착수
4. Q4 2025: 양산 준비

IV. 예산
- 2025년 R&D 예산: 3,000억원
- 인력 충원: 200명
- 테스트 설비 투자: 500억원
        """,
        "created_at": datetime.now() - timedelta(days=10)
    },
    {
        "type": "internal",
        "title": "ESS 화재 리스크 관리 방안",
        "organization": "경영지원",
        "category": "경영지원",
        "content": """
ESS 화재 리스크 종합 관리 방안

I. 배경
- 국내외 ESS 화재 사고 지속 발생
- 안전성이 시장 진출의 핵심 성공 요인
- 보험료 및 배상 리스크 증가

II. 화재 원인 분석
1. 배터리 자체 요인 (40%)
   - 셀 불량 및 내부 단락
   - 열폭주 연쇄 반응
   - 제조 공정 품질 이슈

2. 시스템 요인 (35%)
   - BMS 오작동
   - 냉각 시스템 고장
   - 전기적 절연 불량

3. 운영 요인 (25%)
   - 과충전/과방전
   - 부적절한 유지보수
   - 환경 조건 미준수

III. 대응 방안

1. 기술적 대응
   - LFP 배터리 100% 적용 (열안정성 우수)
   - 4단계 안전 시스템 구축
   - 실시간 모니터링 강화
   - 자동 소화 시스템 탑재

2. 품질 관리
   - 전수 검사 체계 도입
   - AI 기반 불량 예측
   - 공급망 품질 관리 강화

3. 운영 관리
   - 24/7 원격 모니터링 센터
   - 정기 점검 의무화
   - 긴급 대응 팀 운영

4. 보험 및 보증
   - 제조물 책임보험 가입
   - 10년 성능 보증
   - 화재 배상 책임 한도 설정

IV. 인증 및 규제 대응
- KC, UL, IEC 안전 인증
- 소방청 화재안전기준 준수
- 정부 ESS 안전강화 대책 이행

V. 투자 계획
- 안전성 R&D: 연 500억원
- 테스트 설비: 300억원
- 모니터링 시스템: 200억원

VI. 기대 효과
- 화재 사고율 0.001% 이하
- 보험료 30% 절감
- 고객 신뢰도 향상
        """,
        "created_at": datetime.now() - timedelta(days=7)
    },
    {
        "type": "internal",
        "title": "ESS 글로벌 영업 전략",
        "organization": "영업마케팅",
        "category": "영업마케팅",
        "content": """
ESS 글로벌 영업 전략 수립

I. 목표 시장 분석

1. Tier 1 시장 (우선 진출)
   - 미국: IRA 혜택, 전력망 노후화
   - 유럽: RE100, 탄소중립 정책
   - 호주: 높은 전기료, 태양광 보급률

2. Tier 2 시장 (중기 진출)
   - 일본: 재해 대비, 안정성 중시
   - 인도: 전력 부족, 급성장 시장
   - 중동: 신재생에너지 전환

II. 고객 세그먼트별 전략

1. Utility (전력회사)
   - 타겟: 100MWh 이상 대형 프로젝트
   - 강점: 대용량, 고신뢰성
   - 영업 전략: 장기 계약, 성능 보증

2. C&I (상업/산업)
   - 타겟: 1-10MWh 중형 프로젝트
   - 강점: 맞춤형 솔루션
   - 영업 전략: TCO 절감, ESG 가치

3. Residential (가정용)
   - 타겟: 10-30kWh 소형 시스템
   - 강점: 안전성, 디자인
   - 영업 전략: 설치업체 파트너십

III. 영업 채널 구축

1. 직접 영업
   - 글로벌 영업 조직 구축
   - 현지 법인 설립 (미국, 유럽)
   - 기술 영업 인력 200명 확보

2. 파트너 채널
   - EPC 업체 제휴
   - 인버터 업체 협력
   - 신재생에너지 개발사 파트너십

3. 디지털 마케팅
   - 온라인 리드 생성
   - 웨비나 및 가상 전시회
   - SNS 브랜드 마케팅

IV. 가격 전략
- 프리미엄 포지셔닝 (안전성 강조)
- 초기: 경쟁사 대비 5-10% 프리미엄
- 중기: 규모의 경제로 가격 경쟁력 확보
- TCO 기준 경쟁우위 강조

V. 2025년 영업 목표
- 수주: 5GWh
- 매출: 3조원
- 신규 고객: 50개사
- 시장점유율: 5%
        """,
        "created_at": datetime.now() - timedelta(days=5)
    },
    {
        "type": "internal",
        "title": "ESS 생산 계획 및 공급망 전략",
        "organization": "생산",
        "category": "생산",
        "content": """
ESS 생산 체계 구축 계획

I. 생산 능력 확대 계획

1. 신규 공장 건설
   - 충주 ESS 전용공장: 20GWh (2026년 가동)
   - 헝가리 공장 ESS 라인: 10GWh (2027년)
   - 총 생산능력: 30GWh (2027년)

2. 기존 설비 전환
   - EV용 라인 일부 ESS 전환
   - 유연 생산 체계 구축
   - 공용 모듈 설계 적용

II. 공급망 관리

1. 핵심 원자재 확보
   - 리튬: 장기계약 50%, 스팟 50%
   - LFP 전구체: 중국 외 공급선 다변화
   - 전해질: 국내 생산 비중 확대

2. 부품 공급망
   - BMS: 자체 개발 + 외주 생산
   - 인버터: 글로벌 파트너십
   - 함체: 현지 생산 체계

3. 공급망 리스크 관리
   - 이중화 소싱 전략
   - 재고 버퍼 관리
   - 대체재 개발

III. 생산 효율화

1. 자동화 수준
   - 셀 생산: 95% 자동화
   - 모듈 조립: 80% 자동화
   - 시스템 통합: 60% 자동화

2. 품질 관리
   - 실시간 품질 모니터링
   - AI 기반 불량 예측
   - 전수 성능 테스트

3. 원가 절감
   - 목표: 연 10% 원가 절감
   - 주요 수단: 규모의 경제, 수율 개선
   - 2027년 $100/kWh 달성

IV. ESS 특화 요구사항

1. 대용량 모듈
   - 표준 모듈: 100kWh
   - 컨테이너: 3MWh
   - 운송 최적화 설계

2. 장수명 요구사항
   - 20년 수명 보증
   - 열화율 연 2% 이하
   - 원격 업그레이드 지원

V. 투자 및 인력 계획
- 설비 투자: 2조원 (2025-2027)
- 신규 채용: 1,000명
- 교육 훈련: 100억원/년
        """,
        "created_at": datetime.now() - timedelta(days=3)
    }
]

# ESS 관련 외부 뉴스들
external_ess_documents = [
    {
        "type": "external",
        "title": "[속보] 미국, ESS 투자세액공제 40%로 확대",
        "organization": "에너지경제",
        "category": "정책",
        "content": """
미국 정부가 에너지저장장치(ESS) 투자에 대한 세액공제를 기존 30%에서 40%로 확대한다고 발표했다.

이번 조치는 IRA(인플레이션감축법) 세부 규정 개정을 통해 이뤄졌으며, 미국 내 생산 ESS에 대해서는 추가 10% 인센티브를 제공한다.

주요 내용:
- 기본 세액공제: 30% → 40%
- 미국산 배터리 사용 시: +10%
- 저소득 지역 설치: +10%
- 최대 세액공제: 60%

업계 반응:
"한국 배터리 기업들에게 호재" - 업계 관계자
"테슬라, LG에너지솔루션 등 수혜 예상"
"2025년부터 본격적인 투자 붐 예상"

정책 배경:
- 재생에너지 간헐성 해결
- 전력망 안정성 강화
- 중국 의존도 감소
- 일자리 창출

시장 전망:
- 미국 ESS 시장 연 50% 성장 예상
- 2030년까지 200GWh 수요
- 투자 규모 500억 달러
        """,
        "created_at": datetime.now() - timedelta(days=2)
    },
    {
        "type": "external",
        "title": "[글로벌] CATL, 세계 최대 ESS 프로젝트 수주",
        "organization": "배터리인사이트",
        "category": "경쟁사",
        "content": """
중국 CATL이 사우디아라비아에서 10GWh 규모의 세계 최대 ESS 프로젝트를 수주했다.

프로젝트 개요:
- 규모: 10GWh (10,000MWh)
- 금액: 50억 달러
- 기간: 2025-2027년
- 위치: 사우디 네옴시티

CATL 전략:
- LFP 배터리 100% 적용
- 액냉 시스템 탑재
- 20년 성능 보증
- 현지 생산 검토

경쟁사 동향:
- BYD: 중동 5GWh 프로젝트 추진
- 테슬라: Megapack 공급 확대
- LG엔솔: 미국 시장 집중
- 삼성SDI: 유럽 시장 공략

시장 영향:
"중국 기업의 가격 경쟁력 입증"
"한국 기업들 차별화 전략 필요"
"안전성과 품질로 승부해야"

중동 ESS 시장:
- 2030년까지 50GWh 수요
- 태양광 연계 프로젝트 급증
- 석유 의존도 감소 정책
        """,
        "created_at": datetime.now() - timedelta(days=4)
    },
    {
        "type": "external",
        "title": "[리포트] 글로벌 ESS 시장, 2030년 1,200억 달러 전망",
        "organization": "마켓리서치",
        "category": "산업",
        "content": """
글로벌 ESS 시장이 2030년까지 1,200억 달러 규모로 성장할 것이라는 전망이 나왔다.

시장 전망:
- 2024년: 200억 달러
- 2027년: 600억 달러  
- 2030년: 1,200억 달러
- CAGR: 35%

성장 동력:
1. 재생에너지 확산
   - 태양광, 풍력 간헐성 해결
   - Grid Parity 달성
   - 정부 의무화 정책

2. 전력망 현대화
   - 노후 인프라 교체
   - 스마트그리드 구축
   - 분산전원 확대

3. 전기요금 상승
   - Peak Shaving 수요
   - Time-of-Use 요금제
   - 수요반응(DR) 시장

지역별 전망:
- 북미: 35% (IRA 효과)
- 유럽: 25% (RE100)
- 중국: 20% (내수 시장)
- 기타: 20%

기술 트렌드:
- LFP 비중 확대 (70%)
- 나트륨이온 상용화
- AI 기반 운영 최적화
- 모듈러 설계 확산
        """,
        "created_at": datetime.now() - timedelta(days=6)
    },
    {
        "type": "external",
        "title": "[동향] 테슬라 Megapack, 분기 매출 10억 달러 돌파",
        "organization": "테크뉴스",
        "category": "경쟁사",
        "content": """
테슬라의 ESS 사업부문이 분기 매출 10억 달러를 돌파하며 급성장세를 보이고 있다.

실적 하이라이트:
- Q3 매출: 10억 달러
- 전년 대비: +200%
- 영업이익률: 25%
- 백로그: 50GWh

Megapack 경쟁력:
- 용량: 3.9MWh/유닛
- 가격: $300/kWh
- 설치 시간: 3개월
- 수명: 20년

주요 프로젝트:
- 캘리포니아: 1GWh
- 텍사스: 500MWh
- 호주: 300MWh
- 영국: 200MWh

시장 전략:
"자체 배터리 생산으로 원가 경쟁력"
"소프트웨어 차별화"
"Autobidder로 수익 극대화"
"가상발전소(VPP) 사업 확대"

경쟁 구도:
- CATL: 가격 경쟁
- Fluence: 시스템 통합
- Wartsila: 유럽 강세
- 한국 3사: 품질 승부
        """,
        "created_at": datetime.now() - timedelta(days=8)
    },
    {
        "type": "external",
        "title": "[정책] EU, 2030년까지 ESS 200GWh 구축 목표",
        "organization": "EU에너지",
        "category": "정책",
        "content": """
유럽연합(EU)이 2030년까지 200GWh 규모의 ESS를 구축하겠다는 야심찬 목표를 발표했다.

정책 목표:
- 2025년: 30GWh
- 2027년: 100GWh
- 2030년: 200GWh
- 투자액: 2,000억 유로

핵심 정책:
1. ESS 의무화
   - 신재생 발전소 ESS 연계 의무
   - 용량의 20% ESS 설치
   - 그리드 안정성 기여

2. 보조금 지원
   - 설치비 30% 지원
   - 저리 대출 제공
   - 세제 혜택

3. 규제 완화
   - 인허가 간소화
   - Grid Code 개정
   - 안전 기준 통일

시장 기회:
"한국 기업들에게 큰 기회"
"유럽산 요구사항 대응 필요"
"Battery Passport 준비 필수"

주요 프로젝트:
- 독일: 50GWh
- 프랑스: 30GWh
- 스페인: 25GWh
- 이탈리아: 20GWh
        """,
        "created_at": datetime.now() - timedelta(days=11)
    },
    {
        "type": "external",
        "title": "[분석] LFP vs NCM, ESS 시장 주도권 경쟁",
        "organization": "배터리저널",
        "category": "기술",
        "content": """
ESS 시장에서 LFP(리튬인산철) 배터리가 NCM 대비 우위를 점하고 있다.

시장 점유율:
- LFP: 70% (2024년)
- NCM: 25%
- 기타: 5%

LFP 장점:
1. 안전성
   - 열폭주 온도 270°C (NCM 210°C)
   - 화재 위험 현저히 낮음
   - 산소 발생 없음

2. 수명
   - 8,000 사이클 (NCM 3,000)
   - 20년 사용 가능
   - 열화율 낮음

3. 가격
   - $80/kWh (NCM $120/kWh)
   - 원자재 가격 안정
   - 코발트 미사용

NCM 대응:
- 고에너지밀도 강조
- 고출력 응용 타겟
- 하이브리드 ESS

기업별 전략:
- CATL/BYD: LFP 올인
- LG엔솔: NCM+LFP 투트랙
- 삼성SDI: NCM 고도화
- SK온: LFP 라이선스

전망:
"2027년 LFP 80% 전망"
"안전 규제 강화가 변수"
"나트륨이온 대체 가능성"
        """,
        "created_at": datetime.now() - timedelta(days=13)
    },
    {
        "type": "external",
        "title": "[뉴스] 호주 대정전 사태, ESS가 막았다",
        "organization": "글로벌에너지",
        "category": "산업",
        "content": """
호주에서 대규모 정전 위기를 ESS가 성공적으로 방어해 주목받고 있다.

사건 개요:
- 일시: 2024년 11월 15일
- 원인: 송전선로 고장
- 규모: 2GW 전력 손실
- ESS 대응: 500MW 즉시 공급

ESS 역할:
1. 즉각 대응
   - 0.2초 내 전력 공급
   - 주파수 안정화
   - 계통 붕괴 방지

2. 경제적 효과
   - 정전 피해 10조원 방지
   - ESS 투자비 1조원
   - ROI 1,000%

호주 ESS 현황:
- 설치 용량: 5GWh
- 2025년 목표: 15GWh
- 주요 사업자: Tesla, Neoen

글로벌 시사점:
"ESS 필요성 입증"
"전력망 회복력 핵심"
"각국 투자 가속화 예상"

한국 시장:
- 현재: 10GWh
- 목표: 30GWh (2030)
- 과제: 안전 규제 완화
        """,
        "created_at": datetime.now() - timedelta(days=14)
    }
]

async def upload_ess_documents():
    """ESS 관련 문서를 Supabase에 업로드"""
    client = SupabaseClient()
    ingester = DocumentIngester()
    
    all_documents = internal_ess_documents + external_ess_documents
    
    print(f"총 {len(all_documents)}개 ESS 관련 문서 업로드 시작...")
    
    for i, doc in enumerate(all_documents, 1):
        try:
            # 문서 정보 생성
            doc_data = {
                "type": doc["type"],
                "title": doc["title"],
                "organization": doc["organization"],
                "category": doc["category"],
                "created_at": doc["created_at"].isoformat(),
                "source": "ESS Demo Data",
                "file_path": f"demo/ess_{i}.txt",
                "metadata": {
                    "demo": True,
                    "topic": "ESS",
                    "importance": "high"
                }
            }
            
            # 문서 삽입
            doc_response = client.client.table("documents").insert(doc_data).execute()
            
            if doc_response.data:
                document_id = doc_response.data[0]["id"]
                print(f"[{i}/{len(all_documents)}] 문서 업로드 성공: {doc['title']}")
                
                # 청크 생성 및 임베딩
                chunks = await ingester.create_chunks(doc["content"], document_id)
                print(f"  - {len(chunks)}개 청크 생성 완료")
                
                # 임베딩 생성
                for chunk in chunks:
                    await ingester.create_embedding(chunk["id"], chunk["content"])
                print(f"  - 임베딩 생성 완료")
            
        except Exception as e:
            print(f"오류 발생 ({doc['title']}): {str(e)}")
            continue
    
    print("\n✅ ESS 시연 데이터 업로드 완료!")
    print("경영진 시연 시나리오:")
    print("1. ESS 시장 진출 전략 검색 (내부 문서 중심)")
    print("2. 사외 정보 가중치 증가 (경쟁사 및 시장 동향)")
    print("3. Issue Timeline에서 조직별 대응 확인")
    print("4. Smart Alerts에서 실시간 모니터링")

if __name__ == "__main__":
    asyncio.run(upload_ess_documents())