"""
Create strategic documents for executive demonstration
전사적 관점의 경영전략 문서 생성
"""
import sys
import os
from datetime import datetime, timedelta
import random
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from database.supabase_client import SupabaseClient
from langchain_openai import OpenAIEmbeddings
from langchain.text_splitter import RecursiveCharacterTextSplitter

# 내부 전략 문서 템플릿
internal_documents = [
    {
        "title": "2025년 전사 경영전략 방향",
        "organization": "전략기획",
        "category": "전략기획",
        "content": """
2025년 전사 경영전략 방향

I. 핵심 전략 방향
1. 수익성 중심 경영 전환
   - 선택과 집중을 통한 포트폴리오 최적화
   - 저수익 사업 구조조정 및 고부가가치 제품 확대
   - 목표: 2025년 영업이익률 8% → 2027년 12% 달성

2. 기술 리더십 확보
   - 차세대 배터리 기술 개발 가속화
   - 전고체 배터리: 2027년 양산, 2030년 본격 상용화
   - AI/디지털 기술 활용한 생산성 30% 향상

3. 글로벌 시장 확대
   - 미국/유럽 생산기지 확충 (IRA/CRMA 대응)
   - 아시아 시장 점유율 방어 및 확대
   - 신흥시장(인도, 동남아) 진출 본격화

II. 재무 목표
- 매출: 2024년 15조원 → 2025년 20조원 → 2027년 35조원
- 영업이익: 2024년 0.8조원 → 2025년 1.6조원 → 2027년 4.2조원
- CAPEX: 2025-2027년 총 15조원 투자
- ROE: 2025년 10% → 2027년 15% 목표

III. 핵심 실행 과제
1. 원가 경쟁력 확보
   - 원자재 직접 소싱 비중 확대 (30% → 60%)
   - 생산 자동화율 향상 (70% → 90%)
   - 수율 개선 (92% → 95%)

2. 고객 다변화
   - Top 3 고객 의존도 감소 (70% → 50%)
   - 신규 OEM 10개사 확보
   - ESS/전력망 시장 진출

3. 인재 확보 및 육성
   - 핵심 인재 Retention 프로그램
   - 글로벌 R&D 인력 30% 증원
   - 디지털 전환 전문가 영입
        """
    },
    {
        "title": "경쟁사 벤치마킹 분석 보고서",
        "organization": "전략기획",
        "category": "전략기획",
        "content": """
글로벌 경쟁사 벤치마킹 분석 (2025.1)

I. 경쟁사 포지셔닝
1. CATL (중국)
   - 시장점유율: 37% (1위)
   - 강점: 원가경쟁력, 중국내수시장, LFP 기술
   - 약점: 지정학적 리스크, 품질 이슈
   - 전략: 글로벌 생산기지 확대, CTP/CTC 기술 선도

2. BYD (중국)
   - 시장점유율: 16% (2위)
   - 강점: 수직계열화, Blade Battery
   - 약점: 해외시장 경험 부족
   - 전략: 완성차 사업 시너지, 가격 경쟁력

3. LG에너지솔루션 (한국)
   - 시장점유율: 14% (3위)
   - 강점: 품질, 글로벌 고객사, NCM 기술
   - 약점: 원가경쟁력, 중국시장 부재
   - 전략: 미국/유럽 현지화, 차세대 배터리

4. 당사 포지션
   - 시장점유율: 5% (6위)
   - 강점: 기술력, 품질, 주요 OEM 관계
   - 약점: 규모의 경제, 수익성
   - 기회: IRA 수혜, 전고체 선도 가능성

II. 벤치마킹 시사점
1. 원가 경쟁력
   - CATL 대비 20% 높은 제조원가
   - 개선방안: 소재 내재화, 자동화, 수율 개선

2. 기술 차별화
   - 전고체/리튬메탈 집중 투자
   - 안전성/수명 차별화
   - 충전속도 개선 (10분 80%)

3. 사업 모델
   - B2B에서 B2B2C로 확대
   - 배터리 구독/리스 모델
   - 에너지 솔루션 사업 진출

III. 대응 전략
1. 단기 (2025년)
   - 수익성 개선 집중
   - 핵심 고객사 Lock-in
   - 품질 차별화 강화

2. 중기 (2025-2027년)
   - 차세대 기술 상용화
   - 미국/유럽 생산 확대
   - 신사업 모델 구축

3. 장기 (2028년~)
   - 기술 리더십 확보
   - 글로벌 Top 3 진입
   - 토탈 에너지 기업 전환
        """
    },
    {
        "title": "리스크 관리 현황 및 대응 방안",
        "organization": "경영지원",
        "category": "경영지원",
        "content": """
2025년 전사 리스크 관리 현황

I. 주요 리스크 식별
1. 시장 리스크 (확률: 높음, 영향: 높음)
   - 전기차 수요 둔화 지속
   - 중국 업체 가격 공세 심화
   - 대응: 고객/제품 포트폴리오 다변화

2. 원자재 리스크 (확률: 중간, 영향: 높음)
   - 리튬 가격 변동성 확대
   - 희소금속 공급 불안정
   - 대응: 장기계약 확대, 재활용 사업 강화

3. 기술 리스크 (확률: 중간, 영향: 중간)
   - 차세대 배터리 개발 지연
   - 특허 분쟁 가능성
   - 대응: R&D 투자 확대, IP 포트폴리오 강화

4. 규제 리스크 (확률: 높음, 영향: 중간)
   - 환경 규제 강화
   - 보조금 정책 변화
   - 대응: 선제적 대응, 정책 모니터링

5. 운영 리스크 (확률: 낮음, 영향: 높음)
   - 생산시설 사고
   - 품질 문제 발생
   - 대응: 안전 시스템 강화, 품질 관리 고도화

II. 리스크별 대응 전략
1. 헤징 전략
   - 원자재: 선물/옵션 활용 (30% 헤징)
   - 환율: 자연헤지 + 파생상품 (50% 헤징)
   - 고객: 장기계약 + Take-or-Pay

2. 다변화 전략
   - 지역: 미국 40%, 유럽 30%, 아시아 30%
   - 제품: EV 70%, ESS 20%, 기타 10%
   - 고객: Top 10 고객 80% → 60% 축소

3. 보험/준비금
   - 제조물책임보험: 1조원 가입
   - 리스크 준비금: 매출의 2% 적립
   - 긴급 대응 TF 상시 운영

III. 모니터링 체계
- 월간 리스크 대시보드 운영
- 분기별 리스크 위원회 개최
- 연간 리스크 평가 및 전략 수정
        """
    },
    {
        "title": "2025년 투자 계획 및 우선순위",
        "organization": "재무",
        "category": "경영지원",
        "content": """
2025년 투자 계획 및 우선순위

I. 총 투자 규모: 5조원
1. 생산능력 확대: 2.5조원 (50%)
   - 미국 2공장 건설: 1.2조원
   - 유럽 JV 투자: 0.8조원
   - 국내 라인 증설: 0.5조원

2. R&D 투자: 1.5조원 (30%)
   - 전고체 배터리: 0.7조원
   - 차세대 소재: 0.4조원
   - AI/디지털화: 0.4조원

3. 인프라/운영: 0.7조원 (14%)
   - IT 시스템 업그레이드: 0.3조원
   - 물류/SCM 최적화: 0.2조원
   - 안전/환경 설비: 0.2조원

4. M&A/지분투자: 0.3조원 (6%)
   - 소재 기업 인수: 0.2조원
   - 스타트업 투자: 0.1조원

II. 투자 우선순위 평가
우선순위 1: 미국 생산시설 (ROI 25%)
- IRA 혜택 최대화
- 주요 고객사 인접 생산
- 2026년 하반기 양산 목표

우선순위 2: 전고체 배터리 개발 (ROI 40%)
- 게임체인저 기술
- 2027년 파일럿 생산
- 2030년 대량 양산

우선순위 3: 유럽 현지화 (ROI 20%)
- CRMA 대응
- 2026년 생산 시작
- 현지 파트너십 활용

III. 투자 효과 분석
1. 생산능력
   - 2024년: 50GWh
   - 2025년: 80GWh
   - 2027년: 200GWh

2. 매출 기여
   - 신규 투자 매출: 2026년 3조원
   - 2027년: 8조원
   - 2030년: 20조원

3. 수익성 개선
   - 원가 절감: 15%
   - 생산성 향상: 30%
   - 영업이익률: 8% → 12%

IV. 자금 조달 계획
- 영업현금흐름: 2조원
- 차입금: 2조원 (부채비율 60% 유지)
- 유상증자: 1조원 (2025년 3분기)
        """
    },
    {
        "title": "ESG 경영 전략 및 실행 계획",
        "organization": "ESG추진",
        "category": "경영지원",
        "content": """
2025년 ESG 경영 전략

I. ESG 비전 및 목표
비전: "지속가능한 에너지 미래를 선도하는 기업"

2030 목표:
- Environment: 탄소중립 생산 (Scope 1,2)
- Social: 중대재해 Zero, 다양성 30%
- Governance: 이사회 독립성 70%, ESG 연계 보상 50%

II. 환경(E) 전략
1. 탄소 감축
   - 2025년: 2019년 대비 30% 감축
   - RE100: 2025년 60%, 2030년 100%
   - 공급망: Scope 3 관리 체계 구축

2. 순환경제
   - 배터리 재활용률: 95% 달성
   - 원자재 회수율: 리튬 90%, 코발트 95%
   - 폐배터리 수거 체계 구축

3. 친환경 제품
   - LCA 기반 제품 설계
   - 탄소발자국 50% 감축
   - 친환경 인증 100% 취득

III. 사회(S) 전략
1. 안전 최우선
   - 중대재해 Zero 달성
   - 안전투자 2,000억원
   - 안전문화 지수 90점

2. 인재 경영
   - 여성 임원 비율 25%
   - 글로벌 인재 40%
   - 직원 만족도 85점

3. 공급망 관리
   - 책임있는 소싱 100%
   - 공급업체 ESG 평가
   - 동반성장 프로그램

IV. 지배구조(G) 전략
1. 이사회 다양성
   - 독립이사 70%
   - 여성이사 30%
   - ESG 위원회 신설

2. 투명경영
   - 통합보고서 발간
   - TCFD 공시
   - 이해관계자 소통 강화

3. 윤리경영
   - 부패방지 시스템
   - 내부신고 활성화
   - 컴플라이언스 강화

V. 실행 계획
2025년 1분기:
- ESG 위원회 설립
- 탄소중립 로드맵 수립
- 공급망 실사 시작

2025년 2분기:
- RE100 가입
- ESG 평가 등급 A 획득
- 안전 시스템 고도화

2025년 하반기:
- 순환경제 사업 본격화
- ESG 연계 KPI 도입
- 통합보고서 발간
        """
    }
]

# 외부 시장/산업 문서 템플릿
external_documents = [
    {
        "title": "[시장분석] 2025 글로벌 배터리 시장 전망",
        "organization": "마켓리서치",
        "category": "산업",
        "content": """
[시장분석] 2025 글로벌 배터리 시장 전망

□ 시장 규모 및 성장률
ㅇ 2025년 글로벌 배터리 시장: 1,200GWh (YoY +25%)
ㅇ 2030년 전망: 3,500GWh (CAGR 24%)
ㅇ 지역별: 중국 45%, 유럽 25%, 북미 20%, 기타 10%

□ 주요 트렌드
1. 수요 측면
   - EV 침투율 가속화: 2025년 신차의 25%
   - ESS 시장 급성장: YoY +45%
   - 전동 모빌리티 확산

2. 공급 측면
   - 생산능력 과잉 우려 (2025년 1,500GWh)
   - 가격 경쟁 심화 ($100/kWh → $80/kWh)
   - 기술 차별화 가속

3. 기술 트렌드
   - LFP 점유율 확대 (40% → 55%)
   - 실리콘 음극 상용화
   - 전고체 배터리 pilot 생산

□ 경쟁 구도
ㅇ 중국 업체 지배력 강화 (전체 65%)
ㅇ 한국 업체 프리미엄 시장 집중
ㅇ 일본 업체 전고체 선도
ㅇ 구미 업체 로컬 생산 확대

□ 리스크 요인
- 원자재 가격 변동성
- 지정학적 갈등 심화
- 환경 규제 강화
- 기술 표준 경쟁

□ 투자 시사점
- 차별화된 기술력 필수
- 현지화 생산 불가피
- 수직계열화 경쟁력
- ESG 대응 필수
        """
    },
    {
        "title": "[정책동향] 미국 IRA 2025년 개정안 영향",
        "organization": "정책연구원",
        "category": "정책",
        "content": """
[정책동향] 미국 IRA 2025년 개정안 영향 분석

□ 주요 개정 내용
ㅇ 배터리 부품 요건: 60% → 70% (2025.7.1)
ㅇ 핵심광물 요건: 50% → 60% (2025.7.1)
ㅇ 중국산 배제: 2025년부터 전면 시행
ㅇ 세액공제 상한: $7,500 유지

□ 한국 기업 영향
1. 긍정적 영향
   - 한국산 프리미엄 상승
   - 미국 투자 가속화
   - 장기 공급계약 증가

2. 부정적 영향
   - 원가 부담 증가 (10-15%)
   - 공급망 재편 비용
   - 중국 소재 의존도 이슈

□ 대응 전략
ㅇ 미국 현지 생산 확대 필수
ㅇ 우호국 소재 공급망 구축
ㅇ JV/파트너십 활용
ㅇ 정부 지원 최대 활용

□ 중장기 전망
- 2025-2027: 전환기 진통
- 2028-2030: 안정화 및 성장
- 북미 시장 재편 가속화
- 한-미 배터리 동맹 강화

□ 정책 제언
- 정부 차원 협상 지속
- 보조금 지원 확대
- 인력 양성 프로그램
- R&D 세제 혜택
        """
    },
    {
        "title": "[산업분석] 중국 배터리 업체 글로벌 전략",
        "organization": "산업연구소",
        "category": "경쟁사",
        "content": """
[산업분석] 중국 배터리 업체 글로벌 확장 전략

□ CATL 전략
ㅇ 현지화: 독일, 헝가리 공장 가동
ㅇ 기술 라이선싱: Ford, Tesla 협력
ㅇ 가격 공세: $60/kWh 목표
ㅇ M&A: 상류 자원 기업 인수

□ BYD 전략
ㅇ 완성차 시너지: 수직 통합
ㅇ Blade Battery: 안전성 차별화
ㅇ 신흥시장: 동남아, 중남미 진출
ㅇ 가격 경쟁력: 최저가 전략

□ 2군 업체 동향
ㅇ CALB: 항공/선박용 특화
ㅇ Gotion: VW 전략적 파트너
ㅇ EVE: 원통형 배터리 집중
ㅇ SVOLT: 무코발트 배터리

□ 경쟁 전략
1. 가격: 공격적 저가 정책
2. 기술: Fast-follower + 원가 혁신
3. 시장: 중국 내수 + 신흥시장
4. 공급망: 자원 확보 전쟁

□ 대응 방안
- 기술 차별화 강화
- 프리미엄 시장 집중
- 서구 시장 Lock-in
- 차세대 기술 선점
        """
    },
    {
        "title": "[기술리포트] 전고체 배터리 상용화 전망",
        "organization": "기술분석센터",
        "category": "기술",
        "content": """
[기술리포트] 전고체 배터리 2025-2030 상용화 로드맵

□ 기술 현황
ㅇ 일본: 도요타/파나소닉 2027년 양산 목표
ㅇ 한국: 삼성SDI/LG에너지 2027년 시험생산
ㅇ 중국: CATL/BYD 2030년 목표
ㅇ 미국: QuantumScape 2025년 샘플 공급

□ 핵심 과제
1. 기술적 과제
   - 고체전해질 이온전도도
   - 계면 저항 문제
   - 덴드라이트 억제
   - 양산 공정 개발

2. 경제적 과제
   - 높은 제조 비용 (3-5배)
   - 수율 문제 (목표 90%)
   - 설비 투자 규모

□ 상용화 시나리오
ㅇ 2025-2026: 프리미엄 EV 소량 적용
ㅇ 2027-2028: 양산 시작, 고급차 확대
ㅇ 2029-2030: 대중화 시작
ㅇ 2030+: 시장 주류화

□ 시장 영향
- 에너지밀도 2배 향상 (800Wh/L)
- 충전시간 1/3 단축 (10분)
- 화재 위험 제거
- 수명 2배 연장

□ 투자 기회
- 소재 기업 (고체전해질)
- 장비 기업 (신공정)
- 선도 배터리 기업
- 프리미엄 완성차 기업
        """
    },
    {
        "title": "[원자재분석] 리튬 수급 전망 및 가격 시나리오",
        "organization": "원자재연구소",
        "category": "Macro",
        "content": """
[원자재분석] 2025-2027 리튬 시장 전망

□ 수급 전망
ㅇ 2025년 공급: 1,200천톤 LCE
ㅇ 2025년 수요: 1,100천톤 LCE
ㅇ 수급: 100천톤 과잉 (재고 누적)

□ 가격 시나리오
1. Base Case (확률 60%)
   - 탄산리튬: $15,000/톤
   - 수산화리튬: $18,000/톤
   - 점진적 하락 추세

2. Bull Case (확률 20%)
   - 탄산리튬: $25,000/톤
   - 수요 급증 시나리오
   - 공급 차질 발생

3. Bear Case (확률 20%)
   - 탄산리튬: $10,000/톤
   - 수요 둔화 + 공급 과잉
   - 재고 압박

□ 주요 변수
ㅇ 중국 EV 수요
ㅇ 남미 생산 증설
ㅇ DLE 기술 상용화
ㅇ 재활용 시장 성장

□ 공급망 이슈
- 자원 민족주의 강화
- 환경 규제 강화
- 중국 제련 독점
- 서구 공급망 구축

□ 전략적 시사점
- 장기계약 확대 (70% 목표)
- 상류 투자 참여
- 재활용 사업 강화
- 대체 기술 개발
        """
    },
    {
        "title": "[금융분석] 배터리 섹터 투자 전망",
        "organization": "투자분석팀",
        "category": "Macro",
        "content": """
[금융분석] 2025년 글로벌 배터리 섹터 투자 전망

□ 시장 평가
ㅇ 섹터 시가총액: $500B (2024년 대비 -20%)
ㅇ P/E: 25배 (과거 평균 35배)
ㅇ EV/EBITDA: 12배
ㅇ 투자 심리: Neutral to Positive

□ 투자 포인트
1. 긍정 요인
   - 장기 성장성 견고
   - 밸류에이션 매력
   - 정책 지원 지속
   - 기술 혁신 가속

2. 리스크 요인
   - 단기 수급 불균형
   - 가격 경쟁 심화
   - 지정학적 리스크
   - 기술 전환 리스크

□ 기업별 전망
ㅇ CATL: Outperform (목표가 상향)
ㅇ LG에너지: Buy (턴어라운드 기대)
ㅇ BYD: Hold (밸류에이션 부담)
ㅇ 기타: Selective Buy

□ 투자 전략
- Quality over Quantity
- 기술 리더 선별 투자
- 밸류체인 다각화
- Long-term 관점 유지

□ M&A 전망
- 2025년 예상 규모: $20B
- 수직계열화 M&A 증가
- Cross-border 거래 활발
- 기술 기업 인수 증가
        """
    }
]

def create_and_upload_documents():
    """문서 생성 및 업로드"""
    client = SupabaseClient()
    embeddings = OpenAIEmbeddings()
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=200,
        separators=["\n\n", "\n", ".", "!", "?", ",", " ", ""],
        length_function=len
    )
    
    print("=== Creating Strategic Documents for Executive Demo ===\n")
    
    all_documents = []
    
    # 내부 문서 처리
    for doc_data in internal_documents:
        doc_data['type'] = 'internal'  # English for DB constraint
        doc_data['date'] = datetime.now().strftime("%Y-%m-%d")
        all_documents.append(doc_data)
    
    # 외부 문서 처리
    for doc_data in external_documents:
        doc_data['type'] = 'external'  # English for DB constraint
        doc_data['date'] = datetime.now().strftime("%Y-%m-%d")
        all_documents.append(doc_data)
    
    # 문서 업로드
    for i, doc_data in enumerate(all_documents, 1):
        try:
            print(f"[{i}/{len(all_documents)}] Processing: {doc_data['title']}")
            
            # 문서 삽입
            doc_record = {
                'type': doc_data['type'],
                'source': doc_data.get('organization', ''),
                'title': doc_data['title'],
                'organization': doc_data.get('organization', ''),
                'category': doc_data.get('category', ''),
                'created_at': doc_data['date'],
                'file_path': f"strategic_docs/{doc_data['title']}.txt",
                'metadata': {
                    'purpose': 'executive_demo',
                    'strategic_level': 'high'
                }
            }
            
            doc_result = client.client.table('documents').insert(doc_record).execute()
            doc_id = doc_result.data[0]['id']
            
            # 청크 생성 및 삽입
            chunks = text_splitter.split_text(doc_data['content'])
            
            for j, chunk_text in enumerate(chunks):
                # 청크 삽입
                chunk_record = {
                    'document_id': doc_id,
                    'content': chunk_text,
                    'chunk_index': j,
                    'metadata': {
                        'title': doc_data['title'],
                        'type': doc_data['type']
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
    
    print(f"\n=== Upload Complete ===")
    print(f"Total documents: {len(all_documents)}")
    print(f"Internal documents: {len(internal_documents)}")
    print(f"External documents: {len(external_documents)}")

if __name__ == "__main__":
    create_and_upload_documents()