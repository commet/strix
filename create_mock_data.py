"""
Create Mock Data for STRIX Testing
"""
import os
import sys
from datetime import datetime, timedelta
import random

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

def create_mock_files():
    """Create mock documents for testing"""
    
    # Create directories
    base_path = "./mock_data"
    internal_path = os.path.join(base_path, "internal")
    external_path = os.path.join(base_path, "external")
    
    for path in [base_path, internal_path, external_path]:
        os.makedirs(path, exist_ok=True)
    
    # Create subdirectories for internal documents
    orgs = ["전략기획", "R&D", "경영지원", "생산", "영업마케팅"]
    for org in orgs:
        org_path = os.path.join(internal_path, org)
        os.makedirs(org_path, exist_ok=True)
    
    # Create subdirectories for external news
    sources = ["PR팀_AM", "PR팀_PM", "Google_Alert", "Naver_News"]
    for source in sources:
        source_path = os.path.join(external_path, source)
        os.makedirs(source_path, exist_ok=True)
    
    print(f"[OK] Created directory structure at {base_path}")
    
    # Sample internal documents
    internal_docs = [
        {
            "org": "전략기획",
            "filename": "2024_Q1_배터리사업_중장기전략.txt",
            "content": """2024년 1분기 배터리 사업 중장기 전략 보고서

1. 개요
- 글로벌 전기차 시장 성장에 따른 배터리 수요 급증
- 전고체 배터리 기술 개발을 통한 차세대 시장 선점 필요
- 경영진 특별 관심 사항: ESG 경영 강화 및 탄소중립 달성

2. 주요 전략
- NCM811 고니켈 배터리 양산 체제 구축
- 전고체 배터리 파일럿 생산 라인 구축 (2025년 목표)
- 북미 현지 생산 공장 설립 추진 (IRA 대응)

3. 투자 계획
- 총 5조원 규모 투자 (2024-2027)
- R&D 투자 1.5조원 (전고체 배터리 중심)
- 생산 설비 투자 3.5조원

4. 리스크 관리
- 원자재 가격 변동 리스크 헤징
- 지정학적 리스크 대응 (공급망 다변화)
- 기술 유출 방지 체계 강화"""
        },
        {
            "org": "R&D",
            "filename": "2024_01_전고체배터리_개발현황.txt",
            "content": """전고체 배터리 개발 현황 보고

1. 기술 개발 진척도
- 황화물계 고체전해질 개발 완료 (이온전도도 10mS/cm 달성)
- 실험실 규모 셀 제작 성공 (에너지밀도 400Wh/kg)
- 안전성 테스트 통과 (열폭주 없음)

2. 주요 성과
- 특허 출원 15건 (핵심 특허 3건 포함)
- 국제 학회 논문 발표 5건
- 정부 과제 선정 (100억원 지원)

3. 향후 계획
- 파일럿 생산 라인 구축 (2024년 하반기)
- 자동차 업체와 공동 개발 추진
- 양산 기술 개발 (2025년 목표)

4. 필요 지원 사항
- 추가 연구 인력 20명 충원
- 분석 장비 도입 (50억원)
- 해외 전문가 영입"""
        },
        {
            "org": "경영지원",
            "filename": "2024_Q1_리스크관리_현황.txt",
            "content": """2024년 1분기 리스크 관리 현황

1. 주요 리스크 식별
- 리튬 가격 변동성 증가 (전년 대비 30% 상승)
- EU 배터리 규제 강화 (탄소발자국 공시 의무화)
- 중국 경쟁사 공격적 가격 정책

2. 대응 방안
- 장기 공급 계약 체결 (리튬 70%, 니켈 60%)
- LCA(Life Cycle Assessment) 시스템 구축
- 원가 절감 TF 운영 (목표: 10% 절감)

3. ESG 리스크 관리
- RE100 가입 및 재생에너지 전환 계획
- 공급망 실사 강화 (Tier 2까지 확대)
- 안전 관리 체계 고도화

4. 컴플라이언스
- 반부패 교육 실시 (전직원 대상)
- 내부 감사 강화
- 윤리 경영 체계 재정비"""
        }
    ]
    
    # Sample external news
    external_news = [
        {
            "source": "PR팀_AM",
            "filename": f"{datetime.now().strftime('%Y-%m-%d')}_AM_배터리산업동향.txt",
            "content": """[산업] 글로벌 배터리 시장 동향 브리핑

□ 주요 뉴스
1. CATL, 신규 배터리 공장 착공 발표
   - 헝가리 데브레첸 100GWh 규모
   - 2025년 양산 시작 예정
   - 유럽 완성차 업체 공급 목적

2. Tesla, 4680 배터리 생산 확대
   - 텍사스 공장 생산량 2배 증설
   - 자체 배터리 비중 50% 목표
   - 원가 30% 절감 달성

3. 미국 IRA 세부 규정 발표
   - 배터리 핵심 광물 요건 강화
   - 중국산 부품 사용 제한 확대
   - 2024년부터 단계적 적용

□ 시장 지표
- 리튬 가격: $70,000/톤 (전주 대비 +5%)
- 니켈 가격: $20,000/톤 (전주 대비 +3%)
- 코발트 가격: $35,000/톤 (전주 대비 -2%)"""
        },
        {
            "source": "PR팀_PM",
            "filename": f"{datetime.now().strftime('%Y-%m-%d')}_PM_정책규제동향.txt",
            "content": """[정책/규제] 배터리 관련 정책 동향

□ EU 배터리 규제 업데이트
1. 배터리 여권(Battery Passport) 시행 세칙
   - 2026년부터 의무화
   - QR 코드 통한 정보 공개
   - 탄소발자국, 재활용 정보 포함

2. 탄소국경조정메커니즘(CBAM) 확대
   - 배터리 제품 포함 검토
   - 2025년 시범 적용 예정
   - 탄소 배출량 계산 방법론 확정

□ 한국 정부 지원 정책
1. K-배터리 발전 전략 발표
   - 2030년까지 40조원 투자
   - 전고체 배터리 개발 집중 지원
   - 인력 양성 프로그램 확대

2. 배터리 특화단지 조성
   - 새만금, 울산 등 3개 지역
   - 규제 샌드박스 적용
   - 세제 혜택 제공"""
        },
        {
            "source": "Google_Alert",
            "filename": f"{(datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')}_전고체배터리_기술동향.txt",
            "content": """[기술] Solid-State Battery Technology Updates

Recent developments in solid-state battery technology:

1. Toyota announces breakthrough in solid-state battery
   - Energy density: 500 Wh/kg achieved
   - Charging time: 10 minutes to 80%
   - Commercial production targeted for 2027

2. QuantumScape reports Q4 results
   - 24-layer cells testing successful
   - Retained 80% capacity after 800 cycles
   - Seeking automotive partners for validation

3. Samsung SDI solid-state progress
   - Pilot line operational in Suwon
   - Silver-carbon composite anode developed
   - Targeting premium EV market initially

Technical challenges remaining:
- Lithium dendrite formation
- Interface resistance
- Manufacturing scalability
- Cost reduction needs"""
        }
    ]
    
    # Write internal documents
    for doc in internal_docs:
        file_path = os.path.join(internal_path, doc["org"], doc["filename"])
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(doc["content"])
        print(f"[OK] Created: {file_path}")
    
    # Write external news
    for news in external_news:
        file_path = os.path.join(external_path, news["source"], news["filename"])
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(news["content"])
        print(f"[OK] Created: {file_path}")
    
    print(f"\n[OK] Mock data creation complete!")
    print(f"  - Internal documents: {len(internal_docs)}")
    print(f"  - External news: {len(external_news)}")

if __name__ == "__main__":
    create_mock_files()