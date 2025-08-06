"""
Generate Mock Data with Proper Metadata
"""
import os
import json
from datetime import datetime, timedelta
import random

# Base path for mock data
BASE_PATH = "./mock_data/"

# Organizations
ORGANIZATIONS = {
    "internal": {
        "전략기획": "Strategic Planning",
        "R&D": "Research & Development", 
        "경영지원": "Management Support",
        "생산": "Production",
        "영업마케팅": "Sales & Marketing"
    },
    "external": {
        "PR팀_AM": "PR Team Morning",
        "PR팀_PM": "PR Team Afternoon",
        "Google_Alert": "Google Alerts",
        "Naver_News": "Naver News"
    }
}

# Internal documents with metadata
INTERNAL_DOCS = [
    {
        "folder": "전략기획",
        "filename": "2024_Q1_배터리사업_중장기전략.txt",
        "title": "2024년 1분기 배터리 사업 중장기 전략 보고서",
        "date": "2024-01-15",
        "content": """2024년 1분기 배터리 사업 중장기 전략 보고서

1. 시장 환경 분석
- 글로벌 전기차 시장 연평균 25% 성장 전망
- 중국 업체의 가격 경쟁력 위협 증대
- 유럽/미국의 자국 생산 요구 강화

2. 당사 전략 방향
- 고부가가치 제품 중심 포트폴리오 재편
- 차세대 배터리 기술 개발 투자 확대 (전고체 배터리)
- 현지 생산 체제 구축을 통한 시장 대응

3. 투자 계획
- 총 5조원 규모 투자 (2024-2026)
- R&D: 1.5조원 (전고체 배터리 개발)
- 생산 설비: 3.5조원 (미국/유럽 공장)

4. 목표
- 2026년 글로벌 시장 점유율 15% 달성
- 전고체 배터리 2025년 파일럿 생산
- ESG 경영 강화 (탄소중립 달성)"""
    },
    {
        "folder": "R&D",
        "filename": "2024_01_전고체배터리_개발현황.txt",
        "title": "전고체 배터리 개발 현황 보고",
        "date": "2024-01-20",
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
- 양산 기술 개발 (2025년 목표)"""
    },
    {
        "folder": "경영지원",
        "filename": "2024_Q1_리스크관리_보고서.txt",
        "title": "2024년 1분기 리스크 관리 현황",
        "date": "2024-01-25",
        "content": """2024년 1분기 리스크 관리 현황

1. 주요 리스크 식별
- 원자재 가격 변동성 증대 (리튬, 니켈)
- 환율 변동 리스크 (달러 강세)
- 규제 리스크 (EU 배터리 규제)

2. 대응 방안
- 장기 공급 계약 체결 추진
- 환헤지 비율 확대 (70% → 85%)
- 규제 대응 TF 구성 및 운영

3. 모니터링 체계
- 일일 리스크 대시보드 운영
- 월간 리스크 위원회 개최
- 분기별 경영진 보고"""
    }
]

# External news with metadata
EXTERNAL_NEWS = [
    {
        "folder": "Google_Alert",
        "filename": "2024-01-15_CATL_신공장건설.txt",
        "title": "[경쟁사] CATL 유럽 신공장 착공",
        "date": "2024-01-15",
        "category": "경쟁사",
        "content": """[경쟁사 동향] CATL 유럽 신공장 착공

중국 최대 배터리 업체 CATL이 헝가리에 두 번째 유럽 공장을 착공했다고 발표했습니다.

주요 내용:
- 투자 규모: 73억 유로 (약 10조원)
- 생산 규모: 연간 100GWh
- 가동 시기: 2026년 예정
- 고용 인원: 9,000명

시사점:
- 유럽 시장 내 중국 업체 영향력 확대
- 현지 생산을 통한 규제 대응
- 가격 경쟁력 강화 예상"""
    },
    {
        "folder": "PR팀_AM",
        "filename": "2024-01-16_AM_뉴스브리핑.txt",
        "title": "[산업] 글로벌 배터리 수요 급증",
        "date": "2024-01-16",
        "category": "산업",
        "content": """[산업 동향] 글로벌 배터리 수요 급증

시장조사업체 SNE리서치에 따르면 2024년 글로벌 배터리 수요가 전년 대비 40% 증가할 것으로 전망됩니다.

주요 내용:
- 전기차 시장 성장이 주요 동력
- 에너지저장장치(ESS) 수요도 급증
- 공급 부족 우려 지속

당사 대응:
- 생산 능력 확대 필요
- 원자재 확보 전략 재검토
- 고객사 장기 계약 추진"""
    },
    {
        "folder": "Naver_News",
        "filename": "2024-01-17_정부정책_발표.txt",
        "title": "[정책] 정부 K-배터리 지원책 발표",
        "date": "2024-01-17",
        "category": "정책",
        "content": """[정책 동향] 정부 K-배터리 육성 지원책 발표

산업통상자원부는 K-배터리 산업 육성을 위한 종합 지원책을 발표했습니다.

주요 내용:
- R&D 예산 1조원 투입 (3년간)
- 세제 혜택 확대 (투자세액공제율 상향)
- 전문 인력 양성 (연간 1,000명)
- 규제 완화 (인허가 기간 단축)

기대 효과:
- 기술 개발 가속화
- 투자 여건 개선
- 글로벌 경쟁력 강화"""
    }
]

def create_mock_files_with_metadata():
    """Create mock files with metadata JSON"""
    
    # Create directories if not exist
    os.makedirs(BASE_PATH + "internal", exist_ok=True)
    os.makedirs(BASE_PATH + "external", exist_ok=True)
    
    # Create metadata storage
    metadata_all = {
        "internal": [],
        "external": []
    }
    
    # Generate internal documents
    for doc in INTERNAL_DOCS:
        folder_path = os.path.join(BASE_PATH, "internal", doc["folder"])
        os.makedirs(folder_path, exist_ok=True)
        
        file_path = os.path.join(folder_path, doc["filename"])
        
        # Write content
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(doc["content"])
        
        # Store metadata
        metadata_all["internal"].append({
            "file_path": file_path,
            "title": doc["title"],
            "organization": doc["folder"],
            "created_at": doc["date"],
            "type": "internal",
            "category": doc["folder"]
        })
        
        print(f"Created: {file_path}")
    
    # Generate external news
    for news in EXTERNAL_NEWS:
        folder_path = os.path.join(BASE_PATH, "external", news["folder"])
        os.makedirs(folder_path, exist_ok=True)
        
        file_path = os.path.join(folder_path, news["filename"])
        
        # Write content
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(news["content"])
        
        # Store metadata
        metadata_all["external"].append({
            "file_path": file_path,
            "title": news["title"],
            "organization": news["folder"],
            "created_at": news["date"],
            "type": "external",
            "category": news["category"]
        })
        
        print(f"Created: {file_path}")
    
    # Save metadata JSON
    metadata_path = os.path.join(BASE_PATH, "metadata.json")
    with open(metadata_path, 'w', encoding='utf-8') as f:
        json.dump(metadata_all, f, ensure_ascii=False, indent=2)
    
    print(f"\nMetadata saved to: {metadata_path}")
    print(f"Total files created: {len(metadata_all['internal']) + len(metadata_all['external'])}")

if __name__ == "__main__":
    create_mock_files_with_metadata()