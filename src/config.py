"""
STRIX Configuration
"""
import os
from dotenv import load_dotenv

load_dotenv()

# Supabase Configuration
SUPABASE_URL = "https://qxrwyfxwwihskktsmjhj.supabase.co"
SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InF4cnd5Znh3d2loc2trdHNtamhqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTQyODY2MTEsImV4cCI6MjA2OTg2MjYxMX0.RYFV2PFIk6i-Se9Y3MfFbfR8Yz7R9_PzGeGC0F3IIqA"
SUPABASE_SERVICE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InF4cnd5Znh3d2loc2trdHNtamhqIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc1NDI4NjYxMSwiZXhwIjoyMDY5ODYyNjExfQ.nmkCPDvG4Os-Bez9Lcz8MHHP-KlViG6FZCUR2vPTryk"

# OpenAI Configuration (사외 테스트용)
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# Document Processing
CHUNK_SIZE = 1000
CHUNK_OVERLAP = 200

# Categories
INTERNAL_CATEGORIES = ["전략기획", "R&D", "경영지원", "생산", "영업마케팅"]
EXTERNAL_CATEGORIES = ["Macro", "산업", "기술", "리스크", "경쟁사", "정책"]

# Keywords - 확장된 버전
CATEGORY_KEYWORDS = {
    "Macro": [
        "경제", "금리", "환율", "인플레이션", "GDP", "무역", "수출입", 
        "경기", "침체", "성장률", "통화정책", "재정정책", "원자재가격",
        "유가", "달러", "위안화", "엔화", "유로", "연준", "한은",
        "미국", "중국", "유럽", "일본", "신흥국", "공급망", "지정학"
    ],
    "산업": [
        "배터리", "전기차", "반도체", "에너지", "리튬", "니켈", "코발트",
        "양극재", "음극재", "분리막", "전해질", "셀", "모듈", "팩",
        "ESS", "BMS", "충전인프라", "완성차", "OEM", "Tier1",
        "수산화리튬", "탄산리튬", "NCM", "NCA", "LFP", "전구체"
    ],
    "기술": [
        "AI", "자동화", "디지털", "혁신", "전고체", "차세대배터리",
        "리튬메탈", "실리콘음극", "NMC", "LFP", "NCM811", "4680",
        "빅데이터", "IoT", "디지털트윈", "스마트팩토리", "로봇",
        "고니켈", "무코발트", "건식전극", "프리리튬화", "고전압"
    ],
    "리스크": [
        "규제", "제재", "사고", "리콜", "화재", "안전", "환경",
        "ESG", "탄소배출", "폐배터리", "재활용", "공급부족", "가격변동",
        "품질", "불량", "소송", "특허", "덤핑", "무역분쟁", "보조금"
    ],
    "경쟁사": [
        "CATL", "BYD", "Tesla", "Panasonic", "LG에너지솔루션", "삼성SDI",
        "SK온", "Northvolt", "CALB", "Gotion", "EVE", "SVOLT",
        "Contemporary Amperex", "비야디", "테슬라", "파나소닉"
    ],
    "정책": [
        "IRA", "CBAM", "RE100", "탄소중립", "그린뉴딜", "전기차보조금",
        "배터리규제", "Euro7", "CARB", "ZEV", "친환경", "순환경제",
        "배터리여권", "LCA", "EPR", "RoHS", "REACH", "분류체계"
    ]
}

# 동의어 매핑 (검색 성능 향상용)
SYNONYM_MAP = {
    "전고체": ["전고체배터리", "고체전해질", "solid-state", "SSB"],
    "전기차": ["EV", "Electric Vehicle", "BEV", "전동차"],
    "리튬이온": ["Li-ion", "LIB", "리튬이온전지"],
    "CATL": ["Contemporary Amperex", "寧德時代", "닝더스다이"],
    "배터리": ["전지", "이차전지", "Battery", "Cell"]
}

# Mock Data Path (개발용)
MOCK_DATA_PATH = "./mock_data"

# LangChain Settings
DEFAULT_MODEL = "gpt-4o"
EMBEDDING_MODEL = "text-embedding-ada-002"
TEMPERATURE = 0.3