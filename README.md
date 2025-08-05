# STRIX - Strategic Intelligence System

## 빠른 시작

### 1. 환경 변수 설정
`.env` 파일을 생성하고 OpenAI API 키를 추가하세요:
```
OPENAI_API_KEY=your-api-key-here
```

### 2. 의존성 설치
```bash
pip install -r requirements.txt
```

### 3. 데이터베이스 초기화
Supabase SQL Editor에서 `setup_supabase.sql` 파일의 내용을 실행하세요.

### 4. 테스트 데이터 생성 및 인제스트
```bash
# Mock 데이터 생성
python create_mock_data.py

# 문서 인제스트
python test_ingestion.py --auto
```

### 5. Streamlit 앱 실행
```bash
streamlit run streamlit_app.py
```

## 주요 기능
- 내부 문서와 외부 뉴스 통합 검색
- RAG 기반 지능형 답변 생성
- 문서 간 상관관계 분석
- 실시간 인사이트 제공

## 시스템 구조
```
STRIX/
├── src/
│   ├── config.py           # 설정 파일
│   ├── database/          # Supabase 클라이언트
│   ├── document_loaders/  # 문서 로더
│   └── rag/              # RAG 체인
├── streamlit_app.py       # 웹 UI
└── test_*.py             # 테스트 스크립트
```