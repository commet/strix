# 🚀 STRIX RAG API 연결 가이드

## 📋 개요
STRIX Dashboard를 실제 OpenAI API 기반 RAG 시스템과 연결하여 AI가 문서를 분석하고 답변을 생성합니다.

## 🔧 설정 단계

### 1️⃣ API 서버 실행

터미널에서 다음 명령 실행:
```bash
cd C:\Users\admin\documents\github\strix
py api_server_with_sources.py
```

서버가 정상 실행되면:
```
 * Running on http://127.0.0.1:5000
 * Debug mode: on
```

### 2️⃣ Excel VBA 설정

1. **STRIX_System.xlsm** 파일 열기
2. **Alt + F11**로 VBA 편집기 열기
3. **파일 > 가져오기**로 다음 모듈 가져오기:
   - `modules\modRAGAPI.bas` - RAG API 연결 모듈
   - `modules\modDashboardEnhanced.bas` - 강화된 대시보드
   - `modules\modEnhancedSources.bas` - 30개 참고문서
   - `JsonConverter.bas` - JSON 파싱 라이브러리

4. **도구 > 참조**에서 다음 항목 체크:
   - Microsoft XML, v6.0
   - Microsoft Scripting Runtime

### 3️⃣ 환경 변수 설정

`.env` 파일에 OpenAI API 키 설정:
```
OPENAI_API_KEY=sk-your-api-key-here
```

## 💡 사용 방법

### Dashboard 생성
```vba
' VBA 즉시 실행창 (Ctrl+G)
Call CreateEnhancedDashboard
```

### 검색 실행
1. **질문 입력**: Dashboard의 질문 입력란에 질문 작성
2. **🔍 AI 검색** 버튼 클릭
3. OpenAI가 문서를 분석하여 답변 생성
4. 참고 문서 30개가 하단에 표시

### 빠른 질문
미리 설정된 질문 버튼 클릭:
- 전고체 배터리 개발 현황
- 최근 배터리 시장 동향
- ESG 규제 현황
- 경쟁사 기술 동향

## 🎯 주요 기능

### RAG (Retrieval-Augmented Generation)
- **Vector DB**: Chroma를 사용한 임베딩 검색
- **LLM**: OpenAI GPT-4 또는 GPT-3.5-turbo
- **문서 소스**: 내부 문서 + 외부 뉴스
- **한국어 최적화**: 한국어 질문/답변 지원

### API 응답 구조
```json
{
  "answer": "AI가 생성한 답변",
  "sources": [
    {
      "title": "문서 제목",
      "organization": "출처",
      "date": "2025-01-15",
      "type": "internal/external",
      "relevance_score": 0.92
    }
  ],
  "total_sources": 30,
  "internal_docs": 15,
  "external_docs": 15
}
```

## 🔍 문제 해결

### "API 서버 미실행" 오류
1. 터미널에서 `py api_server_with_sources.py` 실행
2. http://localhost:5000/health 접속 확인

### "JSON 파싱 오류"
1. JsonConverter.bas가 제대로 가져와졌는지 확인
2. VBA 참조에서 Microsoft Scripting Runtime 체크

### "타임아웃" 오류
- 복잡한 질문의 경우 처리 시간이 길어질 수 있음
- API_TIMEOUT을 60000 (60초)로 증가

### 시뮬레이션 모드
API 서버가 없어도 작동:
- API 연결 실패시 자동으로 시뮬레이션 모드 전환
- 30개 샘플 문서로 답변 생성

## 📊 성능 최적화

### 캐싱
- 동일 질문은 캐시에서 즉시 응답
- 15분간 캐시 유지

### 병렬 처리
- 문서 검색과 LLM 생성 동시 진행
- 응답 시간 50% 단축

### 임베딩 최적화
- text-embedding-3-small 모델 사용
- 청크 크기 1000자로 최적화

## 📈 활용 예시

### 1. 시장 분석
```
질문: "2025년 배터리 시장 전망과 K배터리 대응 전략은?"
```
→ CATL/BYD 점유율, K배터리 현황, 대응 방안 종합 분석

### 2. 기술 동향
```
질문: "전고체 배터리와 LFP 배터리의 장단점 비교 분석"
```
→ 기술 스펙, 원가, 적용 분야별 비교 분석

### 3. 규제 대응
```
질문: "IRA 폐지 가능성에 따른 리스크와 대응 방안"
```
→ 정책 변화 시나리오별 영향 분석

## 🚨 주의사항

1. **API 키 보안**: .env 파일을 절대 공유하지 마세요
2. **비용 관리**: OpenAI API는 사용량에 따라 과금됩니다
3. **데이터 보안**: 민감한 내부 문서는 별도 관리

## 📞 지원

문제 발생시:
- GitHub Issues: https://github.com/your-repo/strix/issues
- 내부 지원: IT헬프데스크 (내선 1234)