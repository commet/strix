# STRIX Issue Tracker 빠른 시작 가이드

## 1단계: 데이터베이스 설정

### Supabase SQL Editor에서 실행:
```sql
-- create_issue_tracking_tables.sql 파일 내용을 복사하여 실행
```

## 2단계: API 서버 실행

```bash
# Issue Tracking이 통합된 API 서버 실행
python api_server_with_issues.py
```

서버가 http://localhost:5000 에서 실행됩니다.

## 3단계: Excel에서 Issue Timeline 대시보드 생성

### 방법 1: VBA 매크로 직접 실행
1. Excel 파일 열기 (STRIX_System.xlsm)
2. Alt+F11로 VBA 편집기 열기
3. modIssueTimeline 모듈 찾기
4. CreateIssueTimelineDashboard 서브루틴 실행 (F5)

### 방법 2: 버튼으로 실행
1. Excel에서 개발 도구 탭 활성화
2. 삽입 > 단추 추가
3. CreateIssueTimelineDashboard 매크로 할당

## 4단계: 이슈 추적 기능 사용

### 📊 이슈 타임라인 보기
- Issue Timeline 시트에서 이슈 진행 상황 확인
- 각 이슈의 타임라인 시각화
- 문서별 언급 추적

### 🔍 이슈 필터링
- 카테고리별: 전략, 기술, 리스크, 경쟁사, 정책
- 상태별: 미해결, 진행중, 해결됨, 모니터링
- 기간별: 최근 3개월, 6개월, 1년

### 🤖 AI 분석 실행
- "AI 분석" 버튼 클릭
- 미해결 이슈에 대한 예측 생성
- 권장 액션 아이템 확인

### 🔄 데이터 새로고침
- "새로고침" 버튼으로 최신 데이터 로드
- API에서 실시간 이슈 정보 가져오기

## 주요 API 엔드포인트

### 이슈 목록 조회
```
GET /api/issues?category=기술&status=진행중&days=90
```

### 이슈 상세 정보
```
GET /api/issues/{issue_id}
```

### 이슈 타임라인
```
GET /api/issues/{issue_id}/timeline
```

### AI 분석
```
GET /api/issues/{issue_id}/ai-analysis
```

### 대시보드 요약
```
GET /api/issues/dashboard-summary
```

### 문서에서 이슈 추출
```
POST /api/issues/extract
Body: {"document_id": "doc-123"}
```

### AI 예측 생성
```
POST /api/issues/predict
```

## 테스트 시나리오

### 1. 새 문서에서 이슈 추출
```python
import requests

# 문서에서 이슈 자동 추출
response = requests.post('http://localhost:5000/api/issues/extract', 
    json={'document_id': 'your-document-id'})
```

### 2. Excel에서 이슈 목록 업데이트
VBA 편집기에서 실행:
```vba
Call UpdateIssueTimeline
```

### 3. 특정 이슈 상세 보기
타임라인에서 이슈 클릭하면 자동으로 상세 정보 표시

## 문제 해결

### API 서버 연결 실패
- localhost:5000이 실행 중인지 확인
- 방화벽 설정 확인

### 한글 깨짐 문제
- UTF-8 인코딩 확인
- Excel VBA에서 BytesToString 함수 사용

### 데이터베이스 오류
- Supabase 연결 정보 확인 (.env 파일)
- 테이블 생성 여부 확인

## 추가 기능

### 이슈 상태 자동 업데이트
- 문서에서 해결 언급시 자동으로 상태 변경
- 일정 기간 업데이트 없으면 모니터링으로 전환

### 이슈 간 연관관계 분석
- 의존성 파악
- 블로킹 이슈 식별
- 중복 이슈 병합

### 예측 기반 알림
- 리스크 임계값 도달시 알림
- 마감일 임박 알림
- 의사결정 필요 시점 알림