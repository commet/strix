# STRIX Dashboard 설정 가이드

## 1. Dashboard 활성화 단계

### Step 1: VBA 편집기 열기
1. STRIX_Dashboard.xlsm 파일 열기
2. Alt + F11 눌러서 VBA 편집기 열기

### Step 2: 필수 모듈 가져오기
VBA 편집기에서 파일 > 가져오기:
1. `modules\modSTRIXAPI.bas`
2. `modules\modSTRIXExcel.bas` 
3. `modules\frmSTRIX.frm`
4. `JsonConverter.bas` (루트 폴더에 있음)

### Step 3: 참조 설정
도구 > 참조에서 체크:
- Microsoft XML, v6.0
- Microsoft Scripting Runtime

### Step 4: API 서버 실행
명령 프롬프트에서:
```
cd C:\Users\admin\documents\github\strix
py -m streamlit run streamlit_app_with_api.py
```

### Step 5: Dashboard 시트 설정
1. Sheet1로 이동
2. 다음 버튼들 추가:
   - STRIX 열기 버튼 → 매크로: ShowSTRIXDialog
   - 선택 분석 버튼 → 매크로: AskAboutSelection
   - 검색 결과 버튼 → 매크로: DisplaySearchResults

### Step 6: 테스트
셀에 다음 수식 입력:
```
=STRIX("전고체 배터리 개발 현황은?")
```

## 2. Dashboard 레이아웃 예시

### A1:D1 - 제목
"STRIX Intelligence Dashboard"

### A3:D3 - 버튼 영역
- [STRIX 대화창] [선택 분석] [문서 검색] [검색 기록]

### A5:D20 - 질문/답변 영역
질문 입력 셀: B5
답변 표시 영역: B7:D20

### F1:J20 - 최근 검색 결과
검색 기록 자동 표시

## 3. 문제 해결

### "서버 연결 실패" 오류
1. Streamlit 서버가 실행 중인지 확인
2. http://localhost:8501 접속 테스트

### VBA 매크로 보안 경고
1. 파일 > 옵션 > 보안 센터
2. 매크로 설정: "모든 매크로 사용"

### 한글 깨짐
1. Windows 제어판 > 지역 및 언어
2. 관리 탭 > 시스템 로캘 변경
3. "Beta: 세계 언어 지원을 위해 UTF-8 사용" 체크