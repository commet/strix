# STRIX 사외 개발 가이드: Mock 데이터 기반 구현 전략

## 개요
사내 데이터 반출이 불가능한 상황에서 STRIX를 사외에서 개발하기 위한 단계별 구현 가이드입니다. 실제 사내 환경과 동일한 제약조건(VBA only, 외부 API 차단)을 유지하면서 Mock 데이터로 개발·테스트를 진행합니다.

## 1. 개발 환경 셋업 & 파일 구조 준비

### 1.1 Excel 템플릿 생성
```
STRIX_Template.xlsm
├─ Config (설정 시트)
├─ RawData (내부 보고자료 메타)
├─ RawNews (외부 뉴스 데이터)
├─ MetaData (이슈 마스터)
├─ LinkedNews (연관도 매핑)
├─ Dashboard (시각화)
├─ GPT_Interface (프롬프트/응답)
└─ NewsletterTemplate (뉴스레터 양식)
```

### 1.2 Mock 데이터 준비
#### Mock 보고자료 (내부)
```
Mock_Reports/
├─ 2024_Q1_전략기획_배터리사업현황.pptx
├─ 2024_Q1_R&D_신기술개발진척.xlsx
├─ 2024_Q2_경영지원_리스크관리.pdf
└─ ... (10-15개 파일)
```

#### Mock Outlook 뉴스 (외부)
```
Mock_Outlook/
├─ AM_News_2024_01_15.msg
├─ PM_News_2024_01_15.msg
└─ ... (20-30개 메일 파일)
```

### 1.3 VBA 프로젝트 초기 구조
```vb
' modConfig.bas
Public Const cstMockDataPath = "C:\STRIX_Mock\"
Public gblLastScanTime As Date
Public gblIsLocked As Boolean

' modInit.bas
Sub InitializeSTRIX()
    ' 시트 구조 검증
    ' Mock 경로 설정
    ' 초기 데이터 로드
End Sub
```

## 2. Phase 1: 핵심 데이터 적재 (1주차)

### 2.1 내부 Intelligence 스캔 (modInternalIngest)
```vb
Sub ScanInternalFolder()
    Dim fso As Object, folder As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Mock_Reports 폴더 스캔
    Set folder = fso.GetFolder(cstMockDataPath & "Mock_Reports")
    
    For Each file In folder.Files
        ' 파일명에서 메타데이터 추출 (날짜, 조직, 주제)
        ' RawData 시트에 행 추가
    Next file
End Sub
```

### 2.2 외부 Intelligence 스캔 (modExternalIngest)
```vb
Sub ScanOutlookFolder()
    ' Mock .msg 파일 읽기 시뮬레이션
    ' 실제로는 텍스트 파일로 대체
    Dim mockNewsPath As String
    mockNewsPath = cstMockDataPath & "Mock_Outlook\"
    
    ' 각 뉴스 파일 파싱
    ' RawNews 시트에 적재
End Sub
```

### 2.3 메타데이터 입력 폼 (frmMetaEntry)
- 이슈ID, 이슈명, 조직코드, 키워드 입력
- 성공사례/경영진관심 체크박스
- MetaData 시트에 저장

**테스트 포인트:**
- Mock 파일이 정상적으로 스캔되는지
- 메타데이터 입력 후 저장이 원활한지
- 증분 스캔이 중복 없이 작동하는지

## 3. Phase 2: Correlation & 기본 리포트 (2주차)

### 3.1 연관도 계산 함수
```vb
Function CalcCorrelation(issueID As String, newsText As String) As Double
    ' Mock 데이터로 키워드 매칭 테스트
    ' 예: "배터리", "전고체", "리스크" 등
End Function
```

### 3.2 매핑 실행
```vb
Sub LinkNewsToIssues()
    ' 모든 뉴스 × 모든 이슈 조합 계산
    ' 임계값(0.2) 이상만 LinkedNews에 기록
End Sub
```

### 3.3 간이 리포트 생성
- 이슈별 연관 뉴스 건수
- 조직별 이슈 현황
- 기간별 트렌드

**테스트 시나리오:**
1. "배터리 신기술" 이슈 ↔ "전고체 배터리 개발" 뉴스 매칭
2. "리스크 관리" 이슈 ↔ "규제 강화" 뉴스 매칭
3. 임계값 조정에 따른 매칭 결과 변화

## 4. Phase 3: Dashboard & Q&A (3주차)

### 4.1 시각화 구현
- PivotTable: 이슈×기간×조직 크로스탭
- 타임라인 차트: 내부보고 vs 외부뉴스 시계열
- 슬라이서: 동적 필터링

### 4.2 GPT Prompt Builder
```vb
Sub BuildPrompt()
    Dim prompt As String
    prompt = "다음 기간의 주요 이슈를 요약해주세요: " & _
             Range("Dashboard!B1").Value & " ~ " & _
             Range("Dashboard!B2").Value
    
    ' 클립보드 복사
    CopyToClipboard prompt
End Sub
```

### 4.3 응답 수집 시뮬레이션
- Mock GPT 응답 텍스트 준비
- PasteGPTResponse로 붙여넣기 테스트

## 5. Phase 4: Newsletter & 자동화 (4주차)

### 5.1 뉴스레터 템플릿
- 주간 하이라이트 섹션
- 이슈별 요약
- 차트 스냅샷 삽입

### 5.2 Outlook 발송 시뮬레이션
```vb
Sub CreateNewsletterDraft()
    ' Mock 메일 객체 생성
    ' HTML 본문 구성
    ' 첨부파일 추가
End Sub
```

### 5.3 자동화 트리거
- Workbook_Open 이벤트
- OnTime 스케줄링
- 리본 커스텀 버튼

## 6. Mock 데이터 설계 가이드

### 6.1 내부 보고자료 Mock
```
파일명 규칙: YYYY_Qn_조직명_주제.확장자
내용 구성:
- 이슈ID: ISS-2024-001
- 키워드: 배터리, 신기술, 투자
- 경영진 관심: Y/N
```

### 6.2 외부 뉴스 Mock
```
제목: [Macro] 글로벌 배터리 시장 동향
본문: 전고체 배터리 기술이 차세대 핵심으로...
날짜: 2024-01-15
카테고리: Macro, 산업, 기술
```

### 6.3 연관도 테스트 케이스
| 이슈 | 뉴스 | 예상 점수 |
|------|------|-----------|
| 배터리 신기술 개발 | 전고체 배터리 상용화 | 0.8 |
| 리스크 관리 체계 | EU 배터리 규제 강화 | 0.6 |
| 생산 효율화 | 스마트팩토리 도입 | 0.4 |

## 7. 사내 환경 이관 시 체크리스트

1. **경로 변경**
   - Mock 경로 → 실제 공유 폴더
   - Outlook 폴더명 확인

2. **권한 확인**
   - 공유 폴더 읽기/쓰기
   - Outlook 폴더 접근

3. **데이터 마이그레이션**
   - Config 설정값
   - 키워드 사전
   - 이슈 마스터

4. **성능 최적화**
   - 실제 데이터 볼륨 대응
   - 스캔 주기 조정

## 8. 주의사항 & 팁

### VDI 환경 시뮬레이션
- 외부 네트워크 차단 가정
- ActiveX 컨트롤 제한
- 파일 시스템 접근만 허용

### 코드 이식성
- Late Binding 사용 (CreateObject)
- 절대 경로 대신 Config 참조
- 에러 핸들링 강화

### 테스트 자동화
```vb
Sub RunAllTests()
    TestInternalScan
    TestExternalScan
    TestCorrelation
    TestDashboard
    TestNewsletter
End Sub
```

## 9. Q&A 대비 포인트

**Q: Mock 데이터가 실제와 얼마나 유사해야 하나요?**
A: 파일명 규칙, 키워드 분포, 데이터 볼륨이 유사하면 충분합니다.

**Q: VDI 제약을 어떻게 시뮬레이션하나요?**
A: 인터넷 차단 상태에서 개발하고, VBA 표준 라이브러리만 사용합니다.

**Q: 사내 이관 시 가장 주의할 점은?**
A: Outlook 버전 차이, 공유 폴더 권한, 실제 데이터 인코딩 문제입니다.

---

이 가이드를 따라 단계별로 진행하시면, 사외에서도 안정적으로 STRIX를 개발하실 수 있습니다. 각 Phase마다 Mock 데이터로 충분히 테스트한 후 다음 단계로 진행하시길 권장합니다.