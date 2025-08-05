# STRIX: 전략 인텔리전스 통합 플랫폼 기획서

## 1. 배경 및 문제 인식

### 1) 정보 단절과 사일로 현상
- **보고자료·회의록·이슈는 각 조직별로 산재**해 관리가 어렵고, 누적 이력 파악이 불투명
- 사내 PR팀의 Macro·배터리 산업 뉴스 메일도 **매일 2회 수동 확인에 그쳐**, 사내 이슈와의 연계가 부실

### 2) 경영진 의사결정 지원 한계
- 보고서·회의록만으로는 **전사 주요 이슈/사례와 경영진 관심사항이 즉각적으로 공유**되기 어려움
- **외부 시장·정책 변화와 내부 실행 현황을 동시에 고려한 통합 인사이트가 부재**

### 3) 보고 위한 수작업의 비효율성
- 과거 보고자료 수동 검색, 메일 본문 복사·붙여넣기, 보고서 작성용 정리 등에 **과도한 시간 소요**
- 외부 네트워크 차단 환경 및 사내 API 활용 불가 등 개발자 도구 활용 제약下 **자동화 구현이 제한적**

## 2. 목표 (Mission)

> **내·외부 정보를 단일 창에서 모니터링·분석하여, 전략실·경영진의 의사결정을 종합 지원하는 인텔리전스 플랫폼 구축 (VBA 기반)**

## 3. 핵심 가치 제안 (Value Proposition)

### 1) 통합 아카이빙
- 사내 이슈(보고·성공사례·경영진 관심 포인트)와 메일 뉴스(Macro·산업·리스크)를 **동일 DB에 누적**

### 2) 병렬적 정보 추적
- **내부 Intelligence**: 조직별 주요 이슈, 성공사례, 경영진 관심 토픽
- **외부 Intelligence**: 시장·정책·경쟁사 뉴스, Macro 트렌드

### 3) 연계 인사이트
- 내부 이슈와 외부 뉴스 간 **자동 연관도 계산**으로 사각지대 해소
- "우리 주요 성공사례와 연관된 외부 긍정 이슈는?" **즉시 답변 지원**

## 4. 모듈별 주요 기능

### 1) Config 모듈
- 공유 폴더 경로(Teamroom에 전략실 구성원 권한 부여 고려 중), Outlook 폴더, 카테고리·키워드 관리
- 마지막 스캔 시각·잠금 설정

### 2) 내부 Intelligence Tracking
- **폴더 스캔**: 보고자료(PPT/Excel/PDF) 메타 자동 추출
- **UserForm 태깅**: 이슈ID, 조직, 성공사례 여부, 경영진 관심 포인트
- **증분 로드**: 최종 처리 이후 신규 파일만 동기화

### 3) 외부 Intelligence Tracking
- **Outlook VBA**: 지정 폴더("내부뉴스_AM/PM", "외부뉴스_Google/Naver..") 스캔 (*VDI 내 Outlook 접근 제한 시 다른 접근 필요)
- **본문·첨부**: 뉴스 본문, 링크, PDF 첨부 저장
- **자동 분류**: Macro·산업·리스크·경쟁사 키워드 태그

### 4) Correlation
- 내부 이슈명 vs. 뉴스 키워드 텍스트 매칭
- 연관도 지수 등 계산 후 LinkedNews 시트에 자동 기록, UserForm에서 최종 확인·수정

### 5) Dashboard & Q&A
- **타임라인 차트**: 내부 보고·외부 뉴스 동시 시각화
- **슬라이서**: 조직·이슈·카테고리·기간 필터
- **Prompt 빌더**: 선택된 뷰 기반으로 GPT 질의문 자동 생성

### 6) Newsletter & Export
- **템플릿 기반**: 주간 인사이트 자동 생성
- **Outlook 발송**: VBA로 메일 초안 작성·발송
- **PDF Export**: 경영진 보고용 원페이저 생성

## 5. 주요 Feature Code 예시

### 1) 사내 GPT 반자동 연동 (Prompt 빌더 & 응답 파싱)

```vb
'— 참조: Microsoft Forms 2.0 Object Library 필요
Sub BuildAndCopyPrompt()
    Dim promptText As String
    Dim ds As MSForms.DataObject
    
    ' 1) 현재 Dashboard 필터 기준으로 Prompt 생성
    promptText = "Please summarize internal issues and related external news from " & _
                 Sheets("Dashboard").Range("B1").Value & " to " & _
                 Sheets("Dashboard").Range("B2").Value & "."
    
    ' 2) 클립보드에 복사
    Set ds = New MSForms.DataObject
    ds.SetText promptText
    ds.PutInClipboard
    
    MsgBox "Prompt has been copied to clipboard." & vbCrLf & _
           "→ Paste into your internal GPT console (Ctrl+V) and get response."
End Sub

Sub PasteGPTResponse()
    Dim ds As MSForms.DataObject
    Dim response As String
    Dim rw As Long
    
    ' 클립보드에서 GPT 응답 읽어오기
    Set ds = New MSForms.DataObject
    ds.GetFromClipboard
    response = ds.GetText
    
    ' GPT_Interface 시트에 새 행으로 삽입
    With Sheets("GPT_Interface")
        rw = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        .Cells(rw, 1).Value = Now           ' PromptID 대용
        .Cells(rw, 2).Value = ""            ' (원문 저장 칼럼)
        .Cells(rw, 3).Value = response
    End With
    
    MsgBox "GPT response pasted into GPT_Interface."
End Sub
```

- **BuildAndCopyPrompt**: Dashboard 상 필터(날짜, 이슈 등)를 조합해 Prompt 생성 → 클립보드 복사 → 사용자 수동 붙여넣기
- **PasteGPTResponse**: GPT 콘솔에서 복사한 응답을 다시 클립보드에서 읽어와 시트에 자동 저장
- 추후 해당 MVP 시연 등을 통해 사내 DT 조직에 건의하여 **GPT API Key 권한 부여를 요청할 계획** (미정)

### 2) 내·외부 Intelligence Correlation 로직

```vb
Function CalcCorrelation(issueID As String, newsText As String) As Double
    Dim keywords As Variant, kw As Variant
    Dim score As Long: score = 0
    
    ' MetaData 시트에서 해당 이슈 키워드 배열 로드
    keywords = Application.Transpose( _
        Sheets("MetaData").ListObjects("MetaData_tbl") _
              .ListColumns("Keywords").DataBodyRange _
    )
    
    ' 단순 키워드 빈도 집계
    For Each kw In keywords
        If Len(kw) > 0 Then
            score = score + UBound(Split(LCase(newsText), LCase(kw)))
        End If
    Next kw
    
    ' 정규화: (키워드 매칭 건수) / (총 키워드 수)
    If UBound(keywords) >= 0 Then
        CalcCorrelation = score / (UBound(keywords) + 1)
    Else
        CalcCorrelation = 0
    End If
End Function

Sub LinkNewsToIssues()
    Dim nr As ListRow, score As Double
    Dim newsRow As ListRow
    
    For Each newsRow In Sheets("RawNews").ListObjects("RawNews_tbl").ListRows
        For Each nr In Sheets("MetaData").ListObjects("MetaData_tbl").ListRows
            score = CalcCorrelation(nr.Range.Cells(1, "IssueID"), newsRow.Range.Cells(1, "BodyText"))
            If score >= 0.2 Then   ' 임계값 예시
                ' LinkedNews 시트에 매핑 기록
                With Sheets("LinkedNews").ListObjects("LinkedNews_tbl")
                    .ListRows.Add
                    .ListRows(.ListRows.Count).Range.Value = Array( _
                        nr.Range.Cells(1, "IssueID"), _
                        newsRow.Range.Cells(1, "MailID"), _
                        score, "Y" _
                    )
                End With
            End If
        Next nr
    Next newsRow
End Sub
```

- **CalcCorrelation**: 각 이슈에 미리 정리한 키워드(Keywords 컬럼)와 뉴스 본문을 비교해 키워드 매칭 비율 계산
- **LinkNewsToIssues**: 모든 뉴스 ↔ 모든 이슈에 대해 CalcCorrelation 실행 → 임계값 이상일 때 LinkedNews 테이블에 자동 등록
- 사내 GPT 활용하여 **단순 키워드 빈도 매칭을 넘어 Logic 고도화 예정**

## 6. Project Architecture

```
┌──────────────────────┐
│     Config 모듈       │
│ (경로, 키워드, 설정)  │
└──────────┬───────────┘
           │
    ┌──────┴──────┐
    ▼             ▼
┌─────────────┐  ┌─────────────┐
│   내부       │  │   외부       │
│Intelligence  │  │Intelligence  │
│  Tracking    │  │  Tracking    │
├─────────────┤  ├─────────────┤
│• 보고자료    │  │• Outlook     │
│  스캔        │  │  메일 스캔   │
│• 메타데이터  │  │• 본문·첨부   │
│  등록        │  │  저장        │
│• 이슈·조직   │  │• 키워드      │
│  태깅        │  │  기반 분류   │
└──────┬──────┘  └──────┬──────┘
       │                 │
       └────────┬────────┘
                ▼
        ┌───────────────┐
        │  Correlation  │
        │ (매칭·연관도) │
        └───────┬───────┘
                ▼
        ┌───────────────┐
        │  Dashboard &  │
        │     Q&A       │
        └───────┬───────┘
                ▼
        ┌───────────────┐
        │  Newsletter & │
        │   Reporting   │
        └───────────────┘
```

## 7. Claude Code를 활용한 개발 로드맵

### 개발 방식
본 프로젝트는 Claude Code를 활용하여 VBA 코드를 자동 생성하고, Mock 데이터로 테스트를 진행합니다.

### Phase 1: 기초 설정 (1-2시간)
#### 구현 모듈
- **modConfig.bas**: 전역 변수, 상수, 설정 관리
- **modInit.bas**: 초기화, 시트 생성, 검증
- **Mock 데이터 생성기**: 테스트용 파일 자동 생성

#### 주요 기능
- Excel 템플릿 생성 (Config, RawData, RawNews 등 9개 시트)
- 전역 설정 변수 관리
- Mock 데이터 폴더 구조 생성

### Phase 2: 데이터 수집 (2-3시간)
#### 구현 모듈
- **modInternalIngest.bas**: 폴더 스캔, 파일 메타데이터 추출
- **modExternalIngest.bas**: Outlook 시뮬레이션, 뉴스 파싱
- **frmMetaEntry.frm**: 메타데이터 입력 UserForm

#### 주요 기능
- PPT/Excel/PDF 파일 스캔 및 메타데이터 추출
- Mock 뉴스 메일 파싱 (제목, 본문, 첨부파일)
- UserForm을 통한 이슈 태깅 (조직, 성공사례, 경영진 관심)

### Phase 3: 데이터 처리 (2시간)
#### 구현 모듈
- **modCorrelation.bas**: 연관도 계산 알고리즘
- **modClassification.bas**: 카테고리 자동 분류
- **frmCategoryAdjust.frm**: 분류 보정 UI

#### 주요 기능
- 키워드 기반 내부 이슈-외부 뉴스 매칭
- 연관도 점수 계산 (임계값 0.2)
- 자동 분류 후 수동 보정 인터페이스

### Phase 4: 시각화 & 리포팅 (2시간)
#### 구현 모듈
- **modDashboard.bas**: PivotTable, 차트 생성
- **modPromptBuilder.bas**: GPT 프롬프트 생성
- **modNewsletter.bas**: 뉴스레터 템플릿

#### 주요 기능
- 타임라인 차트 (내부 보고 vs 외부 뉴스)
- 동적 슬라이서 필터
- GPT 프롬프트 자동 생성 및 클립보드 복사

### Phase 5: 자동화 & 최적화 (1시간)
#### 구현 모듈
- **modScheduler.bas**: 자동 실행 스케줄러
- **modErrorHandler.bas**: 에러 처리
- **modConcurrency.bas**: 동시 접근 관리

#### 주요 기능
- Workbook_Open 이벤트 자동 실행
- OnTime 스케줄링 (매일 09:00, 18:00)
- 잠금 충돌 처리 및 로깅

### 구현 일정
- **총 소요 시간**: 7-10시간
- **일일 작업량**: 2-3시간씩 3-4일
- **테스트 포함**: 각 Phase별 Mock 데이터 테스트

## 8. 기대 효과

1. **정보 통합 관리**: 흩어진 내·외부 정보를 한 곳에서 관리
2. **의사결정 지원**: 경영진에게 시의적절한 인사이트 제공
3. **업무 효율화**: 수작업 시간 80% 이상 단축
4. **선제적 대응**: 외부 트렌드와 내부 이슈의 연계 분석으로 리스크 조기 감지

## 9. 향후 확장 계획

- **API 통합**: 사내 GPT API Key 확보 시 완전 자동화
- **대용량 최적화**: ADODB + Access 연동으로 대용량 데이터 처리
- **다국어 지원**: 글로벌 뉴스 및 리포트 통합 분석
- **AI 고도화**: 단순 키워드 매칭을 넘어 의미 기반 연관도 분석

---

*이 프로젝트는 외부 네트워크 차단 환경에서도 구동 가능한 VBA 기반 솔루션으로, 외부에서 구현된 코드를 VDI로 복사하여 사용할 예정입니다.*