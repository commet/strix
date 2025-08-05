Attribute VB_Name = "modMockDataGenerator"
Option Explicit

'==============================================================================
' Mock Data Generator Module
' Word, PPT, PDF 형태의 실제와 유사한 Mock 데이터 생성
'==============================================================================

Public Sub GenerateRealisticMockData()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim basePath As String
    basePath = ThisWorkbook.Path & "\mock_data\"
    
    ' 폴더 구조 생성
    CreateFolderStructure fso, basePath
    
    ' Mock 내부 문서 생성 (Word, PPT, PDF 시뮬레이션)
    GenerateMockInternalDocuments fso, basePath & "internal\"
    
    ' Mock 외부 뉴스 생성
    GenerateMockExternalNews fso, basePath & "external\"
    
    ' Config 업데이트
    UpdateConfigPaths basePath
    
    MsgBox "실제 형태의 Mock 데이터 생성 완료!" & vbNewLine & _
           "내부문서: " & basePath & "internal\" & vbNewLine & _
           "외부뉴스: " & basePath & "external\", vbInformation
    
    Exit Sub
ErrorHandler:
    MsgBox "Mock 데이터 생성 중 오류: " & Err.Description, vbCritical
End Sub

Private Sub CreateFolderStructure(fso As Object, basePath As String)
    ' 기본 폴더
    If Not fso.FolderExists(basePath) Then fso.CreateFolder basePath
    If Not fso.FolderExists(basePath & "internal\") Then fso.CreateFolder basePath & "internal\"
    If Not fso.FolderExists(basePath & "external\") Then fso.CreateFolder basePath & "external\"
    
    ' 내부 문서 하위 폴더 (조직별)
    Dim orgs As Variant
    orgs = Array("전략기획", "R&D", "경영지원", "생산", "영업마케팅")
    
    Dim i As Long
    For i = 0 To UBound(orgs)
        If Not fso.FolderExists(basePath & "internal\" & orgs(i) & "\") Then
            fso.CreateFolder basePath & "internal\" & orgs(i) & "\"
        End If
    Next i
    
    ' 외부 뉴스 하위 폴더 (출처별)
    Dim sources As Variant
    sources = Array("PR팀_AM", "PR팀_PM", "Google_Alert", "Naver_News")
    
    For i = 0 To UBound(sources)
        If Not fso.FolderExists(basePath & "external\" & sources(i) & "\") Then
            fso.CreateFolder basePath & "external\" & sources(i) & "\"
        End If
    Next i
End Sub

Private Sub GenerateMockInternalDocuments(fso As Object, internalPath As String)
    ' 실제 회사 보고서 형태의 Mock 파일 생성
    ' 참고: VBA에서는 실제 Word/PPT/PDF를 생성할 수 없으므로,
    ' 확장자만 해당 형태로 하고 내용은 텍스트로 시뮬레이션
    
    Dim docData As Variant
    Dim i As Long
    
    ' [폴더, 파일명, 문서타입, 내용키워드]
    docData = Array( _
        Array("전략기획", "2024_Q1_배터리사업_중장기전략.docx", "Word", "배터리,전략,투자,성장"), _
        Array("전략기획", "2024_01_월간경영회의_보고서.pptx", "PPT", "경영,실적,이슈,대응"), _
        Array("R&D", "2024_01_전고체배터리_개발현황.pptx", "PPT", "전고체,배터리,기술,개발"), _
        Array("R&D", "2024_차세대배터리_기술로드맵.pdf", "PDF", "기술,로드맵,혁신,특허"), _
        Array("경영지원", "2024_Q1_리스크관리_보고서.docx", "Word", "리스크,규제,대응,관리"), _
        Array("경영지원", "2024_ESG경영_추진현황.pdf", "PDF", "ESG,지속가능,환경,안전"), _
        Array("생산", "2024_01_스마트팩토리_구축계획.pptx", "PPT", "스마트팩토리,자동화,효율,품질"), _
        Array("생산", "2024_Q1_생산성향상_프로젝트.docx", "Word", "생산성,개선,원가,절감"), _
        Array("영업마케팅", "2024_글로벌시장_진출전략.pptx", "PPT", "글로벌,시장,진출,전략"), _
        Array("영업마케팅", "2024_Q1_고객사_대응현황.pdf", "PDF", "고객,대응,수주,영업") _
    )
    
    For i = 0 To UBound(docData)
        CreateMockDocument fso, internalPath, docData(i)(0), docData(i)(1), docData(i)(2), docData(i)(3)
    Next i
End Sub

Private Sub CreateMockDocument(fso As Object, basePath As String, org As String, fileName As String, docType As String, keywords As String)
    Dim filePath As String
    filePath = basePath & org & "\" & fileName
    
    Dim ts As Object
    Set ts = fso.CreateTextFile(filePath, True)
    
    ' 문서 메타데이터 시뮬레이션
    ts.WriteLine "=== MOCK DOCUMENT METADATA ==="
    ts.WriteLine "문서명: " & fileName
    ts.WriteLine "조직: " & org
    ts.WriteLine "문서타입: " & docType
    ts.WriteLine "생성일: " & Format(DateAdd("d", -Int(Rnd * 30), Date), "yyyy-mm-dd")
    ts.WriteLine "작성자: " & Choose(Int(Rnd * 5) + 1, "김전략", "이연구", "박경영", "최생산", "정영업")
    ts.WriteLine ""
    ts.WriteLine "=== 주요 내용 ==="
    
    ' 문서 타입별 내용 시뮬레이션
    Select Case docType
        Case "Word"
            ts.WriteLine "1. 개요"
            ts.WriteLine "   - " & org & " 부서의 주요 현안 및 추진 과제"
            ts.WriteLine "2. 주요 성과"
            ts.WriteLine "   - 키워드: " & keywords
            ts.WriteLine "3. 향후 계획"
            ts.WriteLine "   - 지속적인 개선 및 혁신 추진"
            
        Case "PPT"
            ts.WriteLine "Slide 1: 제목"
            ts.WriteLine "Slide 2: 목차"
            ts.WriteLine "Slide 3: 현황 분석"
            ts.WriteLine "   - 키워드: " & keywords
            ts.WriteLine "Slide 4: 주요 이슈"
            ts.WriteLine "Slide 5: 대응 방안"
            ts.WriteLine "Slide 6: 향후 계획"
            
        Case "PDF"
            ts.WriteLine "== 보고서 요약 =="
            ts.WriteLine "주제: " & Left(fileName, InStr(fileName, ".") - 1)
            ts.WriteLine "핵심 키워드: " & keywords
            ts.WriteLine "주요 내용:"
            ts.WriteLine "- 현재 상황 분석"
            ts.WriteLine "- 개선 방향 제시"
            ts.WriteLine "- 실행 계획 수립"
    End Select
    
    ts.WriteLine ""
    ts.WriteLine "=== 경영진 관심사항 ==="
    If InStr(keywords, "전략") > 0 Or InStr(keywords, "투자") > 0 Then
        ts.WriteLine "[경영진 관심사항 표시]"
    End If
    
    If InStr(keywords, "리스크") > 0 Or InStr(keywords, "규제") > 0 Then
        ts.WriteLine "[리스크 관리 필요]"
    End If
    
    ts.Close
End Sub

Private Sub GenerateMockExternalNews(fso As Object, externalPath As String)
    Dim newsData As Variant
    Dim i As Long
    
    ' [출처, 파일명, 제목, 카테고리, 내용]
    newsData = Array( _
        Array("PR팀_AM", "2024-01-15_AM_뉴스브리핑.txt", "[Macro] 美 금리 동결 전망", "Macro", "연준 금리 동결 가능성 증대, 시장 안정화 기대"), _
        Array("PR팀_AM", "2024-01-16_AM_뉴스브리핑.txt", "[산업] 글로벌 배터리 수요 급증", "산업", "전기차 시장 성장으로 배터리 수요 전년 대비 40% 증가"), _
        Array("PR팀_PM", "2024-01-15_PM_뉴스브리핑.txt", "[기술] 전고체 배터리 상용화 임박", "기술", "주요 업체 2025년 양산 계획 발표"), _
        Array("PR팀_PM", "2024-01-16_PM_뉴스브리핑.txt", "[리스크] EU 배터리 규제 강화", "리스크", "탄소발자국 공시 의무화, 2024년 7월 시행"), _
        Array("Google_Alert", "2024-01-15_CATL_신공장건설.txt", "[경쟁사] CATL 유럽 신공장 착공", "경쟁사", "CATL 헝가리 공장 착공, 연산 100GWh 규모"), _
        Array("Google_Alert", "2024-01-16_Tesla_배터리전략.txt", "[경쟁사] Tesla 자체 배터리 생산 확대", "경쟁사", "Tesla 4680 배터리 생산량 두 배 증설 계획"), _
        Array("Naver_News", "2024-01-15_정부정책_발표.txt", "[정책] 정부 K-배터리 지원책 발표", "정책", "배터리 R&D 예산 1조원 투입, 인력양성 확대"), _
        Array("Naver_News", "2024-01-16_산업동향_분석.txt", "[산업] 국내 배터리 3사 점유율 상승", "산업", "글로벌 시장 점유율 30% 돌파, 중국 업체와 격차 축소") _
    )
    
    For i = 0 To UBound(newsData)
        CreateMockNews fso, externalPath, newsData(i)(0), newsData(i)(1), newsData(i)(2), newsData(i)(3), newsData(i)(4)
    Next i
End Sub

Private Sub CreateMockNews(fso As Object, basePath As String, source As String, fileName As String, title As String, category As String, content As String)
    Dim filePath As String
    filePath = basePath & source & "\" & fileName
    
    Dim ts As Object
    Set ts = fso.CreateTextFile(filePath, True)
    
    ' 뉴스 메일 형식 시뮬레이션
    ts.WriteLine "From: " & IIf(InStr(source, "PR팀") > 0, "pr.team@company.com", source & "@alert.com")
    ts.WriteLine "Date: " & Left(fileName, 10)
    ts.WriteLine "Subject: " & title
    ts.WriteLine "Category: " & category
    ts.WriteLine ""
    ts.WriteLine "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    ts.WriteLine title
    ts.WriteLine "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    ts.WriteLine ""
    ts.WriteLine content
    ts.WriteLine ""
    
    ' 상세 내용 추가
    Select Case category
        Case "Macro"
            ts.WriteLine "▶ 주요 경제 지표"
            ts.WriteLine "  - GDP 성장률: 2.5%"
            ts.WriteLine "  - 물가상승률: 3.2%"
            ts.WriteLine "  - 환율: 1,300원/달러"
            
        Case "산업"
            ts.WriteLine "▶ 산업 동향"
            ts.WriteLine "  - 배터리 시장 규모: 500조원"
            ts.WriteLine "  - 전기차 판매: 전년 대비 35% 증가"
            ts.WriteLine "  - 신규 투자: 100조원 규모"
            
        Case "기술"
            ts.WriteLine "▶ 기술 혁신"
            ts.WriteLine "  - 에너지밀도: 400Wh/kg 달성"
            ts.WriteLine "  - 충전시간: 10분 내 80%"
            ts.WriteLine "  - 수명: 100만km 보장"
            
        Case "리스크"
            ts.WriteLine "▶ 리스크 요인"
            ts.WriteLine "  - 규제 강화"
            ts.WriteLine "  - 원자재 가격 상승"
            ts.WriteLine "  - 공급망 불안정"
            
        Case "경쟁사"
            ts.WriteLine "▶ 경쟁사 동향"
            ts.WriteLine "  - 시장 점유율 변화"
            ts.WriteLine "  - 신규 투자 계획"
            ts.WriteLine "  - 기술 개발 현황"
            
        Case "정책"
            ts.WriteLine "▶ 정책 영향"
            ts.WriteLine "  - 지원금 규모"
            ts.WriteLine "  - 규제 변화"
            ts.WriteLine "  - 산업 육성 방안"
    End Select
    
    ts.WriteLine ""
    ts.WriteLine "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    ts.WriteLine "[출처: " & source & "]"
    
    ts.Close
End Sub

Private Sub UpdateConfigPaths(basePath As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    ws.Range("B2").Value = basePath & "internal\"
    ws.Range("B3").Value = basePath & "external\"
    ws.Range("B4").Value = Now
    ws.Range("B5").Value = Now
End Sub