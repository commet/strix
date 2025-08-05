Attribute VB_Name = "modInit"
Option Explicit

'==============================================================================
' STRIX Initialization Module
' 프로젝트 초기화 및 시트 생성
'==============================================================================

' === 메인 초기화 함수 ===
Public Sub InitializeSTRIX()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 1. 시트 구조 생성/검증
    If Not CreateWorksheets() Then
        MsgBox "시트 생성 실패", vbCritical
        GoTo Cleanup
    End If
    
    ' 2. 테이블 구조 생성
    CreateTables
    
    ' 3. Config 시트 초기화
    InitializeConfigSheet
    
    ' 4. 설정 로드
    LoadConfig
    
    ' 5. 리본 메뉴 추가 (선택사항)
    ' AddRibbonMenu
    
    MsgBox "STRIX 초기화 완료!", vbInformation, APP_NAME
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "초기화 중 오류 발생: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' === 워크시트 생성 ===
Private Function CreateWorksheets() As Boolean
    On Error GoTo ErrorHandler
    
    Dim sheetNames As Variant
    Dim i As Long
    
    ' 필수 시트 목록
    sheetNames = Array(SHEET_CONFIG, SHEET_RAWDATA, SHEET_RAWNEWS, _
                      SHEET_METADATA, SHEET_LINKEDNEWS, SHEET_DASHBOARD, _
                      SHEET_GPT, SHEET_NEWSLETTER, SHEET_REPORTS)
    
    ' 각 시트 생성 또는 확인
    For i = 0 To UBound(sheetNames)
        If Not WorksheetExists(CStr(sheetNames(i))) Then
            ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = CStr(sheetNames(i))
        End If
    Next i
    
    CreateWorksheets = True
    Exit Function
    
ErrorHandler:
    CreateWorksheets = False
End Function

' === 테이블 구조 생성 ===
Private Sub CreateTables()
    On Error Resume Next
    
    ' RawData 테이블
    CreateTable SHEET_RAWDATA, "RawData_tbl", _
               Array("FileID", "FileName", "FilePath", "FileType", "FileSize", _
                     "CreatedDate", "ModifiedDate", "UploadDate", "Organization", _
                     "IssueID", "ProcessedFlag")
    
    ' RawNews 테이블
    CreateTable SHEET_RAWNEWS, "RawNews_tbl", _
               Array("MailID", "ReceivedDate", "Subject", "Sender", "BodyText", _
                     "AttachmentPath", "Category", "SubCategory", "ProcessedFlag")
    
    ' MetaData 테이블
    CreateTable SHEET_METADATA, "MetaData_tbl", _
               Array("IssueID", "IssueName", "Organization", "Keywords", _
                     "Priority", "Status", "SuccessCase", "ExecInterest", _
                     "FirstReported", "LastUpdated", "Description")
    
    ' LinkedNews 테이블
    CreateTable SHEET_LINKEDNEWS, "LinkedNews_tbl", _
               Array("LinkID", "IssueID", "MailID", "CorrelationScore", _
                     "VerifiedFlag", "VerifiedBy", "VerifiedDate", "Notes")
    
    ' GPT_Interface 테이블
    CreateTable SHEET_GPT, "GPT_tbl", _
               Array("PromptID", "PromptDate", "PromptText", "ResponseText", _
                     "UsedBy", "Purpose")
    
    ' Reports 테이블
    CreateTable SHEET_REPORTS, "Reports_tbl", _
               Array("ReportID", "ReportType", "GeneratedDate", "GeneratedBy", _
                     "FilePath", "Recipients", "Status")
End Sub

' === Config 시트 초기화 ===
Private Sub InitializeConfigSheet()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    ' 헤더 설정
    ws.Range("A1").Value = "STRIX Configuration"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    
    ' 설정 항목
    ws.Range("A2:A10").Value = Application.Transpose(Array( _
        "내부문서 폴더", "외부뉴스 폴더", "마지막 내부스캔", _
        "마지막 외부스캔", "현재 사용자", "마지막 업데이트", _
        "스캔 주기(분)", "자동 스캔", "이메일 알림"))
    
    ' 기본값 설정
    If ws.Range("B7").Value = "" Then ws.Range("B7").Value = 60
    If ws.Range("B8").Value = "" Then ws.Range("B8").Value = "Yes"
    If ws.Range("B9").Value = "" Then ws.Range("B9").Value = "No"
    
    ' 잠금 상태 영역
    ws.Range("D1").Value = "시스템 상태"
    ws.Range("D1").Font.Bold = True
    ws.Range("D2").Value = "UNLOCKED"
    
    ' 서식 설정
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 20
    
    ' 키워드 설정 영역
    ws.Range("F1").Value = "카테고리 키워드"
    ws.Range("F1").Font.Bold = True
    
    ws.Range("F2:G10").Value = Application.Transpose(Array( _
        Array("Macro", "경제,금리,환율,인플레이션"), _
        Array("산업", "배터리,전기차,반도체,에너지"), _
        Array("기술", "AI,자동화,디지털,혁신"), _
        Array("리스크", "규제,제재,사고,리콜"), _
        Array("경쟁사", "CATL,BYD,Tesla,Panasonic"), _
        Array("정책", "IRA,CBAM,RE100,탄소중립"), _
        Array("", ""), Array("", ""), Array("", "")))
End Sub

' === 유틸리티 함수 ===
Private Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
End Function

Private Sub CreateTable(sheetName As String, tableName As String, headers As Variant)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    ' 기존 테이블 확인
    Set tbl = ws.ListObjects(tableName)
    If Not tbl Is Nothing Then Exit Sub
    
    ' 헤더 작성
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    
    ' 테이블 생성
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = tableName
    tbl.TableStyle = "TableStyleMedium2"
End Sub

' === 초기 Mock 데이터 생성 ===
Public Sub CreateMockData()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim basePath As String
    basePath = ThisWorkbook.Path & "\mock_data\"
    
    ' Mock 데이터 폴더 생성
    If Not fso.FolderExists(basePath) Then
        fso.CreateFolder basePath
    End If
    
    If Not fso.FolderExists(basePath & "internal\") Then
        fso.CreateFolder basePath & "internal\"
    End If
    
    If Not fso.FolderExists(basePath & "external\") Then
        fso.CreateFolder basePath & "external\"
    End If
    
    ' Mock 내부 문서 생성
    CreateMockInternalFiles basePath & "internal\"
    
    ' Mock 외부 뉴스 생성
    CreateMockExternalFiles basePath & "external\"
    
    ' Config에 경로 설정
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    ws.Range("B2").Value = basePath & "internal\"
    ws.Range("B3").Value = basePath & "external\"
    
    MsgBox "Mock 데이터 생성 완료!" & vbNewLine & _
           "경로: " & basePath, vbInformation
    
    Exit Sub
ErrorHandler:
    MsgBox "Mock 데이터 생성 중 오류: " & Err.Description, vbCritical
End Sub

Private Sub CreateMockInternalFiles(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim fileNames As Variant
    Dim i As Long
    
    fileNames = Array( _
        "2024_Q1_전략기획_배터리사업현황.txt", _
        "2024_Q1_R&D_전고체배터리개발.txt", _
        "2024_Q1_경영지원_리스크관리체계.txt", _
        "2024_Q2_생산_스마트팩토리구축.txt", _
        "2024_Q2_영업_글로벌시장확대.txt")
    
    For i = 0 To UBound(fileNames)
        Dim ts As Object
        Set ts = fso.CreateTextFile(folderPath & fileNames(i), True)
        ts.WriteLine "Mock 내부 문서: " & fileNames(i)
        ts.WriteLine "이슈: 배터리 사업 관련"
        ts.WriteLine "키워드: 배터리, 전기차, 혁신"
        ts.Close
    Next i
End Sub

Private Sub CreateMockExternalFiles(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim newsData As Variant
    Dim i As Long
    
    newsData = Array( _
        Array("2024-01-15_AM_Macro경제동향.txt", "[Macro] 글로벌 경제 전망", "금리 인상 우려 지속..."), _
        Array("2024-01-15_PM_배터리산업뉴스.txt", "[산업] 전고체 배터리 상용화 가속", "주요 업체들 투자 확대..."), _
        Array("2024-01-16_AM_규제동향.txt", "[리스크] EU 배터리 규제 강화", "탄소발자국 공시 의무화..."))
    
    For i = 0 To UBound(newsData)
        Dim ts As Object
        Set ts = fso.CreateTextFile(folderPath & newsData(i)(0), True)
        ts.WriteLine "제목: " & newsData(i)(1)
        ts.WriteLine "날짜: " & Left(newsData(i)(0), 10)
        ts.WriteLine "본문: " & newsData(i)(2)
        ts.Close
    Next i
End Sub