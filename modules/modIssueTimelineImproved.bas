Attribute VB_Name = "modIssueTimelineImproved"
' Improved Issue Timeline with Keyword Search
Option Explicit

Private highlightKeyword As String

' Issue Timeline Dashboard 생성 (개선 버전)
Sub CreateImprovedIssueTimeline()
    Dim ws As Worksheet
    Dim timelineWs As Worksheet
    
    ' 기존 시트 삭제
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Issue Timeline").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' 새 시트 생성
    Set timelineWs = ThisWorkbook.Sheets.Add
    timelineWs.Name = "Issue Timeline"
    timelineWs.Activate
    
    ' 전체 배경색
    timelineWs.Cells.Interior.Color = RGB(245, 245, 245)
    
    ' 열 너비 설정
    timelineWs.Columns("A").ColumnWidth = 2
    timelineWs.Columns("B").ColumnWidth = 15  ' 날짜
    timelineWs.Columns("C").ColumnWidth = 40  ' 이슈 제목
    timelineWs.Columns("D").ColumnWidth = 12  ' 카테고리
    timelineWs.Columns("E").ColumnWidth = 12  ' 상태
    timelineWs.Columns("F").ColumnWidth = 15  ' 부서
    timelineWs.Columns("G:K").ColumnWidth = 20 ' 타임라인
    timelineWs.Columns("L").ColumnWidth = 2
    
    ' 헤더 영역
    With timelineWs.Range("B2:K2")
        .Merge
        .Value = "STRIX Issue Timeline & Decision Tracker"
        .Font.Name = "맑은 고딕"
        .Font.Size = 24
        .Font.Bold = True
        .Interior.Color = RGB(41, 128, 185)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 50
    End With
    
    ' 부제목
    With timelineWs.Range("B3:K3")
        .Merge
        .Value = "사내 이슈 진행 현황 및 의사결정 추적 시스템"
        .Font.Size = 14
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' ===== 키워드 검색 영역 추가 =====
    timelineWs.Range("B5").Value = "키워드:"
    timelineWs.Range("B5").Font.Bold = True
    timelineWs.Range("B5").Font.Size = 12
    timelineWs.Range("B5").Font.Color = RGB(255, 0, 0)
    
    With timelineWs.Range("C5")
        .Name = "KeywordInput"
        .Interior.Color = RGB(255, 255, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 0, 0)
        .Font.Size = 12
        .Value = ""
        .Font.Color = RGB(0, 0, 0)
    End With
    
    ' 키워드 검색 버튼
    Dim keywordBtn As Object
    Set keywordBtn = timelineWs.Buttons.Add(timelineWs.Range("D5").Left, _
                                           timelineWs.Range("D5").Top, 60, 25)
    With keywordBtn
        .Caption = "검색"
        .OnAction = "HighlightByKeyword"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' 키워드 초기화 버튼
    Set keywordBtn = timelineWs.Buttons.Add(timelineWs.Range("D5").Left + 65, _
                                           timelineWs.Range("D5").Top, 60, 25)
    With keywordBtn
        .Caption = "해제"
        .OnAction = "ClearKeywordHighlight"
        .Font.Size = 11
    End With
    
    ' ===== 필터 영역 (더 큰 글자) =====
    timelineWs.Range("B6").Value = "필터:"
    timelineWs.Range("B6").Font.Bold = True
    timelineWs.Range("B6").Font.Size = 12
    
    ' 카테고리 필터 (Data Validation 사용)
    With timelineWs.Range("C6")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            Formula1:="전체,전략,기술,리스크,경쟁사,정책"
        .Value = "전체"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 14  ' 글자 크기 증가
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' 상태 필터 (Data Validation 사용)
    With timelineWs.Range("E6")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            Formula1:="전체,미해결,진행중,해결됨,모니터링"
        .Value = "전체"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 14  ' 글자 크기 증가
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' 기간 필터
    With timelineWs.Range("G6")
        .Value = "최근 3개월"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .Font.Size = 12
    End With
    
    ' 새로고침 버튼
    Dim refreshBtn As Object
    Set refreshBtn = timelineWs.Buttons.Add(timelineWs.Range("H6").Left, _
                                           timelineWs.Range("H6").Top, 80, 30)
    With refreshBtn
        .Caption = "새로고침"
        .OnAction = "RefreshImprovedTimeline"
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    ' AI 분석 버튼
    Dim aiBtn As Object
    Set aiBtn = timelineWs.Buttons.Add(timelineWs.Range("I6").Left, _
                                       timelineWs.Range("I6").Top, 80, 30)
    With aiBtn
        .Caption = "AI 분석"
        .OnAction = "RunImprovedAIAnalysis"
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    ' 타임라인 영역 헤더
    With timelineWs.Range("B9:K9")
        .Interior.Color = RGB(52, 73, 94)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Font.Size = 12
    End With
    
    timelineWs.Range("B9").Value = "최초 언급"
    timelineWs.Range("C9").Value = "이슈 제목"
    timelineWs.Range("D9").Value = "카테고리"
    timelineWs.Range("E9").Value = "상태"
    timelineWs.Range("F9").Value = "담당부서"
    
    ' 타임라인 월 헤더
    Dim currentMonth As Date
    Dim col As Integer
    currentMonth = DateSerial(Year(Date), Month(Date) - 2, 1)
    
    For col = 7 To 11
        timelineWs.Cells(9, col).Value = Format(currentMonth, "yyyy-MM")
        currentMonth = DateAdd("m", 1, currentMonth)
    Next col
    
    ' 샘플 이슈 추가
    Call AddImprovedIssues(timelineWs)
    
    ' 화면 설정
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    timelineWs.Range("B9").Select
    
    MsgBox "Improved Issue Timeline이 생성되었습니다!" & Chr(10) & Chr(10) & _
           "새로운 기능:" & Chr(10) & _
           "- 키워드 검색으로 관련 이슈 강조" & Chr(10) & _
           "- 더 큰 필터 드롭다운" & Chr(10) & _
           "- 개선된 새로고침 및 AI 분석", _
           vbInformation, "STRIX Issue Tracker"
End Sub

' 향상된 샘플 이슈 추가 (모든 이슈 포함)
Sub AddImprovedIssues(ws As Worksheet)
    Dim row As Integer
    row = 10
    
    ' 이슈 1: SK온 적자 문제
    ws.Cells(row, 2).Value = "2024-01-05"
    ws.Cells(row, 3).Value = "SK온 연속 적자 및 재무구조 개선 필요"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "재무팀"
    Call DrawTimelineBar(ws, row, 7, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 7, "●", "최초 보고")
    Call AddTimelineMarker(ws, row, 9, "▲", "구조조정 착수")
    
    ' 이슈 2: 전고체 배터리 개발
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-15"
    ws.Cells(row, 3).Value = "전고체 배터리 양산 기술 개발 및 파일럿 라인 구축"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 7, 9, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 7, "●", "최초 언급")
    Call AddTimelineMarker(ws, row, 8, "▲", "파일럿 계획")
    
    ' 이슈 3: CATL 시장점유율 확대
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-10"
    ws.Cells(row, 3).Value = "CATL 점유율 37.9% 달성, 대응 전략 수립 필요"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 7, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 7, "●", "이슈 제기")
    
    ' 이슈 4: 미국 IRA 세액공제 축소
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-20"
    ws.Cells(row, 3).Value = "IRA AMPC 세액공제 2401억→3385억원 급감"
    ws.Cells(row, 4).Value = "정책"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "경영지원"
    Call DrawTimelineBar(ws, row, 7, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 7, "●", "정책 변경")
    
    ' ESS 관련 이슈들 추가
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-25"
    ws.Cells(row, 3).Value = "ESS 화재 안전성 강화 방안 수립"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "안전팀"
    Call DrawTimelineBar(ws, row, 7, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 7, "●", "ESS 안전")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-01"
    ws.Cells(row, 3).Value = "2024년 글로벌 생산능력 152GWh로 70% 확대"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "생산관리"
    Call DrawTimelineBar(ws, row, 8, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 8, "●", "확대 결정")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-05"
    ws.Cells(row, 3).Value = "BYD 수직계열화 전략으로 가격경쟁력 강화"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 8, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 8, "●", "경쟁사 분석")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-10"
    ws.Cells(row, 3).Value = "ESS 시장 진출 전략 보고서 작성"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 8, 9, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 9, "☑", "ESS 전략수립")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-15"
    ws.Cells(row, 3).Value = "전기차 수요 둔화로 인한 시장 성장률 하락"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "마케팅"
    Call DrawTimelineBar(ws, row, 8, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 8, "●", "시장분석")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-20"
    ws.Cells(row, 3).Value = "헝가리 제3공장 증설 완료 및 가동 준비"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "해외사업"
    Call DrawTimelineBar(ws, row, 8, 9, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 9, "☑", "증설 완료")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-01"
    ws.Cells(row, 3).Value = "K배터리 3사 글로벌 점유율 18.4%로 4.7%p 하락"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 9, "●", "실적 발표")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-05"
    ws.Cells(row, 3).Value = "46파이 원통형 배터리 파일럿 라인 구축"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 9, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 9, "●", "착공")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-10"
    ws.Cells(row, 3).Value = "CATL ESS 시장 점유율 확대 대응"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "마케팅"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 9, "●", "ESS 경쟁")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-15"
    ws.Cells(row, 3).Value = "중국 내수시장 CATL-BYD 양강구도 고착화"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "중국사업"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 9, "●", "시장 분석")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-20"
    ws.Cells(row, 3).Value = "미국 ESS 세액공제 정책 변화 분석"
    ws.Cells(row, 4).Value = "정책"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "정책팀"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 9, "●", "ESS 정책")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-25"
    ws.Cells(row, 3).Value = "LFP 배터리 원가 절감 프로젝트"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "생산팀"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 9, "●", "LFP 프로젝트")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-01"
    ws.Cells(row, 3).Value = "2024년 4분기 흑자전환 목표 수립 및 추진"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "경영기획"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "목표 수립")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-05"
    ws.Cells(row, 3).Value = "유럽 ESS 규제 대응 체계 구축"
    ws.Cells(row, 4).Value = "정책"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "법무팀"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "ESS 규제")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2025-07-30"
    ws.Cells(row, 3).Value = "SK온-SK엔무브 합병 결정, 11월 1일 통합법인 출범"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "경영기획"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "합병 결의")
    Call AddTimelineMarker(ws, row, 11, "▲", "11월 출범예정")
End Sub

' 타임라인 바 그리기
Private Sub DrawTimelineBar(ws As Worksheet, row As Integer, startCol As Integer, endCol As Integer, barColor As Long)
    Dim cell As Range
    For Each cell In ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol))
        cell.Interior.Color = barColor
        cell.Interior.Pattern = xlSolid
    Next cell
End Sub

' 타임라인 마커 추가
Private Sub AddTimelineMarker(ws As Worksheet, row As Integer, col As Integer, marker As String, tooltip As String)
    With ws.Cells(row, col)
        .Value = marker
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        If .Comment Is Nothing Then
            .AddComment tooltip
        Else
            .Comment.Text tooltip
        End If
        .Comment.Visible = False
    End With
End Sub

' 키워드로 이슈 강조
Sub HighlightByKeyword()
    Dim ws As Worksheet
    Dim keyword As String
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    keyword = ws.Range("C5").Value
    
    If keyword = "" Then
        MsgBox "키워드를 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).row
    
    ' 모든 이슈 제목 검색
    For i = 10 To lastRow
        If ws.Cells(i, 3).Value <> "" Then
            If InStr(UCase(ws.Cells(i, 3).Value), UCase(keyword)) > 0 Then
                ' 키워드가 포함된 경우 배경색 변경
                ws.Cells(i, 3).Interior.Color = RGB(255, 230, 230)  ' 연한 빨간색 배경
                ws.Cells(i, 3).Font.Color = RGB(255, 0, 0)  ' 빨간색 글자
                ws.Cells(i, 3).Font.Bold = True
            Else
                ' 키워드가 없는 경우 원래대로
                ws.Cells(i, 3).Interior.Color = RGB(255, 255, 255)
                ws.Cells(i, 3).Font.Color = RGB(0, 0, 0)
                ws.Cells(i, 3).Font.Bold = False
            End If
        End If
    Next i
    
    MsgBox "'" & keyword & "' 관련 이슈가 강조되었습니다.", vbInformation
End Sub

' 키워드 강조 해제
Sub ClearKeywordHighlight()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).row
    
    ' 모든 이슈 제목 원래대로
    For i = 10 To lastRow
        If ws.Cells(i, 3).Value <> "" Then
            ws.Cells(i, 3).Interior.Color = RGB(255, 255, 255)
            ws.Cells(i, 3).Font.Color = RGB(0, 0, 0)
            ws.Cells(i, 3).Font.Bold = False
        End If
    Next i
    
    ws.Range("C5").Value = ""
    MsgBox "키워드 강조가 해제되었습니다.", vbInformation
End Sub

' 개선된 새로고침
Sub RefreshImprovedTimeline()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    ' 이슈 데이터 유지하면서 새로고침
    Application.ScreenUpdating = False
    
    ' 기존 데이터 지우기
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).row
    If lastRow >= 10 Then
        ws.Range("B10:K" & lastRow).Clear
    End If
    
    ' 샘플 데이터 다시 로드
    Call AddImprovedIssues(ws)
    
    Application.ScreenUpdating = True
    MsgBox "타임라인이 새로고침되었습니다.", vbInformation
End Sub

' 개선된 AI 분석
Sub RunImprovedAIAnalysis()
    On Error GoTo ErrorHandler
    
    ' Smart Alerts 시트로 이동
    If SheetExists("Smart Alerts") Then
        ThisWorkbook.Sheets("Smart Alerts").Activate
        MsgBox "AI 분석 결과가 Smart Alerts 탭에 표시됩니다.", vbInformation
    Else
        MsgBox "Smart Alerts 탭을 먼저 생성해주세요.", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "AI 분석이 Smart Alerts 탭으로 연결됩니다.", vbInformation
End Sub

' 시트 존재 확인
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function