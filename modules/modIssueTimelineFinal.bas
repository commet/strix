Attribute VB_Name = "modIssueTimelineFinal"
' Final Issue Timeline Module with Complete Visualization
Option Explicit

' API 기본 URL
Private Const API_BASE_URL As String = "http://localhost:5001/api"

' 색상 상수 정의 (진한 원색)
Private Const COLOR_OPEN As Long = 255                ' 빨간색 (RGB 255,0,0)
Private Const COLOR_IN_PROGRESS As Long = 42495       ' 오렌지색 (RGB 255,165,0)
Private Const COLOR_RESOLVED As Long = 32768          ' 초록색 (RGB 0,128,0)
Private Const COLOR_MONITORING As Long = 16711680     ' 파란색 (RGB 0,0,255)

' 최종 이슈 목록 가져오기
Function GetFinalIssueList() As Collection
    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim issues As New Collection
    
    On Error GoTo ErrorHandler
    
    url = API_BASE_URL & "/issues?days=9999"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send
    
    If http.Status = 200 Then
        responseText = http.responseText
        
        Dim pos As Long, endPos As Long
        Dim issueStr As String
        Dim issueCount As Integer
        
        pos = 1
        issueCount = 0
        
        Do While InStr(pos, responseText, """id"":") > 0
            pos = InStr(pos, responseText, """id"":")
            If pos = 0 Then Exit Do
            
            endPos = InStr(pos + 1, responseText, """id"":")
            If endPos = 0 Then
                endPos = InStr(pos, responseText, "]")
            End If
            
            If endPos > pos Then
                issueStr = Mid(responseText, pos, endPos - pos)
                
                Dim issue As Object
                Set issue = CreateObject("Scripting.Dictionary")
                
                issue("id") = ExtractValue(issueStr, "id")
                issue("issue_key") = ExtractValue(issueStr, "issue_key")
                issue("title") = ExtractValue(issueStr, "title")
                issue("category") = ExtractValue(issueStr, "category")
                issue("status") = ExtractValue(issueStr, "status")
                issue("priority") = ExtractValue(issueStr, "priority")
                issue("department") = ExtractValue(issueStr, "department")
                issue("owner") = ExtractValue(issueStr, "owner")
                issue("first_mentioned_date") = ExtractValue(issueStr, "first_mentioned_date")
                issue("last_updated") = ExtractValue(issueStr, "last_updated")
                
                issues.Add issue
                issueCount = issueCount + 1
            End If
            
            pos = pos + 1
        Loop
    End If
    
    Set GetFinalIssueList = issues
    Exit Function
    
ErrorHandler:
    Set GetFinalIssueList = issues
End Function

' 값 추출 함수
Private Function ExtractValue(json As String, key As String) As String
    Dim startPos As Long, endPos As Long
    Dim searchKey As String
    
    searchKey = """" & key & """: """
    startPos = InStr(1, json, searchKey)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchKey)
        endPos = InStr(startPos, json, """")
        If endPos > startPos Then
            ExtractValue = Mid(json, startPos, endPos - startPos)
            Exit Function
        End If
    End If
    
    searchKey = """" & key & """: "
    startPos = InStr(1, json, searchKey)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchKey)
        endPos = InStr(startPos, json, ",")
        If endPos = 0 Then endPos = InStr(startPos, json, "}")
        If endPos > startPos Then
            ExtractValue = Trim(Mid(json, startPos, endPos - startPos))
            ExtractValue = Replace(ExtractValue, """", "")
            Exit Function
        End If
    End If
    
    ExtractValue = ""
End Function

' 최종 타임라인 업데이트
Sub UpdateFinalTimeline()
    Dim ws As Worksheet
    Dim issues As Collection
    Dim issue As Object
    Dim row As Integer
    Dim col As Integer
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    Application.StatusBar = "이슈 목록을 가져오는 중..."
    Application.ScreenUpdating = False
    
    ' API에서 이슈 목록 가져오기
    Set issues = GetFinalIssueList()
    
    ' 기존 데이터 지우기
    ws.Range("B8:M60").ClearContents
    ws.Range("B8:M60").Borders.LineStyle = xlNone
    ws.Range("B8:M60").Interior.Pattern = xlNone
    ws.Range("B8:M60").Font.Bold = False
    ws.Range("B8:M60").Font.Color = RGB(0, 0, 0)
    
    ' 헤더 다시 설정 (타임라인 월 표시 포함)
    With ws.Range("B8:M8")
        .Interior.Color = RGB(52, 73, 94)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ws.Range("B8").Value = "최초 언급"
    ws.Range("C8").Value = "이슈 제목"
    ws.Range("D8").Value = "카테고리"
    ws.Range("E8").Value = "상태"
    ws.Range("F8").Value = "담당부서"
    
    ' 타임라인 월 헤더 설정 (동적으로 현재 기준 -2개월부터 +2개월)
    Dim currentMonth As Date
    currentMonth = DateSerial(Year(Date), Month(Date) - 2, 1)
    
    For col = 7 To 11
        ws.Cells(8, col).Value = Format(currentMonth, "yyyy-mm")
        ws.Cells(8, col).Interior.Color = RGB(52, 73, 94)
        ws.Cells(8, col).Font.Color = RGB(255, 255, 255)
        ws.Cells(8, col).Font.Bold = True
        ws.Cells(8, col).HorizontalAlignment = xlCenter
        ws.Cells(8, col).Borders.LineStyle = xlContinuous
        currentMonth = DateAdd("m", 1, currentMonth)
    Next col
    
    ' 관련 문서 열 헤더
    ws.Range("M8").Value = "관련문서"
    ws.Range("M8").Interior.Color = RGB(52, 73, 94)
    ws.Range("M8").Font.Color = RGB(255, 255, 255)
    ws.Range("M8").Font.Bold = True
    ws.Range("M8").HorizontalAlignment = xlCenter
    ws.Range("M8").Borders.LineStyle = xlContinuous
    
    ' 열 너비 조정
    ws.Columns("B").ColumnWidth = 12  ' 날짜
    ws.Columns("C").ColumnWidth = 40  ' 이슈 제목
    ws.Columns("D").ColumnWidth = 10  ' 카테고리
    ws.Columns("E").ColumnWidth = 10  ' 상태
    ws.Columns("F").ColumnWidth = 12  ' 부서
    ws.Columns("G:K").ColumnWidth = 12 ' 타임라인
    ws.Columns("M").ColumnWidth = 15  ' 관련문서
    
    If issues.Count = 0 Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        MsgBox "가져온 이슈가 없습니다.", vbInformation
        Exit Sub
    End If
    
    ' 각 이슈 표시
    row = 9
    For Each issue In issues
        With ws
            ' 날짜
            If issue("first_mentioned_date") <> "" And issue("first_mentioned_date") <> "null" Then
                .Cells(row, 2).Value = Left(issue("first_mentioned_date"), 10)
                .Cells(row, 2).NumberFormat = "yyyy-mm-dd"
                .Cells(row, 2).Font.Size = 9
            End If
            
            ' 제목
            .Cells(row, 3).Value = issue("title")
            .Cells(row, 3).WrapText = False
            .Cells(row, 3).Font.Size = 10
            
            ' 카테고리
            .Cells(row, 4).Value = issue("category")
            .Cells(row, 4).HorizontalAlignment = xlCenter
            .Cells(row, 4).Font.Size = 9
            
            ' 상태 (한글로 변환 및 진한 원색 적용)
            Dim statusText As String
            Dim statusColor As Long
            
            Select Case issue("status")
                Case "OPEN"
                    statusText = "미해결"
                    statusColor = COLOR_OPEN
                Case "IN_PROGRESS"
                    statusText = "진행중"
                    statusColor = COLOR_IN_PROGRESS
                Case "RESOLVED"
                    statusText = "해결됨"
                    statusColor = COLOR_RESOLVED
                Case "MONITORING"
                    statusText = "모니터링"
                    statusColor = COLOR_MONITORING
                Case Else
                    statusText = issue("status")
                    statusColor = RGB(100, 100, 100)
            End Select
            
            .Cells(row, 5).Value = statusText
            .Cells(row, 5).Font.Color = statusColor
            .Cells(row, 5).Font.Bold = True
            .Cells(row, 5).HorizontalAlignment = xlCenter
            .Cells(row, 5).Font.Size = 10
            
            ' 부서
            .Cells(row, 6).Value = issue("department")
            .Cells(row, 6).HorizontalAlignment = xlCenter
            .Cells(row, 6).Font.Size = 9
            
            ' 타임라인 시각화
            Call DrawFinalTimeline(ws, row, issue)
            
            ' 관련 문서 URL (하이퍼링크처럼 보이게)
            Dim docUrl As String
            docUrl = "http://docs.strix.com/issue/" & issue("issue_key")
            .Cells(row, 13).Value = "문서 보기"
            .Cells(row, 13).Font.Color = RGB(0, 102, 204)  ' 파란색 링크
            .Cells(row, 13).Font.Underline = xlUnderlineStyleSingle
            .Cells(row, 13).HorizontalAlignment = xlCenter
            .Cells(row, 13).Font.Size = 9
            
            ' 행 서식 (흰색 배경)
            With .Range(.Cells(row, 2), .Cells(row, 13))
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200)
                .Borders.Weight = xlThin
                .Interior.Color = RGB(255, 255, 255)  ' 모든 행 흰색 배경
                .RowHeight = 28
            End With
            
            ' 우선순위에 따른 제목 강조
            If issue("priority") = "CRITICAL" Then
                .Cells(row, 3).Font.Bold = True
                .Cells(row, 3).Font.Color = RGB(200, 0, 0)
            ElseIf issue("priority") = "HIGH" Then
                .Cells(row, 3).Font.Bold = True
            End If
            
            ' 이슈 ID 저장 (숨김 열)
            .Cells(row, 14).Value = issue("id")
        End With
        
        row = row + 1
        If row > 60 Then Exit For
    Next issue
    
    ' 전체 서식 조정
    ws.Range("B9:M" & row - 1).Font.Name = "맑은 고딕"
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "타임라인이 업데이트되었습니다!" & Chr(10) & Chr(10) & _
           "총 " & issues.Count & "개 이슈" & Chr(10) & _
           "표시: " & IIf(row - 9 < issues.Count, row - 9, issues.Count) & "개", _
           vbInformation, "Issue Timeline"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "업데이트 중 오류 발생: " & Err.Description, vbCritical
End Sub

' 최종 타임라인 그리기
Private Sub DrawFinalTimeline(ws As Worksheet, row As Integer, issue As Object)
    Dim startCol As Integer, endCol As Integer
    Dim statusColor As Long
    Dim currentMonth As Date
    Dim issueDate As Date
    Dim monthDiff As Integer
    Dim col As Integer
    
    ' 상태별 색상 설정
    Select Case issue("status")
        Case "OPEN"
            statusColor = COLOR_OPEN
        Case "IN_PROGRESS"
            statusColor = COLOR_IN_PROGRESS
        Case "RESOLVED"
            statusColor = COLOR_RESOLVED
        Case "MONITORING"
            statusColor = COLOR_MONITORING
        Case Else
            statusColor = RGB(200, 200, 200)
    End Select
    
    ' 타임라인 기준 월 계산
    currentMonth = DateSerial(Year(Date), Month(Date) - 2, 1)
    
    ' 이슈 날짜 파싱
    On Error Resume Next
    If issue("first_mentioned_date") <> "" And issue("first_mentioned_date") <> "null" Then
        issueDate = CDate(Left(issue("first_mentioned_date"), 10))
    Else
        issueDate = Date
    End If
    On Error GoTo 0
    
    ' 월 차이 계산
    monthDiff = DateDiff("m", currentMonth, issueDate)
    
    ' 타임라인 시작 열 계산
    startCol = 7 + monthDiff
    If startCol < 7 Then startCol = 7
    If startCol > 11 Then startCol = 11
    
    ' 종료 열 계산
    Select Case issue("status")
        Case "IN_PROGRESS", "MONITORING"
            endCol = 9  ' 현재 월
        Case "RESOLVED"
            ' 해결된 이슈는 last_updated 기준
            Dim resolvedDate As Date
            On Error Resume Next
            If issue("last_updated") <> "" And issue("last_updated") <> "null" Then
                resolvedDate = CDate(Left(issue("last_updated"), 10))
                Dim resolveDiff As Integer
                resolveDiff = DateDiff("m", currentMonth, resolvedDate)
                endCol = 7 + resolveDiff
                If endCol < startCol Then endCol = startCol + 1
                If endCol > 11 Then endCol = 11
            Else
                endCol = startCol + 1
            End If
            On Error GoTo 0
        Case "OPEN"
            endCol = 11  ' 미래까지
        Case Else
            endCol = startCol
    End Select
    
    If endCol < startCol Then endCol = startCol
    
    ' 타임라인 바 그리기
    For col = startCol To endCol
        With ws.Cells(row, col)
            .Interior.Color = statusColor
            .Interior.Pattern = xlSolid
            
            ' 마커 추가
            If col = startCol Then
                .Value = Chr(149)  ' 시작 점
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .Font.Size = 12
                .HorizontalAlignment = xlCenter
            ElseIf col = endCol And issue("status") = "RESOLVED" Then
                .Value = Chr(252)  ' 체크 마크
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .Font.Size = 10
                .HorizontalAlignment = xlCenter
            ElseIf col = 9 And (issue("status") = "IN_PROGRESS" Or issue("status") = "MONITORING") Then
                .Value = Chr(187)  ' 화살표
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .Font.Size = 10
                .HorizontalAlignment = xlCenter
            End If
            
            ' 테두리 강조
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = statusColor
        End With
    Next col
End Sub

' RefreshIssueTimeline에서 호출될 함수
Sub RefreshFinalTimeline()
    Call UpdateFinalTimeline
End Sub