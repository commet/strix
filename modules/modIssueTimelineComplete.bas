Attribute VB_Name = "modIssueTimelineComplete"
' Complete Issue Timeline Module with April-October Range
Option Explicit

' API 기본 URL
Private Const API_BASE_URL As String = "http://localhost:5001/api"

' 이슈 목록 가져오기
Function GetCompleteIssueList() As Collection
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
        
        pos = 1
        
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
                issue("first_mentioned_date") = ExtractValue(issueStr, "first_mentioned_date")
                issue("last_updated") = ExtractValue(issueStr, "last_updated")
                
                issues.Add issue
            End If
            
            pos = pos + 1
        Loop
    End If
    
    Set GetCompleteIssueList = issues
    Exit Function
    
ErrorHandler:
    Set GetCompleteIssueList = issues
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

' 초기화 (데이터 지우기)
Sub InitializeTimeline()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    Application.ScreenUpdating = False
    
    ' 데이터 영역만 지우기 (헤더는 유지)
    ws.Range("B9:M60").ClearContents
    ws.Range("B9:M60").Interior.Pattern = xlNone
    ws.Range("B9:M60").Borders.LineStyle = xlNone
    ws.Range("B9:M60").Font.Color = RGB(0, 0, 0)
    
    Application.ScreenUpdating = True
    
    MsgBox "타임라인이 초기화되었습니다." & vbCrLf & vbCrLf & _
           "'새로고침' 버튼을 눌러 데이터를 불러오세요.", _
           vbInformation, "Issue Timeline 초기화"
End Sub

' 완성된 타임라인 업데이트
Sub UpdateCompleteTimeline()
    Dim ws As Worksheet
    Dim issues As Collection
    Dim issue As Object
    Dim row As Integer
    Dim col As Integer
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    Application.StatusBar = "이슈 데이터를 가져오는 중..."
    Application.ScreenUpdating = False
    
    ' API에서 이슈 목록 가져오기
    Set issues = GetCompleteIssueList()
    
    ' 기존 데이터 지우기
    ws.Range("B9:M60").ClearContents
    ws.Range("B9:M60").Interior.Pattern = xlNone
    ws.Range("B9:M60").Borders.LineStyle = xlNone
    ws.Range("B9:M60").Font.Color = RGB(0, 0, 0)
    ws.Range("B9:M60").Font.Bold = False
    
    ' 헤더 재설정 (2025년 4월-10월)
    ws.Range("B8").Value = "최초 언급"
    ws.Range("C8").Value = "이슈 제목"
    ws.Range("D8").Value = "카테고리"
    ws.Range("E8").Value = "상태"
    ws.Range("F8").Value = "담당부서"
    ws.Range("G8").Value = "2025-04"
    ws.Range("H8").Value = "2025-05"
    ws.Range("I8").Value = "2025-06"
    ws.Range("J8").Value = "2025-07"
    ws.Range("K8").Value = "2025-08"
    ws.Range("L8").Value = "2025-09"
    ws.Range("M8").Value = "2025-10"
    ws.Range("N8").Value = "관련문서"
    
    ' 헤더 서식
    With ws.Range("B8:N8")
        .Interior.Color = RGB(52, 73, 94)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 열 너비 조정
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 35
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G:M").ColumnWidth = 10
    ws.Columns("N").ColumnWidth = 12
    
    If issues.Count = 0 Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        MsgBox "이슈 데이터가 없습니다.", vbExclamation
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
            .Cells(row, 3).Font.Size = 10
            
            ' 카테고리
            .Cells(row, 4).Value = issue("category")
            .Cells(row, 4).HorizontalAlignment = xlCenter
            .Cells(row, 4).Font.Size = 9
            
            ' 상태 (한글로 변환 및 색상)
            Dim statusText As String
            Dim statusColor As Long
            
            Select Case issue("status")
                Case "OPEN"
                    statusText = "미해결"
                    statusColor = RGB(255, 0, 0)
                Case "IN_PROGRESS"
                    statusText = "진행중"
                    statusColor = RGB(255, 165, 0)
                Case "RESOLVED"
                    statusText = "해결됨"
                    statusColor = RGB(0, 128, 0)
                Case "MONITORING"
                    statusText = "모니터링"
                    statusColor = RGB(0, 0, 255)
                Case Else
                    statusText = issue("status")
                    statusColor = RGB(100, 100, 100)
            End Select
            
            .Cells(row, 5).Value = statusText
            .Cells(row, 5).Font.Color = statusColor
            .Cells(row, 5).Font.Bold = True
            .Cells(row, 5).HorizontalAlignment = xlCenter
            
            ' 부서
            .Cells(row, 6).Value = issue("department")
            .Cells(row, 6).HorizontalAlignment = xlCenter
            .Cells(row, 6).Font.Size = 9
            
            ' 타임라인 그리기
            Call DrawCompleteTimeline(ws, row, issue)
            
            ' 관련문서 링크
            .Cells(row, 14).Value = "문서 보기"
            .Cells(row, 14).Font.Color = RGB(0, 102, 204)
            .Cells(row, 14).Font.Underline = xlUnderlineStyleSingle
            .Cells(row, 14).HorizontalAlignment = xlCenter
            .Cells(row, 14).Font.Size = 9
            
            ' 행 서식
            With .Range(.Cells(row, 2), .Cells(row, 14))
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200)
                .Borders.Weight = xlThin
                .RowHeight = 28
            End With
            
            ' 우선순위 강조
            If issue("priority") = "CRITICAL" Then
                .Cells(row, 3).Font.Bold = True
                .Cells(row, 3).Font.Color = RGB(200, 0, 0)
            ElseIf issue("priority") = "HIGH" Then
                .Cells(row, 3).Font.Bold = True
            End If
        End With
        
        row = row + 1
        If row > 35 Then Exit For  ' 최대 26개 이슈 표시
    Next issue
    
    ' 전체 서식
    ws.Range("B9:N" & row - 1).Font.Name = "맑은 고딕"
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "타임라인이 업데이트되었습니다!" & vbCrLf & vbCrLf & _
           "총 " & issues.Count & "개 이슈 로드 완료" & vbCrLf & _
           "표시: " & IIf(row - 9 < issues.Count, row - 9, issues.Count) & "개", _
           vbInformation, "Issue Timeline"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "업데이트 중 오류 발생: " & Err.Description, vbCritical
End Sub

' 타임라인 바 그리기
Private Sub DrawCompleteTimeline(ws As Worksheet, row As Integer, issue As Object)
    Dim startCol As Integer, endCol As Integer
    Dim statusColor As Long
    Dim issueDate As Date
    Dim monthDiff As Integer
    Dim col As Integer
    
    ' 상태별 색상
    Select Case issue("status")
        Case "OPEN"
            statusColor = RGB(255, 0, 0)
        Case "IN_PROGRESS"
            statusColor = RGB(255, 165, 0)
        Case "RESOLVED"
            statusColor = RGB(0, 128, 0)
        Case "MONITORING"
            statusColor = RGB(0, 0, 255)
        Case Else
            statusColor = RGB(200, 200, 200)
    End Select
    
    ' 날짜 파싱
    On Error Resume Next
    If issue("first_mentioned_date") <> "" And issue("first_mentioned_date") <> "null" Then
        issueDate = CDate(Left(issue("first_mentioned_date"), 10))
    Else
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 2025년 4월 1일 기준
    Dim baseDate As Date
    baseDate = DateSerial(2025, 4, 1)
    
    ' 월 차이 계산
    monthDiff = DateDiff("m", baseDate, issueDate)
    
    ' 시작 열 (G열=7이 4월)
    startCol = 7 + monthDiff
    If startCol < 7 Then startCol = 7
    If startCol > 13 Then startCol = 13
    
    ' 종료 열 계산
    Select Case issue("status")
        Case "OPEN"
            endCol = 13  ' 10월까지
        Case "IN_PROGRESS", "MONITORING"
            endCol = 11  ' 8월까지 (현재)
        Case "RESOLVED"
            ' last_updated 기준으로 종료
            Dim endDate As Date
            On Error Resume Next
            If issue("last_updated") <> "" Then
                endDate = CDate(Left(issue("last_updated"), 10))
                Dim endDiff As Integer
                endDiff = DateDiff("m", baseDate, endDate)
                endCol = 7 + endDiff
            Else
                endCol = startCol + 1
            End If
            On Error GoTo 0
            If endCol > 13 Then endCol = 13
            If endCol < startCol Then endCol = startCol + 1
    End Select
    
    ' 타임라인 바 그리기
    For col = startCol To endCol
        If col >= 7 And col <= 13 Then
            With ws.Cells(row, col)
                .Interior.Color = statusColor
                .Interior.Pattern = xlSolid
                
                ' 마커 추가 (안전한 문자 사용)
                If col = startCol Then
                    ' 시작점: ● (검은 원)
                    .Value = ChrW(9679)  ' ● 유니코드
                    .Font.Name = "맑은 고딕"
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Bold = True
                    .Font.Size = 12
                    .HorizontalAlignment = xlCenter
                ElseIf col = endCol And issue("status") = "RESOLVED" Then
                    ' 완료: ✓ (체크)
                    .Value = ChrW(10003)  ' ✓ 유니코드
                    .Font.Name = "맑은 고딕"
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Bold = True
                    .Font.Size = 12
                    .HorizontalAlignment = xlCenter
                ElseIf col = 11 And (issue("status") = "IN_PROGRESS" Or issue("status") = "MONITORING") Then
                    ' 진행중: ▶ (삼각형)
                    .Value = ChrW(9654)  ' ▶ 유니코드
                    .Font.Name = "맑은 고딕"
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Bold = True
                    .Font.Size = 11
                    .HorizontalAlignment = xlCenter
                End If
            End With
        End If
    Next col
End Sub

' RefreshIssueTimeline에서 호출
Sub RefreshCompleteTimeline()
    Call UpdateCompleteTimeline
End Sub