Attribute VB_Name = "modIssueTimelineEnhanced"
' Enhanced Issue Timeline Module with Better Visualization
Option Explicit

' API 기본 URL
Private Const API_BASE_URL As String = "http://localhost:5001/api"

' 색상 상수 정의
Private Const COLOR_OPEN As Long = 15123099       ' RGB(231, 76, 60) - 빨간색
Private Const COLOR_IN_PROGRESS As Long = 1023410  ' RGB(241, 196, 15) - 노란색  
Private Const COLOR_RESOLVED As Long = 7664549     ' RGB(46, 204, 113) - 초록색
Private Const COLOR_MONITORING As Long = 13815039  ' RGB(52, 152, 219) - 파란색

' 향상된 이슈 목록 가져오기
Function GetEnhancedIssueList() As Collection
    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim issues As New Collection
    
    On Error GoTo ErrorHandler
    
    url = API_BASE_URL & "/issues?days=9999"  ' 모든 이슈 가져오기
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send
    
    If http.Status = 200 Then
        responseText = http.responseText
        
        ' 간단한 파싱
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
    
    Set GetEnhancedIssueList = issues
    Exit Function
    
ErrorHandler:
    Set GetEnhancedIssueList = issues
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

' 향상된 타임라인 업데이트
Sub UpdateEnhancedTimeline()
    Dim ws As Worksheet
    Dim issues As Collection
    Dim issue As Object
    Dim row As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    Application.StatusBar = "이슈 목록을 가져오는 중..."
    Application.ScreenUpdating = False
    
    ' API에서 이슈 목록 가져오기
    Set issues = GetEnhancedIssueList()
    
    ' 기존 데이터 지우기
    ws.Range("B9:K60").ClearContents
    ws.Range("B9:K60").Borders.LineStyle = xlNone
    ws.Range("B9:K60").Interior.Pattern = xlNone
    ws.Range("B9:K60").Font.Bold = False
    
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
            ' 날짜 (YYYY-MM-DD 형식)
            If issue("first_mentioned_date") <> "" And issue("first_mentioned_date") <> "null" Then
                .Cells(row, 2).Value = Left(issue("first_mentioned_date"), 10)
                .Cells(row, 2).NumberFormat = "yyyy-mm-dd"
            End If
            
            ' 제목
            .Cells(row, 3).Value = issue("title")
            .Cells(row, 3).WrapText = False
            
            ' 카테고리
            .Cells(row, 4).Value = issue("category")
            .Cells(row, 4).HorizontalAlignment = xlCenter
            
            ' 상태 (한글로 변환 및 색상 적용)
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
            
            ' 부서
            .Cells(row, 6).Value = issue("department")
            .Cells(row, 6).HorizontalAlignment = xlCenter
            
            ' 타임라인 시각화 (월별 진행상황)
            Call DrawEnhancedTimeline(ws, row, issue)
            
            ' 행 서식
            With .Range(.Cells(row, 2), .Cells(row, 11))
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200)
                .Borders.Weight = xlThin
                
                ' 짝수 행 배경색
                If row Mod 2 = 0 Then
                    .Interior.Color = RGB(248, 249, 250)
                Else
                    .Interior.Color = RGB(255, 255, 255)
                End If
                
                .RowHeight = 30
            End With
            
            ' 우선순위에 따른 강조
            If issue("priority") = "CRITICAL" Then
                .Cells(row, 3).Font.Bold = True
                .Range(.Cells(row, 2), .Cells(row, 3)).Interior.Color = RGB(255, 243, 224)
            ElseIf issue("priority") = "HIGH" Then
                .Cells(row, 3).Font.Bold = True
            End If
            
            ' 이슈 ID 저장 (숨김 열)
            .Cells(row, 12).Value = issue("id")
        End With
        
        row = row + 1
        If row > 60 Then Exit For ' 최대 표시 수 제한
    Next issue
    
    ' 전체 서식 조정
    ws.Range("B9:K" & row - 1).Font.Name = "맑은 고딕"
    ws.Range("B9:K" & row - 1).Font.Size = 10
    
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

' 향상된 타임라인 그리기
Private Sub DrawEnhancedTimeline(ws As Worksheet, row As Integer, issue As Object)
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
    
    ' 타임라인 기준 월 계산 (현재월 기준 -2개월부터 +2개월까지)
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
    
    ' 타임라인 시작 열 계산 (7열부터 11열까지)
    startCol = 7 + monthDiff
    If startCol < 7 Then startCol = 7
    If startCol > 11 Then startCol = 11
    
    ' 진행 중인 이슈는 현재까지
    If issue("status") = "IN_PROGRESS" Or issue("status") = "MONITORING" Then
        endCol = 9 ' 현재 월
    ElseIf issue("status") = "RESOLVED" Then
        ' 해결된 이슈는 짧게
        endCol = startCol + 1
        If endCol > 11 Then endCol = 11
    Else
        endCol = 11 ' OPEN 이슈는 미래까지
    End If
    
    If endCol < startCol Then endCol = startCol
    
    ' 타임라인 바 그리기
    For col = startCol To endCol
        With ws.Cells(row, col)
            .Interior.Color = statusColor
            .Interior.Pattern = xlSolid
            
            ' 시작점 마커
            If col = startCol Then
                .Value = "●"
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            ElseIf col = endCol And issue("status") = "RESOLVED" Then
                ' 완료 마커
                .Value = "V"
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            ElseIf col = 9 And (issue("status") = "IN_PROGRESS" Or issue("status") = "MONITORING") Then
                ' 현재 진행 중 마커
                .Value = ">"
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End If
        End With
    Next col
    
    ' 우선순위가 높은 이슈는 추가 강조
    If issue("priority") = "CRITICAL" Or issue("priority") = "HIGH" Then
        For col = startCol To endCol
            ws.Cells(row, col).Borders.LineStyle = xlContinuous
            ws.Cells(row, col).Borders.Weight = xlMedium
            ws.Cells(row, col).Borders.Color = statusColor
        Next col
    End If
End Sub

' RefreshIssueTimeline에서 호출될 함수
Sub RefreshEnhancedTimeline()
    Call UpdateEnhancedTimeline
End Sub