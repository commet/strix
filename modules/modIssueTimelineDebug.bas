Attribute VB_Name = "modIssueTimelineDebug"
' Debug Version - Issue Timeline Module
Option Explicit

' API 기본 URL
Private Const API_BASE_URL As String = "http://localhost:5001/api"

' 이슈 목록 가져오기 (디버그)
Function GetDebugIssueList() As Collection
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
    
    Set GetDebugIssueList = issues
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting issues: " & Err.Description
    Set GetDebugIssueList = issues
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

' 디버그 타임라인 업데이트
Sub UpdateDebugTimeline()
    Dim ws As Worksheet
    Dim issues As Collection
    Dim issue As Object
    Dim row As Integer
    Dim col As Integer
    Dim debugMsg As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    Application.StatusBar = "이슈 목록을 가져오는 중..."
    Application.ScreenUpdating = False
    
    ' API에서 이슈 목록 가져오기
    Set issues = GetDebugIssueList()
    
    MsgBox "가져온 이슈 수: " & issues.Count, vbInformation
    
    ' 기존 데이터 지우기
    ws.Range("B9:M60").ClearContents
    ws.Range("B9:M60").Interior.Pattern = xlNone
    ws.Range("B9:M60").Font.Color = RGB(0, 0, 0)
    ws.Range("B9:M60").Font.Bold = False
    
    ' 헤더 설정
    ws.Range("B8").Value = "최초 언급"
    ws.Range("C8").Value = "이슈 제목"
    ws.Range("D8").Value = "카테고리"
    ws.Range("E8").Value = "상태"
    ws.Range("F8").Value = "담당부서"
    
    ' 타임라인 월 헤더 (2025년 6월부터 10월)
    ws.Cells(8, 7).Value = "2025-06"
    ws.Cells(8, 8).Value = "2025-07"
    ws.Cells(8, 9).Value = "2025-08"
    ws.Cells(8, 10).Value = "2025-09"
    ws.Cells(8, 11).Value = "2025-10"
    ws.Range("M8").Value = "관련문서"
    
    ' 헤더 서식
    ws.Range("B8:M8").Interior.Color = RGB(52, 73, 94)
    ws.Range("B8:M8").Font.Color = RGB(255, 255, 255)
    ws.Range("B8:M8").Font.Bold = True
    
    If issues.Count = 0 Then
        MsgBox "이슈가 없습니다"
        Exit Sub
    End If
    
    ' 첫 번째 이슈만 자세히 디버깅
    row = 9
    Dim firstIssue As Boolean
    firstIssue = True
    
    For Each issue In issues
        With ws
            ' 기본 데이터 입력
            .Cells(row, 2).Value = Left(issue("first_mentioned_date"), 10)
            .Cells(row, 3).Value = issue("title")
            .Cells(row, 4).Value = issue("category")
            .Cells(row, 5).Value = issue("status")
            .Cells(row, 6).Value = issue("department")
            
            ' 타임라인 그리기 (디버그)
            Call DrawDebugTimeline(ws, row, issue, firstIssue)
            firstIssue = False
            
            ' 관련문서
            .Cells(row, 13).Value = "문서 보기"
            .Cells(row, 13).Font.Color = RGB(0, 102, 204)
            .Cells(row, 13).Font.Underline = xlUnderlineStyleSingle
            
            ' 행 서식
            .Range(.Cells(row, 2), .Cells(row, 13)).Borders.LineStyle = xlContinuous
        End With
        
        row = row + 1
        If row > 15 Then Exit For  ' 처음 몇 개만 테스트
    Next issue
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "디버그 타임라인 완료", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "오류: " & Err.Description, vbCritical
End Sub

' 디버그 타임라인 그리기
Private Sub DrawDebugTimeline(ws As Worksheet, row As Integer, issue As Object, showDebug As Boolean)
    Dim startCol As Integer, endCol As Integer
    Dim issueDate As Date
    Dim monthDiff As Integer
    Dim col As Integer
    Dim debugMsg As String
    
    ' 날짜 파싱
    On Error Resume Next
    issueDate = CDate(Left(issue("first_mentioned_date"), 10))
    On Error GoTo 0
    
    ' 2025년 6월 1일 기준
    Dim baseDate As Date
    baseDate = DateSerial(2025, 6, 1)
    
    ' 월 차이 계산
    monthDiff = DateDiff("m", baseDate, issueDate)
    
    ' 시작 열 (7열이 6월, 8열이 7월, ...)
    startCol = 7 + monthDiff
    
    ' 범위 제한
    If startCol < 7 Then startCol = 7
    If startCol > 11 Then startCol = 11
    
    ' 종료 열 계산
    Select Case issue("status")
        Case "OPEN"
            endCol = 11  ' 10월까지
        Case "IN_PROGRESS", "MONITORING"
            endCol = 9   ' 8월까지 (현재)
        Case "RESOLVED"
            endCol = startCol + 1
            If endCol > 11 Then endCol = 11
    End Select
    
    ' 첫 번째 이슈만 디버그 정보 표시
    If showDebug Then
        debugMsg = "이슈: " & issue("issue_key") & vbCrLf
        debugMsg = debugMsg & "날짜: " & issue("first_mentioned_date") & vbCrLf
        debugMsg = debugMsg & "상태: " & issue("status") & vbCrLf
        debugMsg = debugMsg & "월차이: " & monthDiff & vbCrLf
        debugMsg = debugMsg & "시작열: " & startCol & " (열" & startCol & ")" & vbCrLf
        debugMsg = debugMsg & "종료열: " & endCol & " (열" & endCol & ")" & vbCrLf
        
        MsgBox debugMsg, vbInformation, "타임라인 디버그"
    End If
    
    ' 타임라인 바 그리기 - 직접 RGB 사용
    For col = startCol To endCol
        With ws.Cells(row, col)
            ' 상태별 색상 직접 적용
            Select Case issue("status")
                Case "OPEN"
                    .Interior.Color = RGB(255, 0, 0)      ' 빨강
                Case "IN_PROGRESS"
                    .Interior.Color = RGB(255, 165, 0)    ' 오렌지
                Case "RESOLVED"
                    .Interior.Color = RGB(0, 128, 0)      ' 초록
                Case "MONITORING"
                    .Interior.Color = RGB(0, 0, 255)      ' 파랑
                Case Else
                    .Interior.Color = RGB(200, 200, 200)  ' 회색
            End Select
            
            ' 시작점 마커
            If col = startCol Then
                .Value = "●"
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End If
        End With
    Next col
End Sub

' RefreshIssueTimeline에서 호출
Sub RefreshDebugTimeline()
    Call UpdateDebugTimeline
End Sub