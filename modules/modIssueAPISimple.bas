Attribute VB_Name = "modIssueAPISimple"
' Simplified Issue API Module for Testing
Option Explicit

' API 기본 URL
Private Const API_BASE_URL As String = "http://localhost:5001/api"

' 단순화된 이슈 목록 가져오기
Function GetSimpleIssueList() As Collection
    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim issues As New Collection
    
    On Error GoTo ErrorHandler
    
    ' URL 구성
    url = API_BASE_URL & "/issues"
    
    ' HTTP 요청
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send
    
    If http.Status = 200 Then
        responseText = http.responseText
        Debug.Print "Response received, length: " & Len(responseText)
        
        ' 간단한 파싱 - 각 이슈 객체를 찾아서 처리
        Dim pos As Long, endPos As Long
        Dim issueStr As String
        Dim issueCount As Integer
        
        pos = 1
        issueCount = 0
        
        Do While InStr(pos, responseText, """id"":") > 0
            pos = InStr(pos, responseText, """id"":")
            If pos = 0 Then Exit Do
            
            ' 다음 이슈 객체의 시작 또는 배열의 끝 찾기
            endPos = InStr(pos + 1, responseText, """id"":")
            If endPos = 0 Then
                endPos = InStr(pos, responseText, "]")
            End If
            
            If endPos > pos Then
                issueStr = Mid(responseText, pos, endPos - pos)
                
                ' 이슈 객체 생성
                Dim issue As Object
                Set issue = CreateObject("Scripting.Dictionary")
                
                ' 간단한 필드 추출
                issue("id") = ExtractValue(issueStr, "id")
                issue("title") = ExtractValue(issueStr, "title")
                issue("category") = ExtractValue(issueStr, "category")
                issue("status") = ExtractValue(issueStr, "status")
                issue("department") = ExtractValue(issueStr, "department")
                issue("owner") = ExtractValue(issueStr, "owner")
                issue("first_mentioned_date") = ExtractValue(issueStr, "first_mentioned_date")
                issue("priority") = ExtractValue(issueStr, "priority")
                
                issues.Add issue
                issueCount = issueCount + 1
                Debug.Print "Added issue #" & issueCount & ": " & issue("title")
            End If
            
            pos = pos + 1
        Loop
    End If
    
    Debug.Print "Total issues parsed: " & issues.Count
    Set GetSimpleIssueList = issues
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in GetSimpleIssueList: " & Err.Description
    Set GetSimpleIssueList = issues
End Function

' 간단한 값 추출 함수
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
    
    ' Try without quotes (for numbers, booleans, null)
    searchKey = """" & key & """: "
    startPos = InStr(1, json, searchKey)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchKey)
        endPos = InStr(startPos, json, ",")
        If endPos = 0 Then endPos = InStr(startPos, json, "}")
        If endPos > startPos Then
            ExtractValue = Trim(Mid(json, startPos, endPos - startPos))
            Exit Function
        End If
    End If
    
    ExtractValue = ""
End Function

' 간단한 타임라인 업데이트
Sub UpdateSimpleTimeline()
    Dim ws As Worksheet
    Dim issues As Collection
    Dim issue As Object
    Dim row As Integer
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    ' 상태 표시
    Application.StatusBar = "이슈 목록을 가져오는 중..."
    
    ' API에서 이슈 목록 가져오기
    Set issues = GetSimpleIssueList()
    
    ' 기존 데이터 지우기 (헤더 제외)
    ws.Range("B9:K50").ClearContents
    ws.Range("B9:K50").Borders.LineStyle = xlNone
    ws.Range("B9:K50").Interior.Color = xlNone
    
    ' 이슈가 없으면 메시지 표시
    If issues.Count = 0 Then
        Application.StatusBar = False
        MsgBox "가져온 이슈가 없습니다.", vbInformation
        Exit Sub
    End If
    
    ' 각 이슈 표시
    row = 9
    For Each issue In issues
        With ws
            ' 날짜
            If issue("first_mentioned_date") <> "" Then
                .Cells(row, 2).Value = Left(issue("first_mentioned_date"), 10)
            End If
            
            ' 제목
            .Cells(row, 3).Value = issue("title")
            
            ' 카테고리
            .Cells(row, 4).Value = issue("category")
            
            ' 상태 (한글로 변환)
            Select Case issue("status")
                Case "OPEN"
                    .Cells(row, 5).Value = "미해결"
                    .Cells(row, 5).Font.Color = RGB(231, 76, 60)
                Case "IN_PROGRESS"
                    .Cells(row, 5).Value = "진행중"
                    .Cells(row, 5).Font.Color = RGB(241, 196, 15)
                Case "RESOLVED"
                    .Cells(row, 5).Value = "해결됨"
                    .Cells(row, 5).Font.Color = RGB(46, 204, 113)
                Case "MONITORING"
                    .Cells(row, 5).Value = "모니터링"
                    .Cells(row, 5).Font.Color = RGB(52, 152, 219)
                Case Else
                    .Cells(row, 5).Value = issue("status")
            End Select
            
            ' 부서
            .Cells(row, 6).Value = issue("department")
            
            ' 타임라인 바 그리기 (임시)
            Dim startCol As Integer, endCol As Integer
            startCol = 7
            endCol = 9
            
            Select Case issue("status")
                Case "OPEN"
                    Call DrawSimpleTimelineBar(ws, row, startCol, endCol, RGB(231, 76, 60))
                Case "IN_PROGRESS"
                    Call DrawSimpleTimelineBar(ws, row, startCol, endCol, RGB(241, 196, 15))
                Case "RESOLVED"
                    Call DrawSimpleTimelineBar(ws, row, startCol, endCol, RGB(46, 204, 113))
                Case "MONITORING"
                    Call DrawSimpleTimelineBar(ws, row, startCol, endCol, RGB(52, 152, 219))
            End Select
            
            ' 행 서식
            With .Range(.Cells(row, 2), .Cells(row, 11))
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200)
                If row Mod 2 = 0 Then
                    .Interior.Color = RGB(248, 248, 248)
                End If
            End With
            
            ' 이슈 ID 저장 (숨김 열)
            .Cells(row, 12).Value = issue("id")
        End With
        
        row = row + 1
        If row > 50 Then Exit For ' 최대 표시 수 제한
    Next issue
    
    Application.StatusBar = False
    MsgBox "타임라인이 업데이트되었습니다: " & issues.Count & "개 이슈", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "업데이트 중 오류 발생: " & Err.Description, vbCritical
End Sub

' 간단한 타임라인 바 그리기
Private Sub DrawSimpleTimelineBar(ws As Worksheet, row As Integer, startCol As Integer, endCol As Integer, barColor As Long)
    Dim cell As Range
    For Each cell In ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol))
        cell.Interior.Color = barColor
        cell.Interior.Pattern = xlSolid
    Next cell
End Sub