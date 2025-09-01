Attribute VB_Name = "modIssueAPI"
' Issue Tracking API Integration Module
Option Explicit

' API 기본 URL
Private Const API_BASE_URL As String = "http://localhost:5001/api"

' 이슈 목록 가져오기
Function GetIssueList(Optional category As String = "", _
                     Optional status As String = "", _
                     Optional days As Integer = 90) As Collection
    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim issues As New Collection
    
    On Error GoTo ErrorHandler
    
    ' URL 구성
    url = API_BASE_URL & "/issues?"
    If category <> "" And category <> "전체" Then
        url = url & "category=" & URLEncode(category) & "&"
    End If
    If status <> "" And status <> "전체" Then
        url = url & "status=" & URLEncode(status) & "&"
    End If
    url = url & "days=" & days
    
    ' HTTP 요청
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send
    
    If http.Status = 200 Then
        On Error Resume Next
        responseText = BytesToString(http.responseBody, "UTF-8")
        If Err.Number <> 0 Then
            ' BytesToString 실패시 대체 방법
            responseText = http.responseText
        End If
        On Error GoTo ErrorHandler
        Set issues = ParseIssueList(responseText)
    End If
    
    Set GetIssueList = issues
    Exit Function
    
ErrorHandler:
    Debug.Print "Error getting issue list: " & Err.Description
    Set GetIssueList = issues
End Function

' 이슈 상세 정보 가져오기
Function GetIssueDetail(issueId As String) As Object
    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim detail As Object
    
    On Error GoTo ErrorHandler
    
    url = API_BASE_URL & "/issues/" & issueId
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send
    
    If http.Status = 200 Then
        On Error Resume Next
        responseText = BytesToString(http.responseBody, "UTF-8")
        If Err.Number <> 0 Then
            responseText = http.responseText
        End If
        On Error GoTo ErrorHandler
        Set detail = ParseIssueDetail(responseText)
    End If
    
    Set GetIssueDetail = detail
    Exit Function
    
ErrorHandler:
    Debug.Print "Error getting issue detail: " & Err.Description
    Set GetIssueDetail = Nothing
End Function

' 이슈 타임라인 가져오기
Function GetIssueTimeline(issueId As String) As Collection
    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim timeline As New Collection
    
    On Error GoTo ErrorHandler
    
    url = API_BASE_URL & "/issues/" & issueId & "/timeline"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send
    
    If http.Status = 200 Then
        On Error Resume Next
        responseText = BytesToString(http.responseBody, "UTF-8")
        If Err.Number <> 0 Then
            responseText = http.responseText
        End If
        On Error GoTo ErrorHandler
        Set timeline = ParseTimeline(responseText)
    End If
    
    Set GetIssueTimeline = timeline
    Exit Function
    
ErrorHandler:
    Debug.Print "Error getting timeline: " & Err.Description
    Set GetIssueTimeline = timeline
End Function

' AI 분석 가져오기
Function GetAIAnalysis(issueId As String) As Object
    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim analysis As Object
    
    On Error GoTo ErrorHandler
    
    url = API_BASE_URL & "/issues/" & issueId & "/ai-analysis"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send
    
    If http.Status = 200 Then
        On Error Resume Next
        responseText = BytesToString(http.responseBody, "UTF-8")
        If Err.Number <> 0 Then
            responseText = http.responseText
        End If
        On Error GoTo ErrorHandler
        Set analysis = ParseAIAnalysis(responseText)
    End If
    
    Set GetAIAnalysis = analysis
    Exit Function
    
ErrorHandler:
    Debug.Print "Error getting AI analysis: " & Err.Description
    Set GetAIAnalysis = Nothing
End Function

' 대시보드 요약 정보 가져오기
Function GetDashboardSummary() As Object
    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim summary As Object
    
    On Error GoTo ErrorHandler
    
    url = API_BASE_URL & "/issues/dashboard-summary"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send
    
    If http.Status = 200 Then
        On Error Resume Next
        responseText = BytesToString(http.responseBody, "UTF-8")
        If Err.Number <> 0 Then
            responseText = http.responseText
        End If
        On Error GoTo ErrorHandler
        Set summary = ParseDashboardSummary(responseText)
    End If
    
    Set GetDashboardSummary = summary
    Exit Function
    
ErrorHandler:
    Debug.Print "Error getting dashboard summary: " & Err.Description
    Set GetDashboardSummary = Nothing
End Function

' ===== 파싱 함수들 =====

' 이슈 목록 파싱
Private Function ParseIssueList(jsonStr As String) As Collection
    Dim issues As New Collection
    Dim startPos As Long, endPos As Long
    Dim issueJson As String
    
    ' 디버그 출력
    Debug.Print "ParseIssueList - JSON Length: " & Len(jsonStr)
    Debug.Print "ParseIssueList - First 200 chars: " & Left(jsonStr, 200)
    
    ' 간단한 JSON 배열 파싱
    startPos = InStr(1, jsonStr, "[")
    endPos = InStrRev(jsonStr, "]")
    
    If startPos > 0 And endPos > startPos Then
        Dim arrayContent As String
        arrayContent = Mid(jsonStr, startPos + 1, endPos - startPos - 1)
        Debug.Print "ParseIssueList - Array content length: " & Len(arrayContent)
        
        ' 각 객체 파싱
        Dim objects() As String
        objects = SplitJSONObjects(arrayContent)
        Debug.Print "ParseIssueList - Objects count: " & (UBound(objects) + 1)
        
        Dim i As Integer
        For i = 0 To UBound(objects)
            If Trim(objects(i)) <> "" Then
                Dim issue As Object
                Set issue = ParseSingleIssue(objects(i))
                If Not issue Is Nothing Then
                    issues.Add issue
                    Debug.Print "ParseIssueList - Added issue: " & issue("title")
                End If
            End If
        Next i
    Else
        Debug.Print "ParseIssueList - No valid JSON array found"
    End If
    
    Debug.Print "ParseIssueList - Total issues: " & issues.Count
    Set ParseIssueList = issues
End Function

' 단일 이슈 파싱
Private Function ParseSingleIssue(issueJson As String) As Object
    Dim issue As Object
    Set issue = CreateObject("Scripting.Dictionary")
    
    issue("id") = ExtractJsonValue(issueJson, "id")
    issue("issue_key") = ExtractJsonValue(issueJson, "issue_key")
    issue("title") = ExtractJsonValue(issueJson, "title")
    issue("category") = ExtractJsonValue(issueJson, "category")
    issue("priority") = ExtractJsonValue(issueJson, "priority")
    issue("status") = ExtractJsonValue(issueJson, "status")
    issue("department") = ExtractJsonValue(issueJson, "department")
    issue("owner") = ExtractJsonValue(issueJson, "owner")
    issue("first_mentioned_date") = ExtractJsonValue(issueJson, "first_mentioned_date")
    issue("last_updated") = ExtractJsonValue(issueJson, "last_updated")
    issue("document_count") = Val(ExtractJsonValue(issueJson, "document_count"))
    
    Set ParseSingleIssue = issue
End Function

' 이슈 상세 정보 파싱
Private Function ParseIssueDetail(jsonStr As String) As Object
    Dim detail As Object
    Set detail = CreateObject("Scripting.Dictionary")
    
    ' 기본 이슈 정보
    Dim issueStart As Long
    issueStart = InStr(1, jsonStr, """issue"":")
    If issueStart > 0 Then
        ' issue 객체 추출 및 파싱
        ' ... 구현 ...
    End If
    
    ' 문서 목록
    Dim docsStart As Long
    docsStart = InStr(1, jsonStr, """documents"":")
    If docsStart > 0 Then
        ' documents 배열 추출 및 파싱
        ' ... 구현 ...
    End If
    
    Set ParseIssueDetail = detail
End Function

' 타임라인 파싱
Private Function ParseTimeline(jsonStr As String) As Collection
    Dim timeline As New Collection
    ' JSON 배열 파싱 로직
    ' ... 구현 ...
    Set ParseTimeline = timeline
End Function

' AI 분석 파싱
Private Function ParseAIAnalysis(jsonStr As String) As Object
    Dim analysis As Object
    Set analysis = CreateObject("Scripting.Dictionary")
    
    analysis("summary") = ExtractJsonValue(jsonStr, "summary")
    analysis("confidence") = Val(ExtractJsonValue(jsonStr, "confidence"))
    
    ' risks 배열 파싱
    ' recommendations 배열 파싱
    ' ... 구현 ...
    
    Set ParseAIAnalysis = analysis
End Function

' 대시보드 요약 파싱
Private Function ParseDashboardSummary(jsonStr As String) As Object
    Dim summary As Object
    Set summary = CreateObject("Scripting.Dictionary")
    
    ' statistics 객체 파싱
    ' category_distribution 배열 파싱
    ' ... 구현 ...
    
    Set ParseDashboardSummary = summary
End Function

' ===== 유틸리티 함수들 =====

' URL 인코딩
Private Function URLEncode(text As String) As String
    Dim i As Integer
    Dim result As String
    
    For i = 1 To Len(text)
        Dim char As String
        char = Mid(text, i, 1)
        
        Select Case Asc(char)
            Case 48 To 57, 65 To 90, 97 To 122 ' 0-9, A-Z, a-z
                result = result & char
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(char)), 2)
        End Select
    Next i
    
    URLEncode = result
End Function

' JSON 객체 배열 분리
Private Function SplitJSONObjects(jsonArrayStr As String) As String()
    ' modSTRIXwithSources에서 가져온 함수
    Dim result() As String
    Dim currentObj As String
    Dim braceCount As Integer
    Dim inString As Boolean
    Dim escaped As Boolean
    Dim i As Long
    Dim objCount As Integer
    Dim ch As String
    
    ReDim result(0)
    currentObj = ""
    braceCount = 0
    inString = False
    escaped = False
    objCount = 0
    
    For i = 1 To Len(jsonArrayStr)
        ch = Mid(jsonArrayStr, i, 1)
        
        If ch = """" And Not escaped Then
            inString = Not inString
        End If
        
        If ch = "\" And Not escaped Then
            escaped = True
        Else
            escaped = False
        End If
        
        If Not inString Then
            If ch = "{" Then
                braceCount = braceCount + 1
            ElseIf ch = "}" Then
                braceCount = braceCount - 1
            End If
        End If
        
        currentObj = currentObj & ch
        
        If braceCount = 0 And Len(Trim(currentObj)) > 0 And InStr(currentObj, "}") > 0 Then
            ReDim Preserve result(objCount)
            result(objCount) = Trim(currentObj)
            objCount = objCount + 1
            currentObj = ""
        End If
    Next i
    
    SplitJSONObjects = result
End Function

' JSON 값 추출
Private Function ExtractJsonValue(json As String, key As String) As String
    ' modSTRIXwithSources에서 가져온 함수
    Dim startPos As Long, endPos As Long
    Dim searchKey As String
    
    searchKey = """" & key & """: """
    startPos = InStr(1, json, searchKey)
    
    If startPos = 0 Then
        searchKey = """" & key & """: "
        startPos = InStr(1, json, searchKey)
        If startPos > 0 Then
            startPos = startPos + Len(searchKey)
            endPos = InStr(startPos, json, ",")
            If endPos = 0 Then endPos = InStr(startPos, json, "}")
            ExtractJsonValue = Trim(Mid(json, startPos, endPos - startPos))
        Else
            ExtractJsonValue = ""
        End If
    Else
        startPos = startPos + Len(searchKey)
        endPos = startPos
        
        Dim escaped As Boolean
        escaped = False
        Do While endPos <= Len(json)
            If Mid(json, endPos, 1) = "\" And Not escaped Then
                escaped = True
            ElseIf Mid(json, endPos, 1) = """" And Not escaped Then
                Exit Do
            Else
                escaped = False
            End If
            endPos = endPos + 1
        Loop
        
        If endPos > startPos Then
            ExtractJsonValue = Mid(json, startPos, endPos - startPos)
            ExtractJsonValue = Replace(ExtractJsonValue, "\""", """")
            ExtractJsonValue = Replace(ExtractJsonValue, "\\", "\")
            ExtractJsonValue = Replace(ExtractJsonValue, "\/", "/")
        Else
            ExtractJsonValue = ""
        End If
    End If
End Function

' BytesToString (UTF-8 변환)
Private Function BytesToString(bytes() As Byte, charset As String) As String
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    objStream.Type = 1  ' adTypeBinary
    objStream.Open
    objStream.Write bytes
    objStream.Position = 0
    objStream.Type = 2  ' adTypeText
    objStream.charset = charset
    
    BytesToString = objStream.ReadText
    objStream.Close
End Function

' ===== 타임라인 업데이트 함수 =====

Sub UpdateIssueTimeline()
    Dim ws As Worksheet
    Dim issues As Collection
    Dim issue As Object
    Dim row As Integer
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    ' 상태 표시
    Application.StatusBar = "이슈 목록을 가져오는 중..."
    
    ' API에서 이슈 목록 가져오기
    Set issues = GetIssueList()
    
    ' 기존 데이터 지우기
    ws.Range("B9:K50").Clear
    
    ' 헤더 다시 설정
    row = 9
    
    ' 각 이슈 표시
    For Each issue In issues
        With ws
            .Cells(row, 2).Value = Left(issue("first_mentioned_date"), 10)
            .Cells(row, 3).Value = issue("title")
            .Cells(row, 4).Value = issue("category")
            .Cells(row, 5).Value = MapStatusKorean(issue("status"))
            .Cells(row, 6).Value = issue("department")
            
            ' 상태별 색상
            Select Case issue("status")
                Case "OPEN"
                    .Cells(row, 5).Font.Color = RGB(231, 76, 60)
                Case "IN_PROGRESS"
                    .Cells(row, 5).Font.Color = RGB(241, 196, 15)
                Case "RESOLVED"
                    .Cells(row, 5).Font.Color = RGB(46, 204, 113)
                Case "MONITORING"
                    .Cells(row, 5).Font.Color = RGB(52, 152, 219)
            End Select
            
            ' 행 서식
            With .Range(.Cells(row, 2), .Cells(row, 11))
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200)
                If row Mod 2 = 0 Then
                    .Interior.Color = RGB(248, 248, 248)
                End If
            End With
            
            ' 이슈 ID 저장 (숨김 열에)
            .Cells(row, 12).Value = issue("id")
        End With
        
        row = row + 1
        If row > 50 Then Exit For ' 최대 표시 수 제한
    Next issue
    
    Application.StatusBar = False
    MsgBox "타임라인이 업데이트되었습니다: " & issues.Count & "개 이슈", vbInformation
End Sub

' 상태 한글 변환
Private Function MapStatusKorean(status As String) As String
    Select Case status
        Case "OPEN": MapStatusKorean = "미해결"
        Case "IN_PROGRESS": MapStatusKorean = "진행중"
        Case "RESOLVED": MapStatusKorean = "해결됨"
        Case "MONITORING": MapStatusKorean = "모니터링"
        Case Else: MapStatusKorean = status
    End Select
End Function