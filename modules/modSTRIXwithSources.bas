Attribute VB_Name = "modSTRIXwithSources"
' Module for STRIX with Source References
Option Explicit

' STRIX에 질문하고 소스 포함 답변 받기
Function AskSTRIXWithSources(question As String) As Variant
    Dim http As Object
    Dim url As String
    Dim jsonBody As String
    Dim responseBytes() As Byte
    Dim responseText As String
    Dim result(0 To 2) As Variant ' 0: answer, 1: sources collection, 2: error
    
    On Error GoTo ErrorHandler
    
    ' WinHTTP 사용
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://localhost:5000/api/query"
    
    ' JSON 본문
    jsonBody = "{""question"":""" & question & """,""doc_type"":""both""}"
    
    ' HTTP 요청
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send jsonBody
    
    If http.Status = 200 Then
        ' 바이트 배열로 받아서 UTF-8 변환
        responseBytes = http.responseBody
        responseText = BytesToString(responseBytes, "UTF-8")
        
        ' JSON 파싱
        result(0) = ExtractAnswer(responseText)
        Set result(1) = ExtractSourcesAsCollection(responseText)
        result(2) = ""
    Else
        result(0) = ""
        Set result(1) = New Collection
        result(2) = "Error: HTTP " & http.Status
    End If
    
    AskSTRIXWithSources = result
    Exit Function
    
ErrorHandler:
    result(0) = ""
    Set result(1) = New Collection
    result(2) = "Error: " & Err.Description
    AskSTRIXWithSources = result
End Function

' 소스 문서를 Collection으로 추출
Function ExtractSourcesAsCollection(jsonStr As String) As Collection
    Dim sources As New Collection
    Dim startPos As Long, endPos As Long
    Dim sourcesJson As String
    Dim i As Integer
    
    ' "sources": [ 찾기
    startPos = InStr(1, jsonStr, """sources"": [")
    If startPos = 0 Then
        startPos = InStr(1, jsonStr, """sources"":[")
        If startPos > 0 Then startPos = startPos + 11
    Else
        startPos = startPos + 12
    End If
    
    If startPos = 0 Then
        Set ExtractSourcesAsCollection = sources
        Exit Function
    End If
    
    ' sources 배열의 끝 찾기
    Dim bracketCount As Integer
    bracketCount = 1
    endPos = startPos
    
    Do While bracketCount > 0 And endPos < Len(jsonStr)
        If Mid(jsonStr, endPos, 1) = "[" Then
            bracketCount = bracketCount + 1
        ElseIf Mid(jsonStr, endPos, 1) = "]" Then
            bracketCount = bracketCount - 1
        End If
        endPos = endPos + 1
    Loop
    
    If endPos > startPos Then
        sourcesJson = Mid(jsonStr, startPos, endPos - startPos - 1)
        
        ' 각 소스 객체 파싱
        Dim sourceTexts() As String
        sourceTexts = SplitJSONObjects(sourcesJson)
        
        For i = 0 To UBound(sourceTexts)
            If Trim(sourceTexts(i)) <> "" Then
                Dim sourceDict As Object
                Set sourceDict = ParseSourceToDict(sourceTexts(i))
                If Not sourceDict Is Nothing Then
                    sources.Add sourceDict
                End If
            End If
        Next i
    End If
    
    Set ExtractSourcesAsCollection = sources
End Function

' JSON 객체 배열을 개별 객체로 분리
Function SplitJSONObjects(jsonArrayStr As String) As String()
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
        
        ' 문자열 내부 여부 체크
        If ch = """" And Not escaped Then
            inString = Not inString
        End If
        
        ' 이스케이프 체크
        If ch = "\" And Not escaped Then
            escaped = True
        Else
            escaped = False
        End If
        
        ' 중괄호 카운트 (문자열 외부에서만)
        If Not inString Then
            If ch = "{" Then
                braceCount = braceCount + 1
            ElseIf ch = "}" Then
                braceCount = braceCount - 1
            End If
        End If
        
        currentObj = currentObj & ch
        
        ' 객체 완성
        If braceCount = 0 And Len(Trim(currentObj)) > 0 And InStr(currentObj, "}") > 0 Then
            ReDim Preserve result(objCount)
            result(objCount) = Trim(currentObj)
            objCount = objCount + 1
            currentObj = ""
        End If
    Next i
    
    SplitJSONObjects = result
End Function

' 개별 소스를 Dictionary로 파싱
Function ParseSourceToDict(sourceJson As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 각 필드 추출
    dict("number") = Val(ExtractJsonValue(sourceJson, "number"))
    dict("type") = ExtractJsonValue(sourceJson, "type")
    dict("title") = ExtractJsonValue(sourceJson, "title")
    dict("organization") = ExtractJsonValue(sourceJson, "organization")
    dict("date") = ExtractJsonValue(sourceJson, "date")
    dict("snippet") = ExtractJsonValue(sourceJson, "snippet")
    
    Set ParseSourceToDict = dict
End Function

' JSON 값 추출 헬퍼
Function ExtractJsonValue(json As String, key As String) As String
    Dim startPos As Long, endPos As Long
    Dim searchKey As String
    
    searchKey = """" & key & """: """
    startPos = InStr(1, json, searchKey)
    
    If startPos = 0 Then
        ' 숫자 값인 경우
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
        
        ' 이스케이프된 따옴표 처리
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
            ' 이스케이프 문자 처리
            ExtractJsonValue = Replace(ExtractJsonValue, "\""", """")
            ExtractJsonValue = Replace(ExtractJsonValue, "\\", "\")
            ExtractJsonValue = Replace(ExtractJsonValue, "\/", "/")
        Else
            ExtractJsonValue = ""
        End If
    End If
End Function

' Dashboard에서 소스 포함 검색 실행
Sub RunSearchWithSources()
    Dim ws As Worksheet
    Dim question As String
    Dim result As Variant
    Dim answer As String
    Dim sources As Collection
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    question = ws.Range("C5").Value
    
    If question = "" Or question = "여기에 질문을 입력하세요" Then
        MsgBox "질문을 입력하세요", vbExclamation
        Exit Sub
    End If
    
    ' 상태 표시
    ws.Range("B10").Value = "검색 중..."
    ws.Range("B10").Font.Color = RGB(0, 0, 255)
    ws.Range("B41").Value = "🔍 검색 중..."
    DoEvents
    
    ' API 호출
    result = AskSTRIXWithSources(question)
    
    If result(2) <> "" Then
        ' 오류 처리
        ws.Range("B10").Value = result(2)
        ws.Range("B10").Font.Color = RGB(255, 0, 0)
        ws.Range("B41").Value = "❌ " & result(2)
        Exit Sub
    End If
    
    ' 답변 표시
    answer = result(0)
    Set sources = result(1)
    
    ' 답변 영역에 표시
    With ws.Range("B10")
        .Value = answer
        .Font.Color = RGB(0, 0, 0)
        .WrapText = True
    End With
    
    ' 레퍼런스 표시
    DisplaySourcesCollection ws, sources
    
    ' 상태 업데이트
    ws.Range("B41").Value = "✅ 검색 완료 - " & sources.Count & "개 참고문서 - " & Format(Now, "yyyy-mm-dd hh:mm:ss")
End Sub

' 소스 문서 표시 (Collection 버전)
Sub DisplaySourcesCollection(ws As Worksheet, sources As Collection)
    Dim startRow As Integer
    Dim i As Integer
    Dim src As Object
    
    startRow = 24 ' 레퍼런스 시작 행
    
    ' 기존 레퍼런스 영역 초기화
    ws.Range("B24:F35").Clear
    With ws.Range("B24:F35")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    
    ' 레퍼런스가 없으면 종료
    If sources.Count = 0 Then
        ws.Range("B24").Value = "참고 문서가 없습니다"
        Exit Sub
    End If
    
    ' 각 소스 표시
    For i = 1 To sources.Count
        Set src = sources(i)
        
        ' 데이터 입력
        ws.Cells(startRow, 2).Value = "[" & src("number") & "]"
        ws.Cells(startRow, 3).Value = src("title")
        ws.Cells(startRow, 4).Value = src("organization")
        ws.Cells(startRow, 5).Value = src("date")
        ws.Cells(startRow, 6).Value = IIf(src("type") = "internal", "내부문서", "외부뉴스")
        
        ' 서식 설정
        ws.Range(ws.Cells(startRow, 2), ws.Cells(startRow, 6)).Borders.LineStyle = xlContinuous
        
        ' 타입별 색상
        If src("type") = "internal" Then
            ws.Cells(startRow, 6).Font.Color = RGB(0, 100, 0)
        Else
            ws.Cells(startRow, 6).Font.Color = RGB(0, 0, 200)
        End If
        
        ' 번호 굵게
        ws.Cells(startRow, 2).Font.Bold = True
        
        startRow = startRow + 1
        
        ' 최대 표시 개수 제한
        If startRow > 35 Then Exit For
    Next i
End Sub

' 바이트 배열을 문자열로 변환 (UTF-8)
Function BytesToString(bytes() As Byte, charset As String) As String
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

' 기본 답변 추출
Function ExtractAnswer(jsonStr As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim answer As String
    
    startPos = InStr(1, jsonStr, """answer"": """)
    If startPos = 0 Then
        startPos = InStr(1, jsonStr, """answer"":""")
        If startPos > 0 Then startPos = startPos + 10
    Else
        startPos = startPos + 11
    End If
    
    If startPos > 10 Then
        ' 답변의 끝 찾기
        Dim i As Long
        Dim escaped As Boolean
        escaped = False
        
        For i = startPos To Len(jsonStr)
            If Mid(jsonStr, i, 1) = "\" And Not escaped Then
                escaped = True
            ElseIf Mid(jsonStr, i, 1) = """" And Not escaped Then
                endPos = i
                Exit For
            Else
                escaped = False
            End If
        Next i
        
        If endPos > startPos Then
            answer = Mid(jsonStr, startPos, endPos - startPos)
            
            ' 이스케이프 문자 처리
            answer = Replace(answer, "\n", vbLf)
            answer = Replace(answer, "\\", "\")
            answer = Replace(answer, "\""", """")
            answer = Replace(answer, "\/", "/")
            
            ExtractAnswer = answer
        Else
            ExtractAnswer = "답변을 파싱할 수 없습니다"
        End If
    Else
        ExtractAnswer = "답변을 찾을 수 없습니다"
    End If
End Function