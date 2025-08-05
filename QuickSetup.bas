Option Explicit

' STRIX 빠른 설정 - 모든 것을 한 번에!
' 이 파일만 import하면 모든 기능이 자동으로 설정됩니다

' ===== API 통신 함수 =====
Function AskSTRIX(question As String, Optional docType As String = "both") As String
    Dim http As Object
    Dim url As String
    Dim jsonBody As String
    Dim response As String
    
    On Error GoTo ErrorHandler
    
    ' HTTP 객체 생성
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' API URL
    url = "http://localhost:5000/search"
    
    ' JSON 요청 본문
    jsonBody = "{""question"":""" & question & """,""doc_type"":""" & docType & """}"
    
    ' API 호출
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonBody
    
    ' 응답 처리
    If http.Status = 200 Then
        response = http.responseText
        ' JSON 파싱 (간단한 답변 추출)
        Dim startPos As Long, endPos As Long
        startPos = InStr(response, """answer"":""") + 10
        endPos = InStr(startPos, response, """")
        AskSTRIX = Mid(response, startPos, endPos - startPos)
        ' 이스케이프 문자 처리
        AskSTRIX = Replace(AskSTRIX, "\n", vbLf)
        AskSTRIX = Replace(AskSTRIX, "\\", "\")
        AskSTRIX = Replace(AskSTRIX, "\""", """")
    Else
        AskSTRIX = "Error: API 서버 응답 오류 (" & http.Status & ")"
    End If
    
    Exit Function
    
ErrorHandler:
    AskSTRIX = "Error: " & Err.Description & vbLf & "API 서버가 실행 중인지 확인하세요."
End Function

' ===== 대화창 표시 =====
Sub ShowSTRIXDialog()
    Dim question As String
    Dim answer As String
    
    question = InputBox("STRIX에게 질문하세요:", "STRIX Intelligence", "전고체 배터리 개발 현황은?")
    
    If question <> "" Then
        answer = AskSTRIX(question)
        MsgBox answer, vbInformation, "STRIX 답변"
    End If
End Sub

' ===== 선택 영역 분석 =====
Sub AskAboutSelection()
    Dim selectedText As String
    Dim answer As String
    
    If TypeName(Selection) = "Range" Then
        selectedText = Selection.Value
        If selectedText <> "" Then
            answer = AskSTRIX("다음 내용을 분석해주세요: " & selectedText)
            Call DisplayAnswer(answer, selectedText)
        Else
            MsgBox "선택한 셀이 비어있습니다.", vbExclamation
        End If
    Else
        MsgBox "셀을 선택해주세요.", vbExclamation
    End If
End Sub

' ===== 답변 표시 =====
Sub DisplayAnswer(answer As String, Optional question As String = "")
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("STRIX Dashboard")
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ' Dashboard가 있으면 거기에 표시
        ws.Range("QuestionInput").Value = question
        ws.Range("AnswerDisplay").Value = answer
        ws.Range("StatusBar").Value = "✅ 검색 완료 - " & Now()
    Else
        ' Dashboard가 없으면 메시지 박스로 표시
        MsgBox answer, vbInformation, "STRIX 답변"
    End If
End Sub

' ===== Dashboard 생성 =====
Sub CreateSTRIXDashboard()
    Dim ws As Worksheet
    Dim btn As Object
    Dim shp As Shape
    
    ' Dashboard 시트 생성 또는 초기화
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("STRIX Dashboard")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "STRIX Dashboard"
    Else
        ' 기존 내용 모두 삭제
        ws.Cells.Clear
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If
    On Error GoTo 0
    
    ' 시트 활성화
    ws.Activate
    
    ' ==== 1. 전체 레이아웃 설정 ====
    With ws
        .Cells.Interior.Color = RGB(245, 245, 245)
        .Columns("A").ColumnWidth = 2
        .Columns("B:F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 2
        .Columns("H:J").ColumnWidth = 20
        .Columns("K").ColumnWidth = 2
    End With
    
    ' ==== 2. 헤더 영역 ====
    With ws.Range("B2:J2")
        .Merge
        .Value = "STRIX Intelligence Dashboard"
        .Font.Name = "Segoe UI"
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(41, 128, 185)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 40
    End With
    
    ' ==== 3. 질문 입력 영역 ====
    ws.Range("B4").Value = "질문:"
    ws.Range("B4").Font.Bold = True
    
    With ws.Range("C4:F4")
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Name = "QuestionInput"
        .Value = "여기에 질문을 입력하세요..."
        .Font.Color = RGB(150, 150, 150)
    End With
    
    ' ==== 4. 버튼 생성 ====
    Set btn = ws.Buttons.Add(ws.Range("B6").Left, ws.Range("B6").Top, 100, 30)
    With btn
        .Caption = "STRIX 대화창"
        .OnAction = "ShowSTRIXDialog"
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("C6").Left + 10, ws.Range("C6").Top, 100, 30)
    With btn
        .Caption = "검색 실행"
        .OnAction = "ExecuteSearch"
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("D6").Left + 20, ws.Range("D6").Top, 100, 30)
    With btn
        .Caption = "선택 분석"
        .OnAction = "AskAboutSelection"
    End With
    
    ' ==== 5. 답변 표시 영역 ====
    With ws.Range("B8:F8")
        .Merge
        .Value = "답변:"
        .Font.Bold = True
        .Interior.Color = RGB(46, 204, 113)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("B9:F20")
        .Merge
        .Name = "AnswerDisplay"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Value = "답변이 여기에 표시됩니다..."
        .Font.Color = RGB(150, 150, 150)
    End With
    
    ' ==== 6. 상태 표시 ====
    With ws.Range("B22:F22")
        .Merge
        .Name = "StatusBar"
        .Value = "✅ 준비 완료"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    ' ==== 7. 빠른 질문 버튼 ====
    ws.Range("H4").Value = "빠른 질문:"
    ws.Range("H4").Font.Bold = True
    
    Set btn = ws.Buttons.Add(ws.Range("H6").Left, ws.Range("H6").Top, 200, 25)
    With btn
        .Caption = "전고체 배터리 개발 현황"
        .OnAction = "QuickQuestion1"
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("H8").Left, ws.Range("H8").Top, 200, 25)
    With btn
        .Caption = "최근 배터리 시장 동향"
        .OnAction = "QuickQuestion2"
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("H10").Left, ws.Range("H10").Top, 200, 25)
    With btn
        .Caption = "경쟁사 기술 개발 현황"
        .OnAction = "QuickQuestion3"
    End With
    
    MsgBox "STRIX Dashboard가 생성되었습니다!" & vbCrLf & vbCrLf & _
           "API 서버 실행 명령:" & vbCrLf & _
           "py api_server.py", vbInformation, "STRIX"
End Sub

' ===== 검색 실행 =====
Sub ExecuteSearch()
    Dim question As String
    Dim answer As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("STRIX Dashboard")
    question = ws.Range("QuestionInput").Value
    
    If question = "" Or question = "여기에 질문을 입력하세요..." Then
        MsgBox "질문을 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    ws.Range("StatusBar").Value = "🔄 검색 중..."
    answer = AskSTRIX(question)
    
    With ws.Range("AnswerDisplay")
        .Value = answer
        .Font.Color = RGB(0, 0, 0)
    End With
    
    ws.Range("StatusBar").Value = "✅ 검색 완료 - " & Now()
End Sub

' ===== 빠른 질문 =====
Sub QuickQuestion1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("STRIX Dashboard")
    ws.Range("QuestionInput").Value = "전고체 배터리 개발 현황은?"
    ExecuteSearch
End Sub

Sub QuickQuestion2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("STRIX Dashboard")
    ws.Range("QuestionInput").Value = "최근 배터리 시장 동향은?"
    ExecuteSearch
End Sub

Sub QuickQuestion3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("STRIX Dashboard")
    ws.Range("QuestionInput").Value = "경쟁사의 기술 개발 현황은?"
    ExecuteSearch
End Sub

' ===== 셀 함수 =====
Function STRIX(question As String) As String
    On Error GoTo ErrorHandler
    STRIX = AskSTRIX(question)
    Exit Function
ErrorHandler:
    STRIX = "Error: " & Err.Description
End Function

' ===== 문서 업로드 (간단 버전) =====
Sub BulkUploadDocuments()
    MsgBox "문서 업로드 기능은 별도의 Python 스크립트를 사용하세요." & vbCrLf & _
           "py test_ingestion.py", vbInformation, "STRIX"
End Sub

' ===== 검색 기록 표시 =====
Sub ShowRecentSearches()
    MsgBox "검색 기록은 Dashboard에 자동으로 표시됩니다.", vbInformation, "STRIX"
End Sub