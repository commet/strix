Attribute VB_Name = "modRAGAPI"
' STRIX RAG API Connection Module
Option Explicit

Private Const API_TIMEOUT As Long = 30000

' 실제 RAG API 호출 함수 - Dictionary 방식
Function CallRAGAPI(question As String, Optional docType As String = "both") As Object
    Dim xmlHttp As Object
    Dim jsonRequest As Object
    Dim jsonResponse As Object
    Dim result As Object
    
    Set result = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrorHandler
    
    ' HTTP 요청 객체 생성
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' JSON 요청 데이터 생성
    Set jsonRequest = CreateObject("Scripting.Dictionary")
    jsonRequest.Add "question", question
    jsonRequest.Add "doc_type", docType
    
    ' API 호출
    With xmlHttp
        .Open "POST", "http://localhost:5000/api/query", False
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "Accept", "application/json"
        .setTimeouts API_TIMEOUT, API_TIMEOUT, API_TIMEOUT, API_TIMEOUT
        
        ' JSON 변환 및 전송
        .send JsonConverter.ConvertToJson(jsonRequest)
        
        ' 응답 확인
        If .Status = 200 Then
            ' JSON 응답 파싱
            Set jsonResponse = JsonConverter.ParseJson(.responseText)
            
            ' 결과 Dictionary에 저장
            result.Add "answer", jsonResponse("answer")
            result.Add "total_sources", jsonResponse("total_sources")
            result.Add "internal_docs", jsonResponse("internal_docs")
            result.Add "external_docs", jsonResponse("external_docs")
            result.Add "sources", jsonResponse("sources")
            result.Add "error", ""
            
        Else
            result.Add "error", "API 오류: " & .Status & " - " & .statusText
        End If
    End With
    
    Set CallRAGAPI = result
    Exit Function
    
ErrorHandler:
    result.Add "error", "오류 발생: " & Err.Description
    Set CallRAGAPI = result
End Function

' RAG 검색 실행
Sub RunRAGSearchWithSources()
    Dim ws As Worksheet
    Dim question As String
    Dim apiResponse As Object
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    question = ws.Range("C5").Value
    
    If question = "" Or question = "여기에 질문을 입력하세요" Then
        MsgBox "질문을 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' 상태 표시
    ws.Range("B64").Value = "⏳ AI 분석 중... (OpenAI API 호출)"
    ws.Range("B64").Font.Color = RGB(255, 140, 0)
    Application.StatusBar = "RAG 시스템에서 답변을 생성중입니다..."
    DoEvents
    
    ' RAG API 호출
    Set apiResponse = CallRAGAPI(question)
    
    ' 오류 확인
    If apiResponse("error") <> "" Then
        ' API 서버가 실행되지 않은 경우
        If InStr(apiResponse("error"), "개체가 필요합니다") > 0 Or _
           InStr(apiResponse("error"), "연결") > 0 Then
            
            ws.Range("B64").Value = "⚠️ API 서버 미실행"
            ws.Range("B64").Font.Color = RGB(255, 165, 0)
            MsgBox "API 서버가 실행되지 않았습니다." & vbCrLf & _
                   "터미널에서 다음 명령을 실행하세요:" & vbCrLf & vbCrLf & _
                   "py api_server_with_sources.py", vbInformation
            Exit Sub
        Else
            MsgBox "API 호출 실패: " & apiResponse("error"), vbCritical
            ws.Range("B64").Value = "❌ 오류 발생"
            ws.Range("B64").Font.Color = RGB(255, 0, 0)
            Exit Sub
        End If
    End If
    
    ' 답변 표시
    ws.Range("B10").Value = apiResponse("answer")
    ws.Range("B10").Font.Color = RGB(0, 0, 0)
    
    ' 소스 문서 표시
    DisplayRAGSourcesFromDict ws, 24, apiResponse("sources")
    
    ' 통계 정보 표시
    Dim statsMsg As String
    statsMsg = "✅ 검색 완료 - " & Format(Now, "hh:mm:ss") & _
               " | 참고문서: " & apiResponse("total_sources") & "개" & _
               " (내부: " & apiResponse("internal_docs") & ", 외부: " & apiResponse("external_docs") & ")"
    
    ws.Range("B64").Value = statsMsg
    ws.Range("B64").Font.Color = RGB(0, 150, 0)
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ws.Range("B64").Value = "❌ 오류 발생: " & Err.Description
    ws.Range("B64").Font.Color = RGB(255, 0, 0)
    Application.StatusBar = False
End Sub

' Dictionary에서 소스 문서 표시
Sub DisplayRAGSourcesFromDict(ws As Worksheet, startRow As Integer, sources As Variant)
    Dim currentRow As Integer
    Dim i As Integer
    Dim src As Object
    
    currentRow = startRow
    
    ' 기존 내용 지우기
    ws.Range("B" & startRow & ":F" & (startRow + 35)).Clear
    
    ' 헤더 스타일
    With ws.Range("B" & currentRow & ":G" & currentRow)
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
    End With
    
    ws.Cells(currentRow, 2).Value = "번호"
    ws.Cells(currentRow, 3).Value = "제목"
    ws.Cells(currentRow, 4).Value = "출처/조직"
    ws.Cells(currentRow, 5).Value = "날짜"
    ws.Cells(currentRow, 6).Value = "유형"
    ws.Cells(currentRow, 7).Value = "문서링크"
    
    currentRow = currentRow + 1
    
    ' 소스 데이터 표시
    If Not IsEmpty(sources) Then
        For i = 1 To sources.Count
            If currentRow > startRow + 35 Then Exit For
            
            Set src = sources(i)
            
            ws.Cells(currentRow, 2).Value = "[" & i & "]"
            ws.Cells(currentRow, 2).Font.Bold = True
            ws.Cells(currentRow, 2).Font.Color = RGB(0, 112, 192)
            
            ws.Cells(currentRow, 3).Value = src("title")
            ws.Cells(currentRow, 3).WrapText = True
            
            ws.Cells(currentRow, 4).Value = src("organization")
            ws.Cells(currentRow, 5).Value = src("date")
            
            ' 영어를 한글로 변환하여 표시
            Dim displayType As String
            Select Case src("type")
                Case "internal"
                    displayType = "사내"
                Case "external"
                    displayType = "사외"
                Case Else
                    displayType = src("type")
            End Select
            ws.Cells(currentRow, 6).Value = displayType
            
            ' 관련도 점수가 있으면 표시
            If src.Exists("relevance_score") Then
                If src("relevance_score") > 0 Then
                    ws.Cells(currentRow, 3).Value = src("title") & " (" & Format(src("relevance_score") * 100, "0") & "%)"
                End If
            End If
            
            ' URL 추가 (하이퍼링크로)
            If src.Exists("url") Then
                Dim linkUrl As String
                linkUrl = src("url")
                If linkUrl <> "" Then
                    ws.Hyperlinks.Add Anchor:=ws.Cells(currentRow, 7), _
                                     Address:=linkUrl, _
                                     TextToDisplay:="열기 →"
                    ws.Cells(currentRow, 7).Font.Color = RGB(0, 102, 204)
                    ws.Cells(currentRow, 7).Font.Underline = True
                End If
            End If
            
            ' 행 서식 (줄무늬) - 유형 열은 제외
            With ws.Range("B" & currentRow & ":E" & currentRow)
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200)
                If i Mod 2 = 0 Then
                    .Interior.Color = RGB(248, 248, 248)
                End If
            End With
            
            ' F열(유형), G열(링크) 테두리만 추가
            With ws.Range("F" & currentRow & ":G" & currentRow)
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200)
            End With
            
            ' 유형별 색상 코딩 (줄무늬 뒤에 적용)
            Select Case src("type")
                Case "internal"
                    ws.Cells(currentRow, 6).Interior.Color = RGB(255, 242, 204)
                Case "external"  
                    ws.Cells(currentRow, 6).Interior.Color = RGB(217, 234, 211)
            End Select
            
            ws.Rows(currentRow).RowHeight = 20
            currentRow = currentRow + 1
        Next i
    End If
End Sub

' API 서버 상태 확인
Function CheckAPIServer() As Boolean
    Dim xmlHttp As Object
    
    On Error Resume Next
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    With xmlHttp
        .Open "GET", "http://localhost:5000/health", False
        .setTimeouts 3000, 3000, 3000, 3000
        .send
        
        If .Status = 200 Then
            CheckAPIServer = True
        Else
            CheckAPIServer = False
        End If
    End With
    
    On Error GoTo 0
End Function