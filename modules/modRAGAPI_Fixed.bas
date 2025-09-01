Attribute VB_Name = "modRAGAPI"
' STRIX RAG API Connection Module
' OpenAI API를 활용한 실제 RAG 시스템 연결
Option Explicit

Private Const API_TIMEOUT As Long = 30000 ' 30초 타임아웃

' API URL 함수로 정의
Private Function API_URL() As String
    API_URL = "http://localhost:5000/api/query"
End Function

' JSON 파싱을 위한 타입 정의
Type SourceDocument
    title As String
    organization As String
    docDate As String
    docType As String
    content As String
    relevance As Double
End Type

' API 응답 타입
Type RAGAPIResponse
    answer As String
    sources As Collection
    totalSources As Integer
    internalDocs As Integer
    externalDocs As Integer
    errorMsg As String
End Type

' 실제 RAG API 호출 함수
Function CallRAGAPI(question As String, Optional docType As String = "both") As RAGAPIResponse
    Dim xmlHttp As Object
    Dim jsonRequest As Object
    Dim jsonResponse As Object
    Dim response As RAGAPIResponse
    Dim startTime As Double
    
    On Error GoTo ErrorHandler
    
    ' HTTP 요청 객체 생성
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' JSON 요청 데이터 생성
    Set jsonRequest = CreateObject("Scripting.Dictionary")
    jsonRequest.Add "question", question
    jsonRequest.Add "doc_type", docType
    
    ' API 호출
    With xmlHttp
        .Open "POST", API_URL(), False
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "Accept", "application/json"
        .setTimeouts API_TIMEOUT, API_TIMEOUT, API_TIMEOUT, API_TIMEOUT
        
        ' JSON 변환 및 전송
        .send JsonConverter.ConvertToJson(jsonRequest)
        
        ' 응답 확인
        If .Status = 200 Then
            ' JSON 응답 파싱
            Set jsonResponse = JsonConverter.ParseJson(.responseText)
            
            ' 응답 데이터 추출
            response.answer = jsonResponse("answer")
            response.totalSources = jsonResponse("total_sources")
            response.internalDocs = jsonResponse("internal_docs")
            response.externalDocs = jsonResponse("external_docs")
            
            ' 소스 문서 파싱
            Set response.sources = New Collection
            Dim sourceItem As Variant
            Dim src As SourceDocument
            
            If jsonResponse.Exists("sources") Then
                For Each sourceItem In jsonResponse("sources")
                    src.title = sourceItem("title")
                    src.organization = sourceItem("organization")
                    src.docDate = sourceItem("date")
                    src.docType = sourceItem("type")
                    If sourceItem.Exists("content") Then
                        src.content = sourceItem("content")
                    End If
                    If sourceItem.Exists("relevance_score") Then
                        src.relevance = sourceItem("relevance_score")
                    End If
                    response.sources.Add src
                Next sourceItem
            End If
            
        Else
            response.errorMsg = "API 오류: " & .Status & " - " & .statusText
        End If
    End With
    
    CallRAGAPI = response
    Exit Function
    
ErrorHandler:
    response.errorMsg = "오류 발생: " & Err.Description
    CallRAGAPI = response
End Function

' Enhanced Sources와 통합된 RAG 검색 실행
Sub RunRAGSearchWithSources()
    Dim ws As Worksheet
    Dim question As String
    Dim apiResponse As RAGAPIResponse
    Dim answer As String
    
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
    apiResponse = CallRAGAPI(question)
    
    ' 오류 확인
    If apiResponse.errorMsg <> "" Then
        ' API 서버가 실행되지 않은 경우 시뮬레이션 모드로 전환
        If InStr(apiResponse.errorMsg, "개체가 필요합니다") > 0 Or _
           InStr(apiResponse.errorMsg, "연결") > 0 Then
            
            ws.Range("B64").Value = "⚠️ API 서버 미실행 - 시뮬레이션 모드"
            ws.Range("B64").Font.Color = RGB(255, 165, 0)
            
            ' 시뮬레이션 모드로 전환
            Call modEnhancedSources.RunSearchWithEnhancedSources
            Exit Sub
        Else
            MsgBox "API 호출 실패: " & apiResponse.errorMsg, vbCritical
            ws.Range("B64").Value = "❌ 오류 발생"
            ws.Range("B64").Font.Color = RGB(255, 0, 0)
            Exit Sub
        End If
    End If
    
    ' 답변 표시
    ws.Range("B10").Value = apiResponse.answer
    ws.Range("B10").Font.Color = RGB(0, 0, 0)
    
    ' 소스 문서 표시
    DisplayRAGSources ws, 24, apiResponse.sources
    
    ' 통계 정보 표시
    Dim statsMsg As String
    statsMsg = "✅ 검색 완료 - " & Format(Now, "hh:mm:ss") & _
               " | 참고문서: " & apiResponse.totalSources & "개" & _
               " (내부: " & apiResponse.internalDocs & ", 외부: " & apiResponse.externalDocs & ")"
    
    ws.Range("B64").Value = statsMsg
    ws.Range("B64").Font.Color = RGB(0, 150, 0)
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ws.Range("B64").Value = "❌ 오류 발생: " & Err.Description
    ws.Range("B64").Font.Color = RGB(255, 0, 0)
    Application.StatusBar = False
    
    ' 오류 발생시 시뮬레이션 모드 제안
    If MsgBox("API 연결 실패. 시뮬레이션 모드로 실행하시겠습니까?", vbYesNo + vbQuestion) = vbYes Then
        Call modEnhancedSources.RunSearchWithEnhancedSources
    End If
End Sub

' RAG API에서 받은 소스 문서 표시
Sub DisplayRAGSources(ws As Worksheet, startRow As Integer, sources As Collection)
    Dim currentRow As Integer
    Dim i As Integer
    Dim src As SourceDocument
    
    currentRow = startRow
    
    ' 헤더 스타일
    With ws.Range("B" & currentRow & ":F" & currentRow)
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
    End With
    
    ws.Cells(currentRow, 2).Value = "번호"
    ws.Cells(currentRow, 3).Value = "제목"
    ws.Cells(currentRow, 4).Value = "출처/조직"
    ws.Cells(currentRow, 5).Value = "날짜"
    ws.Cells(currentRow, 6).Value = "유형"
    
    currentRow = currentRow + 1
    
    ' 소스 데이터 표시
    If sources.Count > 0 Then
        For i = 1 To sources.Count
            If currentRow > startRow + 35 Then Exit For
            
            src = sources(i)
            
            ws.Cells(currentRow, 2).Value = "[" & i & "]"
            ws.Cells(currentRow, 2).Font.Bold = True
            ws.Cells(currentRow, 2).Font.Color = RGB(0, 112, 192)
            
            ws.Cells(currentRow, 3).Value = src.title
            ws.Cells(currentRow, 3).WrapText = True
            
            ws.Cells(currentRow, 4).Value = src.organization
            ws.Cells(currentRow, 5).Value = src.docDate
            
            ' 영어를 한글로 변환하여 표시
            Dim displayType As String
            Select Case src.docType
                Case "internal"
                    displayType = "사내"
                Case "external"
                    displayType = "사외"
                Case Else
                    displayType = src.docType
            End Select
            ws.Cells(currentRow, 6).Value = displayType
            
            ' 유형별 색상 코딩
            Select Case src.docType
                Case "internal", "내부", "사내"
                    ws.Cells(currentRow, 6).Interior.Color = RGB(255, 242, 204)
                Case "external", "외부", "사외"
                    ws.Cells(currentRow, 6).Interior.Color = RGB(217, 234, 211)
                Case "urgent", "긴급"
                    ws.Cells(currentRow, 6).Interior.Color = RGB(255, 199, 206)
            End Select
            
            ' 관련도 점수가 있으면 표시
            If src.relevance > 0 Then
                ws.Cells(currentRow, 3).Value = src.title & " (" & Format(src.relevance * 100, "0") & "%)"
            End If
            
            ' 행 서식
            With ws.Range("B" & currentRow & ":F" & currentRow)
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200)
                If i Mod 2 = 0 Then
                    .Interior.Color = RGB(248, 248, 248)
                End If
            End With
            
            ws.Rows(currentRow).RowHeight = 20
            currentRow = currentRow + 1
        Next i
    Else
        ' 소스가 없는 경우 Enhanced Sources 사용
        Call modEnhancedSources.DisplayEnhancedSources(ws, currentRow - 1)
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

' API 서버 시작 안내
Sub ShowAPIServerGuide()
    Dim msg As String
    
    msg = "RAG API 서버가 실행되지 않았습니다." & Chr(10) & Chr(10) & _
          "터미널에서 다음 명령을 실행하세요:" & Chr(10) & Chr(10) & _
          "cd C:\Users\admin\documents\github\strix" & Chr(10) & _
          "py api_server_with_sources.py" & Chr(10) & Chr(10) & _
          "서버가 실행되면 다시 검색해주세요."
    
    MsgBox msg, vbInformation, "API 서버 실행 필요"
End Sub