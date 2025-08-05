Attribute VB_Name = "modSTRIXAPI"
Option Explicit

' STRIX API 연동 모듈
' Streamlit 앱의 API 엔드포인트와 통신

Private Const STRIX_API_URL As String = "http://localhost:8501/"
Private Const SUPABASE_URL As String = "https://qxrwyfxwwihskktsmjhj.supabase.co"
Private Const SUPABASE_ANON_KEY As String = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InF4cnd5Znh3d2loc2trdHNtamhqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTQyODY2MTEsImV4cCI6MjA2OTg2MjYxMX0.RYFV2PFIk6i-Se9Y3MfFbfR8Yz7R9_PzGeGC0F3IIqA"

' HTTP 요청을 위한 함수
Function SendHTTPRequest(url As String, method As String, Optional body As String = "", Optional headers As Object = Nothing) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    On Error GoTo ErrorHandler
    
    http.Open method, url, False
    
    ' 기본 헤더 설정
    http.setRequestHeader "Content-Type", "application/json"
    
    ' 추가 헤더 설정
    If Not headers Is Nothing Then
        Dim key As Variant
        For Each key In headers.Keys
            http.setRequestHeader CStr(key), headers(key)
        Next key
    End If
    
    ' 요청 전송
    If body = "" Then
        http.send
    Else
        http.send body
    End If
    
    ' 응답 반환
    SendHTTPRequest = http.responseText
    Exit Function
    
ErrorHandler:
    SendHTTPRequest = "Error: " & Err.Description
End Function

' STRIX에 질문하기 (임시 테스트 버전)
Function AskSTRIX(question As String, Optional docType As String = "both") As String
    On Error GoTo ErrorHandler
    
    ' 테스트를 위해 간단한 응답 반환
    AskSTRIX = "STRIX 시스템이 준비 중입니다. 질문: " & question
    Exit Function
    
ErrorHandler:
    AskSTRIX = "Error: " & Err.Description
End Function

' 원본 함수 (나중에 사용)
Function AskSTRIX_Original(question As String, Optional docType As String = "both") As String
    Dim url As String
    Dim response As String
    Dim jsonResponse As Object
    
    ' URL 인코딩
    question = Application.WorksheetFunction.EncodeURL(question)
    
    ' URL 생성
    url = STRIX_API_URL & "?api=true&question=" & question & "&doc_type=" & docType
    
    ' API 호출
    response = SendHTTPRequest(url, "GET")
    
    ' 디버그를 위해 응답 확인
    Debug.Print "Response: " & Left(response, 200)
    
    ' JSON 파싱
    Set jsonResponse = JsonConverter.ParseJson(response)
    
    If jsonResponse.Exists("answer") Then
        AskSTRIX_Original = jsonResponse("answer")
    Else
        AskSTRIX_Original = "Error: No answer received"
    End If
End Function

' Supabase에서 직접 문서 검색
Function SearchDocuments(keyword As String, Optional limit As Integer = 10) As Collection
    Dim url As String
    Dim headers As Object
    Dim response As String
    Dim jsonResponse As Object
    Dim documents As New Collection
    
    ' URL 생성
    url = SUPABASE_URL & "/rest/v1/documents?title=ilike.*" & keyword & "*&limit=" & limit
    
    ' 헤더 설정
    Set headers = CreateObject("Scripting.Dictionary")
    headers.Add "apikey", SUPABASE_ANON_KEY
    headers.Add "Authorization", "Bearer " & SUPABASE_ANON_KEY
    
    ' API 호출
    response = SendHTTPRequest(url, "GET", "", headers)
    
    ' JSON 파싱
    Set jsonResponse = JsonConverter.ParseJson(response)
    
    ' 결과를 컬렉션으로 변환
    Dim doc As Variant
    For Each doc In jsonResponse
        documents.Add doc
    Next doc
    
    Set SearchDocuments = documents
End Function

' 문서 업로드 (Supabase에 직접)
Function UploadDocument(filePath As String, docType As String, organization As String) As Boolean
    Dim fso As Object
    Dim textStream As Object
    Dim content As String
    Dim jsonBody As String
    Dim headers As Object
    Dim response As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 파일 읽기
    Set textStream = fso.OpenTextFile(filePath, 1, False, -2) ' UTF-8
    content = textStream.ReadAll
    textStream.Close
    
    ' JSON 본문 생성
    jsonBody = "{"
    jsonBody = jsonBody & """type"":""" & docType & ""","
    jsonBody = jsonBody & """title"":""" & fso.GetFileName(filePath) & ""","
    jsonBody = jsonBody & """organization"":""" & organization & ""","
    jsonBody = jsonBody & """file_path"":""" & filePath & ""","
    jsonBody = jsonBody & """source"":""VBA Upload"","
    jsonBody = jsonBody & """metadata"":{""uploaded_from"":""Excel VBA""}"
    jsonBody = jsonBody & "}"
    
    ' 헤더 설정
    Set headers = CreateObject("Scripting.Dictionary")
    headers.Add "apikey", SUPABASE_ANON_KEY
    headers.Add "Authorization", "Bearer " & SUPABASE_ANON_KEY
    headers.Add "Prefer", "return=representation"
    
    ' API 호출
    response = SendHTTPRequest(SUPABASE_URL & "/rest/v1/documents", "POST", jsonBody, headers)
    
    ' 성공 여부 확인
    UploadDocument = (InStr(response, "id") > 0)
End Function

' 검색 로그 조회
Function GetSearchLogs(Optional limit As Integer = 10) As Collection
    Dim url As String
    Dim headers As Object
    Dim response As String
    Dim jsonResponse As Object
    Dim logs As New Collection
    
    ' URL 생성
    url = SUPABASE_URL & "/rest/v1/search_logs?order=created_at.desc&limit=" & limit
    
    ' 헤더 설정
    Set headers = CreateObject("Scripting.Dictionary")
    headers.Add "apikey", SUPABASE_ANON_KEY
    headers.Add "Authorization", "Bearer " & SUPABASE_ANON_KEY
    
    ' API 호출
    response = SendHTTPRequest(url, "GET", "", headers)
    
    ' JSON 파싱
    Set jsonResponse = JsonConverter.ParseJson(response)
    
    ' 결과를 컬렉션으로 변환
    Dim log As Variant
    For Each log In jsonResponse
        logs.Add log
    Next log
    
    Set GetSearchLogs = logs
End Function