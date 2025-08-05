Attribute VB_Name = "modConfig"
Option Explicit

'==============================================================================
' STRIX Configuration Module
' 프로젝트 전역 설정 및 상수 관리
'==============================================================================

' === 전역 상수 ===
Public Const APP_NAME As String = "STRIX"
Public Const APP_VERSION As String = "1.0.0"
Public Const DEFAULT_SCAN_INTERVAL As Long = 60 ' 분 단위

' 파일 타입 필터
Public Const FILE_FILTER_INTERNAL As String = "*.ppt*;*.xls*;*.pdf"
Public Const FILE_FILTER_EXTERNAL As String = "*.msg;*.txt"

' 연관도 임계값
Public Const CORRELATION_THRESHOLD As Double = 0.2

' 시트 이름 상수
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_RAWDATA As String = "RawData"
Public Const SHEET_RAWNEWS As String = "RawNews"
Public Const SHEET_METADATA As String = "MetaData"
Public Const SHEET_LINKEDNEWS As String = "LinkedNews"
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_GPT As String = "GPT_Interface"
Public Const SHEET_NEWSLETTER As String = "NewsletterTemplate"
Public Const SHEET_REPORTS As String = "Reports"

' === 전역 변수 ===
Public gblInternalFolderPath As String
Public gblExternalFolderPath As String
Public gblLastInternalScan As Date
Public gblLastExternalScan As Date
Public gblIsLocked As Boolean
Public gblCurrentUser As String

' === 설정 로드/저장 ===
Public Sub LoadConfig()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    ' 기본 경로 설정
    gblInternalFolderPath = ws.Range("B2").Value
    gblExternalFolderPath = ws.Range("B3").Value
    
    ' 마지막 스캔 시간
    If IsDate(ws.Range("B4").Value) Then
        gblLastInternalScan = ws.Range("B4").Value
    Else
        gblLastInternalScan = DateAdd("d", -1, Now)
    End If
    
    If IsDate(ws.Range("B5").Value) Then
        gblLastExternalScan = ws.Range("B5").Value
    Else
        gblLastExternalScan = DateAdd("d", -1, Now)
    End If
    
    ' 사용자 정보
    gblCurrentUser = Environ("USERNAME")
    
    Exit Sub
ErrorHandler:
    MsgBox "설정 로드 중 오류 발생: " & Err.Description, vbCritical, APP_NAME
End Sub

Public Sub SaveConfig()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    ws.Range("B2").Value = gblInternalFolderPath
    ws.Range("B3").Value = gblExternalFolderPath
    ws.Range("B4").Value = gblLastInternalScan
    ws.Range("B5").Value = gblLastExternalScan
    ws.Range("B6").Value = gblCurrentUser
    ws.Range("B7").Value = Now
    
    Exit Sub
ErrorHandler:
    MsgBox "설정 저장 중 오류 발생: " & Err.Description, vbCritical, APP_NAME
End Sub

' === 잠금 관리 ===
Public Function AcquireLock() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    ' 현재 잠금 상태 확인
    If ws.Range("D2").Value = "LOCKED" Then
        ' 잠금 시간 확인 (30분 이상 경과 시 강제 해제)
        If DateDiff("n", ws.Range("D3").Value, Now) > 30 Then
            ReleaseLock
        Else
            MsgBox "다른 사용자가 사용 중입니다: " & ws.Range("D4").Value, vbExclamation
            AcquireLock = False
            Exit Function
        End If
    End If
    
    ' 잠금 설정
    ws.Range("D2").Value = "LOCKED"
    ws.Range("D3").Value = Now
    ws.Range("D4").Value = gblCurrentUser
    gblIsLocked = True
    
    AcquireLock = True
    
    Exit Function
ErrorHandler:
    AcquireLock = False
End Function

Public Sub ReleaseLock()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    ws.Range("D2").Value = "UNLOCKED"
    ws.Range("D3").Value = ""
    ws.Range("D4").Value = ""
    gblIsLocked = False
End Sub

' === 유틸리티 함수 ===
Public Function GetConfigValue(key As String) As Variant
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    ' Config 시트에서 Named Range로 값 가져오기
    GetConfigValue = ws.Range(key).Value
End Function

Public Sub SetConfigValue(key As String, value As Variant)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    ws.Range(key).Value = value
End Sub

' === 초기화 검증 ===
Public Function ValidateConfig() As Boolean
    Dim result As Boolean
    result = True
    
    ' 필수 폴더 경로 확인
    If Len(gblInternalFolderPath) = 0 Then
        MsgBox "내부 문서 폴더 경로가 설정되지 않았습니다.", vbExclamation
        result = False
    ElseIf Not FolderExists(gblInternalFolderPath) Then
        MsgBox "내부 문서 폴더를 찾을 수 없습니다: " & gblInternalFolderPath, vbExclamation
        result = False
    End If
    
    If Len(gblExternalFolderPath) = 0 Then
        MsgBox "외부 뉴스 폴더 경로가 설정되지 않았습니다.", vbExclamation
        result = False
    ElseIf Not FolderExists(gblExternalFolderPath) Then
        MsgBox "외부 뉴스 폴더를 찾을 수 없습니다: " & gblExternalFolderPath, vbExclamation
        result = False
    End If
    
    ValidateConfig = result
End Function

Private Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
End Function