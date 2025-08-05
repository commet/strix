Attribute VB_Name = "modInternalIngest"
Option Explicit

'==============================================================================
' Internal Document Ingestion Module
' 내부 보고자료(Word, PPT, PDF) 스캔 및 메타데이터 추출
'==============================================================================

' === 메인 스캔 함수 ===
Public Sub ScanInternalDocuments()
    On Error GoTo ErrorHandler
    
    ' 설정 로드
    If Not ValidateConfig() Then Exit Sub
    
    ' 잠금 획득
    If Not AcquireLock() Then
        MsgBox "다른 사용자가 스캔 중입니다. 잠시 후 다시 시도하세요.", vbExclamation
        Exit Sub
    End If
    
    Application.StatusBar = "내부 문서 스캔 시작..."
    
    ' 스캔 실행
    Dim scanCount As Long
    scanCount = PerformInternalScan(gblInternalFolderPath)
    
    ' 마지막 스캔 시간 업데이트
    gblLastInternalScan = Now
    SaveConfig
    
    ' 잠금 해제
    ReleaseLock
    
    Application.StatusBar = False
    MsgBox "내부 문서 스캔 완료!" & vbNewLine & _
           "처리된 파일: " & scanCount & "개", vbInformation, APP_NAME
    
    Exit Sub
ErrorHandler:
    ReleaseLock
    Application.StatusBar = False
    MsgBox "스캔 중 오류 발생: " & Err.Description, vbCritical
End Sub

' === 증분 스캔 함수 ===
Public Sub IncrementalScanInternal()
    On Error GoTo ErrorHandler
    
    If Not ValidateConfig() Then Exit Sub
    
    Application.StatusBar = "증분 스캔 중..."
    
    ' 마지막 스캔 이후 변경된 파일만 처리
    Dim scanCount As Long
    scanCount = PerformIncrementalScan(gblInternalFolderPath, gblLastInternalScan)
    
    gblLastInternalScan = Now
    SaveConfig
    
    Application.StatusBar = False
    
    If scanCount > 0 Then
        MsgBox "증분 스캔 완료! 새로운 파일: " & scanCount & "개", vbInformation
    End If
    
    Exit Sub
ErrorHandler:
    Application.StatusBar = False
    Debug.Print "증분 스캔 오류: " & Err.Description
End Sub

' === 실제 스캔 수행 ===
Private Function PerformInternalScan(folderPath As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim processedCount As Long
    processedCount = 0
    
    ' 재귀적으로 폴더 스캔
    ProcessFolder fso, fso.GetFolder(folderPath), processedCount
    
    PerformInternalScan = processedCount
End Function

Private Sub ProcessFolder(fso As Object, folder As Object, ByRef count As Long)
    On Error Resume Next
    
    Dim file As Object
    Dim subFolder As Object
    
    ' 현재 폴더의 파일 처리
    For Each file In folder.Files
        If IsValidDocumentFile(file.Name) Then
            If ProcessInternalFile(file) Then
                count = count + 1
                
                ' 진행 상황 표시
                If count Mod 10 = 0 Then
                    Application.StatusBar = "처리 중... " & count & "개 파일 완료"
                    DoEvents
                End If
            End If
        End If
    Next file
    
    ' 하위 폴더 재귀 처리
    For Each subFolder In folder.SubFolders
        ProcessFolder fso, subFolder, count
    Next subFolder
End Sub

' === 파일 처리 ===
Private Function ProcessInternalFile(file As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_RAWDATA)
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("RawData_tbl")
    
    ' 중복 확인
    If IsFileProcessed(tbl, file.Path) Then
        ProcessInternalFile = False
        Exit Function
    End If
    
    ' 메타데이터 추출
    Dim metadata As Object
    Set metadata = ExtractFileMetadata(file)
    
    ' 테이블에 추가
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    
    With newRow
        .Range(1, 1).Value = GenerateFileID() ' FileID
        .Range(1, 2).Value = file.Name ' FileName
        .Range(1, 3).Value = file.Path ' FilePath
        .Range(1, 4).Value = GetFileType(file.Name) ' FileType
        .Range(1, 5).Value = file.Size ' FileSize
        .Range(1, 6).Value = file.DateCreated ' CreatedDate
        .Range(1, 7).Value = file.DateLastModified ' ModifiedDate
        .Range(1, 8).Value = Now ' UploadDate
        .Range(1, 9).Value = metadata("Organization") ' Organization
        .Range(1, 10).Value = "" ' IssueID (나중에 매핑)
        .Range(1, 11).Value = "N" ' ProcessedFlag
    End With
    
    ProcessInternalFile = True
    
    Exit Function
ErrorHandler:
    Debug.Print "파일 처리 오류: " & file.Name & " - " & Err.Description
    ProcessInternalFile = False
End Function

' === 메타데이터 추출 ===
Private Function ExtractFileMetadata(file As Object) As Object
    Dim metadata As Object
    Set metadata = CreateObject("Scripting.Dictionary")
    
    ' 파일명에서 정보 추출
    Dim fileName As String
    fileName = file.Name
    
    ' 조직 추출 (파일 경로에서)
    Dim pathParts As Variant
    pathParts = Split(file.Path, "\")
    
    Dim org As String
    org = "기타"
    
    ' 폴더명에서 조직 찾기
    Dim i As Long
    For i = UBound(pathParts) - 1 To 0 Step -1
        If IsOrganization(pathParts(i)) Then
            org = pathParts(i)
            Exit For
        End If
    Next i
    
    metadata.Add "Organization", org
    
    ' 날짜 추출 (파일명에서)
    Dim datePattern As String
    datePattern = ExtractDateFromFileName(fileName)
    If datePattern <> "" Then
        metadata.Add "ReportDate", datePattern
    End If
    
    ' 문서 유형 추출
    Dim docType As String
    Select Case LCase(Right(fileName, 4))
        Case "docx", ".doc": docType = "Word"
        Case "pptx", ".ppt": docType = "PPT"
        Case ".pdf": docType = "PDF"
        Case ".xls", "xlsx": docType = "Excel"
        Case Else: docType = "기타"
    End Select
    
    metadata.Add "DocumentType", docType
    
    Set ExtractFileMetadata = metadata
End Function

' === 유틸리티 함수 ===
Private Function IsValidDocumentFile(fileName As String) As Boolean
    Dim validExtensions As Variant
    validExtensions = Array(".doc", ".docx", ".ppt", ".pptx", ".pdf", ".xls", ".xlsx")
    
    Dim ext As String
    ext = LCase(Right(fileName, InStrRev(fileName, ".") - 1))
    
    Dim i As Long
    For i = 0 To UBound(validExtensions)
        If ext = Mid(validExtensions(i), 2) Then
            IsValidDocumentFile = True
            Exit Function
        End If
    Next i
    
    IsValidDocumentFile = False
End Function

Private Function IsFileProcessed(tbl As ListObject, filePath As String) As Boolean
    On Error Resume Next
    
    Dim row As ListRow
    For Each row In tbl.ListRows
        If row.Range(1, 3).Value = filePath Then
            IsFileProcessed = True
            Exit Function
        End If
    Next row
    
    IsFileProcessed = False
End Function

Private Function GenerateFileID() As String
    ' 고유 ID 생성 (날짜시간 + 랜덤)
    GenerateFileID = "FID-" & Format(Now, "yyyymmddhhmmss") & "-" & _
                     Format(Int(Rnd * 9999), "0000")
End Function

Private Function GetFileType(fileName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    
    If dotPos > 0 Then
        GetFileType = UCase(Mid(fileName, dotPos + 1))
    Else
        GetFileType = "UNKNOWN"
    End If
End Function

Private Function IsOrganization(folderName As String) As Boolean
    Dim orgs As Variant
    orgs = Array("전략기획", "R&D", "경영지원", "생산", "영업마케팅", _
                 "재무", "인사", "구매", "품질", "IT")
    
    Dim i As Long
    For i = 0 To UBound(orgs)
        If InStr(folderName, orgs(i)) > 0 Then
            IsOrganization = True
            Exit Function
        End If
    Next i
    
    IsOrganization = False
End Function

Private Function ExtractDateFromFileName(fileName As String) As String
    ' 파일명에서 날짜 패턴 찾기
    ' 예: 2024_01_15, 2024-01-15, 20240115
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    regEx.Pattern = "(\d{4})[_-]?(\d{1,2})[_-]?(\d{1,2})"
    regEx.Global = True
    
    Dim matches As Object
    Set matches = regEx.Execute(fileName)
    
    If matches.Count > 0 Then
        ExtractDateFromFileName = matches(0).Value
    Else
        ExtractDateFromFileName = ""
    End If
End Function

' === 증분 스캔 ===
Private Function PerformIncrementalScan(folderPath As String, lastScanDate As Date) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim newFileCount As Long
    newFileCount = 0
    
    CheckNewFiles fso, fso.GetFolder(folderPath), lastScanDate, newFileCount
    
    PerformIncrementalScan = newFileCount
End Function

Private Sub CheckNewFiles(fso As Object, folder As Object, lastScanDate As Date, ByRef count As Long)
    On Error Resume Next
    
    Dim file As Object
    Dim subFolder As Object
    
    ' 새 파일 확인
    For Each file In folder.Files
        If file.DateLastModified > lastScanDate Then
            If IsValidDocumentFile(file.Name) Then
                If ProcessInternalFile(file) Then
                    count = count + 1
                End If
            End If
        End If
    Next file
    
    ' 하위 폴더 확인
    For Each subFolder In folder.SubFolders
        CheckNewFiles fso, subFolder, lastScanDate, count
    Next subFolder
End Sub