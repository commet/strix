Attribute VB_Name = "modExternalIngest"
Option Explicit

'==============================================================================
' External News Ingestion Module
' 외부 뉴스 메일 스캔 및 분류
'==============================================================================

' === 메인 스캔 함수 ===
Public Sub ScanExternalNews()
    On Error GoTo ErrorHandler
    
    ' 설정 검증
    If Not ValidateConfig() Then Exit Sub
    
    ' 잠금 획득
    If Not AcquireLock() Then
        MsgBox "다른 사용자가 스캔 중입니다.", vbExclamation
        Exit Sub
    End If
    
    Application.StatusBar = "외부 뉴스 스캔 시작..."
    
    ' Mock 환경에서는 파일 시스템 스캔
    ' 실제 환경에서는 Outlook 폴더 스캔으로 전환
    Dim scanCount As Long
    
    If IsOutlookAvailable() Then
        scanCount = ScanOutlookFolders()
    Else
        scanCount = ScanMockNewsFiles(gblExternalFolderPath)
    End If
    
    ' 자동 분류 실행
    If scanCount > 0 Then
        ClassifyNews
    End If
    
    ' 마지막 스캔 시간 업데이트
    gblLastExternalScan = Now
    SaveConfig
    
    ' 잠금 해제
    ReleaseLock
    
    Application.StatusBar = False
    MsgBox "외부 뉴스 스캔 완료!" & vbNewLine & _
           "처리된 뉴스: " & scanCount & "개", vbInformation, APP_NAME
    
    Exit Sub
ErrorHandler:
    ReleaseLock
    Application.StatusBar = False
    MsgBox "스캔 중 오류 발생: " & Err.Description, vbCritical
End Sub

' === Outlook 스캔 ===
Private Function ScanOutlookFolders() As Long
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim olNs As Object
    Dim olFolder As Object
    Dim olMail As Object
    Dim processedCount As Long
    
    ' Outlook 객체 생성
    Set olApp = CreateObject("Outlook.Application")
    Set olNs = olApp.GetNamespace("MAPI")
    
    ' 스캔할 폴더 목록
    Dim folderNames As Variant
    folderNames = Array("PR팀_AM", "PR팀_PM", "외부뉴스", "Google_Alert")
    
    Dim i As Long
    For i = 0 To UBound(folderNames)
        On Error Resume Next
        Set olFolder = olNs.Folders("받은편지함").Folders(folderNames(i))
        On Error GoTo ErrorHandler
        
        If Not olFolder Is Nothing Then
            processedCount = processedCount + ProcessOutlookFolder(olFolder)
        End If
    Next i
    
    ScanOutlookFolders = processedCount
    
    ' 정리
    Set olMail = Nothing
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
    
    Exit Function
ErrorHandler:
    ScanOutlookFolders = 0
End Function

Private Function ProcessOutlookFolder(olFolder As Object) As Long
    Dim olMail As Object
    Dim processedCount As Long
    Dim item As Object
    
    For Each item In olFolder.Items
        If TypeName(item) = "MailItem" Then
            Set olMail = item
            
            ' 마지막 스캔 이후 메일만 처리
            If olMail.ReceivedTime > gblLastExternalScan Then
                If ProcessMailItem(olMail) Then
                    processedCount = processedCount + 1
                End If
            End If
        End If
    Next item
    
    ProcessOutlookFolder = processedCount
End Function

Private Function ProcessMailItem(olMail As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_RAWNEWS)
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("RawNews_tbl")
    
    ' 중복 확인
    If IsMailProcessed(tbl, olMail.EntryID) Then
        ProcessMailItem = False
        Exit Function
    End If
    
    ' 새 행 추가
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    
    With newRow
        .Range(1, 1).Value = olMail.EntryID ' MailID
        .Range(1, 2).Value = olMail.ReceivedTime ' ReceivedDate
        .Range(1, 3).Value = olMail.Subject ' Subject
        .Range(1, 4).Value = olMail.SenderName ' Sender
        .Range(1, 5).Value = Left(olMail.Body, 5000) ' BodyText (5000자 제한)
        .Range(1, 6).Value = SaveAttachments(olMail) ' AttachmentPath
        .Range(1, 7).Value = "" ' Category (자동 분류 예정)
        .Range(1, 8).Value = "" ' SubCategory
        .Range(1, 9).Value = "N" ' ProcessedFlag
    End With
    
    ProcessMailItem = True
    
    Exit Function
ErrorHandler:
    Debug.Print "메일 처리 오류: " & Err.Description
    ProcessMailItem = False
End Function

' === Mock 파일 스캔 (Outlook 없는 환경) ===
Private Function ScanMockNewsFiles(folderPath As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        MsgBox "외부 뉴스 폴더를 찾을 수 없습니다: " & folderPath, vbExclamation
        ScanMockNewsFiles = 0
        Exit Function
    End If
    
    Dim processedCount As Long
    processedCount = ProcessNewsFolder(fso, fso.GetFolder(folderPath))
    
    ScanMockNewsFiles = processedCount
End Function

Private Function ProcessNewsFolder(fso As Object, folder As Object) As Long
    Dim file As Object
    Dim subFolder As Object
    Dim count As Long
    
    ' 현재 폴더의 파일 처리
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".txt" Or _
           LCase(Right(file.Name, 4)) = ".msg" Then
            If ProcessNewsFile(file) Then
                count = count + 1
            End If
        End If
    Next file
    
    ' 하위 폴더 처리
    For Each subFolder In folder.SubFolders
        count = count + ProcessNewsFolder(fso, subFolder)
    Next subFolder
    
    ProcessNewsFolder = count
End Function

Private Function ProcessNewsFile(file As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_RAWNEWS)
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("RawNews_tbl")
    
    ' 중복 확인
    If IsFileInTable(tbl, file.Path) Then
        ProcessNewsFile = False
        Exit Function
    End If
    
    ' 파일 내용 읽기
    Dim newsData As Object
    Set newsData = ParseNewsFile(file)
    
    ' 테이블에 추가
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    
    With newRow
        .Range(1, 1).Value = "NEWS-" & Format(Now, "yyyymmddhhmmss") & "-" & Format(Int(Rnd * 9999), "0000") ' MailID
        .Range(1, 2).Value = newsData("Date") ' ReceivedDate
        .Range(1, 3).Value = newsData("Subject") ' Subject
        .Range(1, 4).Value = newsData("Sender") ' Sender
        .Range(1, 5).Value = newsData("Body") ' BodyText
        .Range(1, 6).Value = file.Path ' AttachmentPath
        .Range(1, 7).Value = newsData("Category") ' Category
        .Range(1, 8).Value = "" ' SubCategory
        .Range(1, 9).Value = "N" ' ProcessedFlag
    End With
    
    ProcessNewsFile = True
    
    Exit Function
ErrorHandler:
    Debug.Print "뉴스 파일 처리 오류: " & file.Name & " - " & Err.Description
    ProcessNewsFile = False
End Function

' === 뉴스 파일 파싱 ===
Private Function ParseNewsFile(file As Object) As Object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(file.Path, 1) ' ForReading
    
    Dim newsData As Object
    Set newsData = CreateObject("Scripting.Dictionary")
    
    ' 기본값 설정
    newsData.Add "Date", file.DateCreated
    newsData.Add "Subject", file.Name
    newsData.Add "Sender", "Unknown"
    newsData.Add "Body", ""
    newsData.Add "Category", ""
    
    ' 파일 내용 파싱
    Dim line As String
    Dim bodyStarted As Boolean
    Dim bodyText As String
    
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        
        If InStr(line, "From:") = 1 Then
            newsData("Sender") = Trim(Mid(line, 6))
        ElseIf InStr(line, "Date:") = 1 Then
            On Error Resume Next
            newsData("Date") = CDate(Trim(Mid(line, 6)))
            On Error GoTo 0
        ElseIf InStr(line, "Subject:") = 1 Then
            newsData("Subject") = Trim(Mid(line, 9))
        ElseIf InStr(line, "Category:") = 1 Then
            newsData("Category") = Trim(Mid(line, 10))
        ElseIf line = "" And Not bodyStarted Then
            bodyStarted = True
        ElseIf bodyStarted Then
            bodyText = bodyText & line & vbNewLine
        End If
    Loop
    
    newsData("Body") = Left(bodyText, 5000) ' 5000자 제한
    
    ts.Close
    
    Set ParseNewsFile = newsData
End Function

' === 자동 분류 ===
Private Sub ClassifyNews()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_RAWNEWS)
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("RawNews_tbl")
    
    Dim row As ListRow
    Dim category As String
    Dim subCategory As String
    
    For Each row In tbl.ListRows
        If row.Range(1, 7).Value = "" Then ' Category가 비어있으면
            ' 제목과 본문에서 카테고리 추출
            category = ExtractCategory(row.Range(1, 3).Value & " " & row.Range(1, 5).Value)
            row.Range(1, 7).Value = category
            
            ' 세부 카테고리 추출
            subCategory = ExtractSubCategory(category, row.Range(1, 3).Value & " " & row.Range(1, 5).Value)
            row.Range(1, 8).Value = subCategory
        End If
    Next row
    
    Exit Sub
ErrorHandler:
    Debug.Print "자동 분류 오류: " & Err.Description
End Sub

Private Function ExtractCategory(text As String) As String
    ' Config 시트에서 키워드 가져오기
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    Dim categories As Variant
    categories = Array("Macro", "산업", "기술", "리스크", "경쟁사", "정책")
    
    Dim i As Long
    Dim keywords As String
    Dim keywordArr As Variant
    Dim j As Long
    
    For i = 0 To UBound(categories)
        ' F2:G10 범위에서 키워드 찾기
        keywords = ws.Range("G" & (i + 2)).Value
        
        If keywords <> "" Then
            keywordArr = Split(keywords, ",")
            
            For j = 0 To UBound(keywordArr)
                If InStr(1, text, Trim(keywordArr(j)), vbTextCompare) > 0 Then
                    ExtractCategory = categories(i)
                    Exit Function
                End If
            Next j
        End If
    Next i
    
    ' 기본값
    ExtractCategory = "기타"
End Function

Private Function ExtractSubCategory(category As String, text As String) As String
    Select Case category
        Case "Macro"
            If InStr(text, "금리") > 0 Then ExtractSubCategory = "금리"
            If InStr(text, "환율") > 0 Then ExtractSubCategory = "환율"
            If InStr(text, "인플레이션") > 0 Then ExtractSubCategory = "인플레이션"
            
        Case "산업"
            If InStr(text, "배터리") > 0 Then ExtractSubCategory = "배터리"
            If InStr(text, "전기차") > 0 Then ExtractSubCategory = "전기차"
            If InStr(text, "반도체") > 0 Then ExtractSubCategory = "반도체"
            
        Case "기술"
            If InStr(text, "전고체") > 0 Then ExtractSubCategory = "전고체"
            If InStr(text, "AI") > 0 Then ExtractSubCategory = "AI"
            If InStr(text, "자동화") > 0 Then ExtractSubCategory = "자동화"
            
        Case "리스크"
            If InStr(text, "규제") > 0 Then ExtractSubCategory = "규제"
            If InStr(text, "사고") > 0 Then ExtractSubCategory = "사고"
            If InStr(text, "리콜") > 0 Then ExtractSubCategory = "리콜"
            
        Case "경쟁사"
            If InStr(text, "CATL") > 0 Then ExtractSubCategory = "CATL"
            If InStr(text, "BYD") > 0 Then ExtractSubCategory = "BYD"
            If InStr(text, "Tesla") > 0 Then ExtractSubCategory = "Tesla"
            
        Case "정책"
            If InStr(text, "IRA") > 0 Then ExtractSubCategory = "IRA"
            If InStr(text, "CBAM") > 0 Then ExtractSubCategory = "CBAM"
            If InStr(text, "탄소중립") > 0 Then ExtractSubCategory = "탄소중립"
    End Select
    
    If ExtractSubCategory = "" Then ExtractSubCategory = "일반"
End Function

' === 유틸리티 함수 ===
Private Function IsOutlookAvailable() As Boolean
    On Error Resume Next
    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
    IsOutlookAvailable = Not olApp Is Nothing
    Set olApp = Nothing
End Function

Private Function IsMailProcessed(tbl As ListObject, mailID As String) As Boolean
    Dim row As ListRow
    For Each row In tbl.ListRows
        If row.Range(1, 1).Value = mailID Then
            IsMailProcessed = True
            Exit Function
        End If
    Next row
    IsMailProcessed = False
End Function

Private Function IsFileInTable(tbl As ListObject, filePath As String) As Boolean
    Dim row As ListRow
    For Each row In tbl.ListRows
        If row.Range(1, 6).Value = filePath Then
            IsFileInTable = True
            Exit Function
        End If
    Next row
    IsFileInTable = False
End Function

Private Function SaveAttachments(olMail As Object) As String
    ' 첨부파일 저장 (실제 구현 시)
    ' 현재는 첨부파일 존재 여부만 표시
    If olMail.Attachments.Count > 0 Then
        SaveAttachments = "첨부파일 " & olMail.Attachments.Count & "개"
    Else
        SaveAttachments = ""
    End If
End Function