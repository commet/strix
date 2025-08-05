Attribute VB_Name = "modSTRIXExcel"
Option Explicit

' Excel에서 STRIX 사용을 위한 헬퍼 함수들

' 리본 메뉴 또는 버튼에서 호출할 메인 함수
Sub ShowSTRIXDialog()
    frmSTRIX.Show
End Sub

' 선택된 셀의 내용으로 질문하기
Sub AskAboutSelection()
    Dim selectedText As String
    Dim answer As String
    
    ' 선택된 텍스트 가져오기
    If TypeName(Selection) = "Range" Then
        selectedText = Selection.Value
    Else
        MsgBox "셀을 선택해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' STRIX에 질문
    answer = AskSTRIX("다음 내용을 분석해주세요: " & selectedText)
    
    ' 결과를 새 시트에 표시
    Call DisplayAnswer(answer, selectedText)
End Sub

' 답변을 새 시트에 표시
Sub DisplayAnswer(answer As String, Optional question As String = "")
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    
    ' 새 시트 생성
    Set newSheet = ThisWorkbook.Sheets.Add
    newSheet.Name = "STRIX_" & Format(Now, "yyyymmdd_hhmmss")
    
    ' 헤더 설정
    With newSheet
        .Range("A1").Value = "STRIX Intelligence Report"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "질문:"
        .Range("B3").Value = question
        .Range("A3:B3").Font.Bold = True
        
        .Range("A5").Value = "답변:"
        .Range("A5").Font.Bold = True
        
        ' 답변 내용 (줄바꿈 처리)
        Dim lines() As String
        lines = Split(answer, vbLf)
        Dim i As Integer
        For i = 0 To UBound(lines)
            .Range("A" & (6 + i)).Value = lines(i)
        Next i
        
        ' 서식 조정
        .Columns("A:B").AutoFit
        .Range("A:A").ColumnWidth = 80
    End With
End Sub

' 문서 일괄 업로드
Sub BulkUploadDocuments()
    Dim folderPath As String
    Dim fileName As String
    Dim docType As String
    Dim organization As String
    Dim successCount As Integer
    Dim failCount As Integer
    
    ' 폴더 선택
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "업로드할 문서가 있는 폴더를 선택하세요"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' 문서 타입 입력
    docType = InputBox("문서 타입을 입력하세요 (internal/external):", "문서 타입", "internal")
    organization = InputBox("조직명을 입력하세요:", "조직명", "전략기획")
    
    ' 폴더 내 파일 처리
    fileName = Dir(folderPath & "\*.txt")
    Do While fileName <> ""
        If UploadDocument(folderPath & "\" & fileName, docType, organization) Then
            successCount = successCount + 1
        Else
            failCount = failCount + 1
        End If
        fileName = Dir
    Loop
    
    MsgBox "업로드 완료!" & vbCrLf & _
           "성공: " & successCount & "개" & vbCrLf & _
           "실패: " & failCount & "개", vbInformation
End Sub

' 검색 결과를 시트에 표시
Sub DisplaySearchResults()
    Dim keyword As String
    Dim documents As Collection
    Dim ws As Worksheet
    Dim row As Integer
    
    ' 검색어 입력
    keyword = InputBox("검색할 키워드를 입력하세요:", "문서 검색")
    If keyword = "" Then Exit Sub
    
    ' 문서 검색
    Set documents = SearchDocuments(keyword)
    
    ' 결과 시트 생성
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "검색결과_" & Format(Now, "hhmmss")
    
    ' 헤더 생성
    With ws
        .Range("A1:E1").Value = Array("ID", "제목", "타입", "조직", "생성일")
        .Range("A1:E1").Font.Bold = True
        
        ' 데이터 표시
        row = 2
        Dim doc As Variant
        For Each doc In documents
            .Cells(row, 1).Value = doc("id")
            .Cells(row, 2).Value = doc("title")
            .Cells(row, 3).Value = doc("type")
            .Cells(row, 4).Value = doc("organization")
            .Cells(row, 5).Value = doc("created_at")
            row = row + 1
        Next doc
        
        ' 서식 조정
        .Columns("A:E").AutoFit
        .Range("A1:E" & row - 1).Borders.LineStyle = xlContinuous
    End With
End Sub

' 최근 검색 기록 표시
Sub ShowRecentSearches()
    Dim logs As Collection
    Dim ws As Worksheet
    Dim row As Integer
    
    ' 검색 로그 가져오기
    Set logs = GetSearchLogs(20)
    
    ' 결과 시트 생성
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "검색기록_" & Format(Now, "hhmmss")
    
    ' 헤더 생성
    With ws
        .Range("A1:C1").Value = Array("시간", "검색어", "결과수")
        .Range("A1:C1").Font.Bold = True
        
        ' 데이터 표시
        row = 2
        Dim log As Variant
        For Each log In logs
            .Cells(row, 1).Value = log("created_at")
            .Cells(row, 2).Value = log("query")
            .Cells(row, 3).Value = log("results")("internal_docs") + log("results")("external_docs")
            row = row + 1
        Next log
        
        ' 서식 조정
        .Columns("A:C").AutoFit
        .Range("A1:C" & row - 1).Borders.LineStyle = xlContinuous
    End With
End Sub

' 커스텀 함수: 셀에서 직접 STRIX 호출
Function STRIX(question As String) As String
    On Error GoTo ErrorHandler
    STRIX = AskSTRIX(question)
    Exit Function
ErrorHandler:
    STRIX = "Error: " & Err.Description
End Function