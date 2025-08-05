VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSTRIX 
   Caption         =   "STRIX Intelligence System"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10680
   OleObjectBlob   =   "frmSTRIX.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSTRIX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    ' 콤보박스 초기화
    cboDocType.AddItem "전체"
    cboDocType.AddItem "내부 문서만"
    cboDocType.AddItem "외부 뉴스만"
    cboDocType.ListIndex = 0
    
    ' 텍스트 박스 초기화
    txtQuestion.Value = ""
    txtAnswer.Value = ""
    
    ' 레이블 설정
    lblStatus.Caption = "준비됨"
End Sub

Private Sub cmdAsk_Click()
    Dim question As String
    Dim docType As String
    Dim answer As String
    
    ' 입력 확인
    question = Trim(txtQuestion.Value)
    If question = "" Then
        MsgBox "질문을 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' 문서 타입 설정
    Select Case cboDocType.ListIndex
        Case 0: docType = "both"
        Case 1: docType = "internal"
        Case 2: docType = "external"
    End Select
    
    ' 상태 표시
    lblStatus.Caption = "답변 생성 중..."
    cmdAsk.Enabled = False
    DoEvents
    
    ' STRIX 호출
    On Error GoTo ErrorHandler
    answer = AskSTRIX(question, docType)
    
    ' 답변 표시
    txtAnswer.Value = answer
    
    ' 히스토리에 추가
    lstHistory.AddItem Now & " - " & Left(question, 50) & "..."
    
    lblStatus.Caption = "완료"
    cmdAsk.Enabled = True
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "오류 발생"
    cmdAsk.Enabled = True
    MsgBox "오류가 발생했습니다: " & Err.Description, vbCritical
End Sub

Private Sub cmdClear_Click()
    txtQuestion.Value = ""
    txtAnswer.Value = ""
    lblStatus.Caption = "준비됨"
End Sub

Private Sub cmdExport_Click()
    Dim ws As Worksheet
    
    If txtAnswer.Value = "" Then
        MsgBox "내보낼 답변이 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 새 시트에 결과 내보내기
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "STRIX_Export_" & Format(Now, "hhmmss")
    
    With ws
        .Range("A1").Value = "STRIX Intelligence Report"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "질문:"
        .Range("B3").Value = txtQuestion.Value
        
        .Range("A5").Value = "문서 타입:"
        .Range("B5").Value = cboDocType.Text
        
        .Range("A7").Value = "답변:"
        .Range("A8").Value = txtAnswer.Value
        
        .Columns("A:B").AutoFit
    End With
    
    MsgBox "결과가 새 시트에 저장되었습니다.", vbInformation
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lstHistory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' 히스토리 항목 더블클릭 시 상세 보기
    If lstHistory.ListIndex >= 0 Then
        MsgBox lstHistory.List(lstHistory.ListIndex), vbInformation, "검색 기록"
    End If
End Sub