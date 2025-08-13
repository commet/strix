VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmailCompose 
   Caption         =   "이메일 작성 및 전송"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   OleObjectBlob   =   "frmEmailCompose.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEmailCompose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    ' 수신자 불러오기 (설정에서)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    If Not ws Is Nothing Then
        If ws.Range("B4").Value <> "" Then
            txtTo.Value = ws.Range("B4").Value
        Else
            txtTo.Value = "ceo@company.com; coo@company.com"
        End If
    Else
        txtTo.Value = "ceo@company.com; coo@company.com"
    End If
    
    ' CC 기본값
    txtCC.Value = "risk-management@company.com"
    
    ' 제목 자동 생성
    txtSubject.Value = "[STRIX Alert] " & Format(Date, "yyyy-mm-dd") & " Critical Issues Report"
    
    ' 본문 자동 생성
    Call GenerateEmailBody
    
    ' 우선순위 설정
    cboPriority.AddItem "높음"
    cboPriority.AddItem "보통"
    cboPriority.AddItem "낮음"
    cboPriority.Value = "높음"
End Sub

Private Sub GenerateEmailBody()
    Dim body As String
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Smart Alerts")
    
    body = "안녕하세요," & vbLf & vbLf
    body = body & "STRIX Smart Alert System에서 발송하는 " & Format(Date, "yyyy년 mm월 dd일") & " Critical Issues 보고서입니다." & vbLf & vbLf
    
    body = body & "========================================" & vbLf
    body = body & "TOP 5 CRITICAL ISSUES" & vbLf
    body = body & "========================================" & vbLf & vbLf
    
    ' Smart Alerts 시트에서 데이터 가져오기
    If Not ws Is Nothing Then
        Dim i As Integer
        For i = 13 To 17
            If ws.Cells(i, 3).Value <> "" Then
                body = body & ws.Cells(i, 2).Value & ". " & ws.Cells(i, 3).Value & vbLf
                body = body & "   위험도: " & ws.Cells(i, 4).Value & vbLf
                body = body & "   권장 액션: " & ws.Cells(i, 6).Value & vbLf
                body = body & "   담당: " & ws.Cells(i, 7).Value & vbLf
                body = body & "   구분: " & ws.Cells(i, 8).Value & vbLf & vbLf
            End If
        Next i
    Else
        body = body & "1. SK온-SK엔무브 합병 통합법인 출범 준비" & vbLf
        body = body & "   위험도: 92%" & vbLf
        body = body & "   권장 액션: 통합 실무 TF 구성" & vbLf & vbLf
        
        body = body & "2. 트럼프 IRA 폐지 가능성, AMPC 세액공제 위기" & vbLf
        body = body & "   위험도: 90%" & vbLf
        body = body & "   권장 액션: 정책 대응 시나리오 수립" & vbLf & vbLf
    End If
    
    body = body & "========================================" & vbLf
    body = body & "ACTION REQUIRED" & vbLf
    body = body & "========================================" & vbLf & vbLf
    body = body & "위 이슈들에 대한 즉각적인 검토와 대응이 필요합니다." & vbLf
    body = body & "상세 내용은 첨부된 보고서를 참조하시기 바랍니다." & vbLf & vbLf
    
    body = body & "감사합니다." & vbLf & vbLf
    body = body & "STRIX Alert System" & vbLf
    body = body & "자동 생성 시간: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    txtBody.Value = body
End Sub

Private Sub btnAddRecipient_Click()
    ' 수신자 추가
    Dim newRecipient As String
    newRecipient = InputBox("추가할 이메일 주소를 입력하세요:", "수신자 추가")
    If newRecipient <> "" Then
        If txtTo.Value = "" Then
            txtTo.Value = newRecipient
        Else
            txtTo.Value = txtTo.Value & "; " & newRecipient
        End If
    End If
End Sub

Private Sub btnAddCC_Click()
    ' CC 추가
    Dim newCC As String
    newCC = InputBox("CC에 추가할 이메일 주소를 입력하세요:", "CC 추가")
    If newCC <> "" Then
        If txtCC.Value = "" Then
            txtCC.Value = newCC
        Else
            txtCC.Value = txtCC.Value & "; " & newCC
        End If
    End If
End Sub

Private Sub btnAttach_Click()
    ' 첨부파일 추가 (시뮬레이션)
    txtAttachments.Value = "Critical_Issues_Report_" & Format(Date, "yyyymmdd") & ".xlsx"
    MsgBox "보고서 파일이 첨부되었습니다.", vbInformation
End Sub

Private Sub btnSend_Click()
    ' 입력 검증
    If txtTo.Value = "" Then
        MsgBox "수신자를 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    If txtSubject.Value = "" Then
        MsgBox "제목을 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' 발송 확인
    Dim result As VbMsgBoxResult
    result = MsgBox("다음 내용으로 이메일을 발송하시겠습니까?" & vbLf & vbLf & _
                    "수신: " & txtTo.Value & vbLf & _
                    "CC: " & txtCC.Value & vbLf & _
                    "제목: " & txtSubject.Value & vbLf & _
                    "우선순위: " & cboPriority.Value & vbLf & _
                    "첨부: " & txtAttachments.Value, _
                    vbYesNo + vbQuestion, "이메일 발송 확인")
    
    If result = vbYes Then
        ' 발송 시뮬레이션
        Application.StatusBar = "이메일 발송 중..."
        Application.Wait Now + TimeValue("00:00:02")
        
        ' 발송 기록 저장
        Call SaveEmailLog
        
        Application.StatusBar = False
        MsgBox "이메일이 성공적으로 발송되었습니다!" & vbLf & vbLf & _
               "발송 시간: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbLf & _
               "수신자 수: " & UBound(Split(txtTo.Value, ";")) + 1 & "명", _
               vbInformation, "발송 완료"
        
        Me.Hide
    End If
End Sub

Private Sub SaveEmailLog()
    ' 이메일 발송 기록 저장
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Email Log")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Email Log"
        ws.Visible = xlSheetHidden
        
        ' 헤더 생성
        ws.Range("A1").Value = "발송일시"
        ws.Range("B1").Value = "수신자"
        ws.Range("C1").Value = "제목"
        ws.Range("D1").Value = "우선순위"
        ws.Range("E1").Value = "상태"
    End If
    
    ' 새 로그 추가
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ws.Cells(lastRow, 1).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Cells(lastRow, 2).Value = txtTo.Value
    ws.Cells(lastRow, 3).Value = txtSubject.Value
    ws.Cells(lastRow, 4).Value = cboPriority.Value
    ws.Cells(lastRow, 5).Value = "발송완료"
End Sub

Private Sub btnPreview_Click()
    ' 미리보기
    MsgBox "이메일 미리보기" & vbLf & vbLf & _
           "제목: " & txtSubject.Value & vbLf & vbLf & _
           "본문:" & vbLf & _
           Left(txtBody.Value, 500) & "...", _
           vbInformation, "미리보기"
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnSaveDraft_Click()
    ' 임시저장
    MsgBox "이메일이 임시저장되었습니다.", vbInformation
    Me.Hide
End Sub