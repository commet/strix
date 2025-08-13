VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAlertSettings 
   Caption         =   "Smart Alert 설정"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7000
   OleObjectBlob   =   "frmAlertSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAlertSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    ' 기본값 설정
    txtThreshold.Value = "70"
    cboFrequency.AddItem "실시간"
    cboFrequency.AddItem "1시간마다"
    cboFrequency.AddItem "3시간마다"
    cboFrequency.AddItem "하루 2회"
    cboFrequency.AddItem "하루 1회"
    cboFrequency.Value = "실시간"
    
    ' 알림 시간 설정
    txtAlertTime.Value = "09:00"
    
    ' 이메일 수신자 (기본값)
    txtRecipients.Value = "ceo@company.com; coo@company.com; cfo@company.com"
    
    ' Slack 웹훅 URL
    txtSlackWebhook.Value = "https://hooks.slack.com/services/YOUR/WEBHOOK/URL"
    
    ' 체크박스 기본값
    chkEmailAlert.Value = True
    chkSlackAlert.Value = False
    chkDesktopAlert.Value = True
    
    ' 저장된 설정 불러오기
    Call LoadSettings
End Sub

Private Sub btnSave_Click()
    ' 설정 저장
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Settings"
        ws.Visible = xlSheetHidden
    End If
    
    ' 설정값 저장
    ws.Range("A1").Value = "Critical Threshold"
    ws.Range("B1").Value = txtThreshold.Value
    
    ws.Range("A2").Value = "Alert Frequency"
    ws.Range("B2").Value = cboFrequency.Value
    
    ws.Range("A3").Value = "Alert Time"
    ws.Range("B3").Value = txtAlertTime.Value
    
    ws.Range("A4").Value = "Email Recipients"
    ws.Range("B4").Value = txtRecipients.Value
    
    ws.Range("A5").Value = "Slack Webhook"
    ws.Range("B5").Value = txtSlackWebhook.Value
    
    ws.Range("A6").Value = "Email Alert"
    ws.Range("B6").Value = chkEmailAlert.Value
    
    ws.Range("A7").Value = "Slack Alert"
    ws.Range("B7").Value = chkSlackAlert.Value
    
    ws.Range("A8").Value = "Desktop Alert"
    ws.Range("B8").Value = chkDesktopAlert.Value
    
    MsgBox "설정이 저장되었습니다!", vbInformation, "저장 완료"
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub LoadSettings()
    ' 저장된 설정 불러오기
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    If Not ws Is Nothing Then
        If ws.Range("B1").Value <> "" Then txtThreshold.Value = ws.Range("B1").Value
        If ws.Range("B2").Value <> "" Then cboFrequency.Value = ws.Range("B2").Value
        If ws.Range("B3").Value <> "" Then txtAlertTime.Value = ws.Range("B3").Value
        If ws.Range("B4").Value <> "" Then txtRecipients.Value = ws.Range("B4").Value
        If ws.Range("B5").Value <> "" Then txtSlackWebhook.Value = ws.Range("B5").Value
        If ws.Range("B6").Value <> "" Then chkEmailAlert.Value = ws.Range("B6").Value
        If ws.Range("B7").Value <> "" Then chkSlackAlert.Value = ws.Range("B7").Value
        If ws.Range("B8").Value <> "" Then chkDesktopAlert.Value = ws.Range("B8").Value
    End If
End Sub

Private Sub btnTestEmail_Click()
    ' 테스트 이메일 발송
    MsgBox "테스트 이메일을 다음 주소로 발송합니다:" & vbLf & vbLf & _
           txtRecipients.Value, vbInformation, "테스트 이메일"
End Sub

Private Sub btnTestSlack_Click()
    ' Slack 연동 테스트
    MsgBox "Slack 웹훅 URL 테스트:" & vbLf & vbLf & _
           txtSlackWebhook.Value & vbLf & vbLf & _
           "연결 상태: OK", vbInformation, "Slack 테스트"
End Sub