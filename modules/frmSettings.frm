VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "STRIX 검색 설정"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' STRIX 설정 다이얼로그
Option Explicit

Private Sub UserForm_Initialize()
    ' 폼 초기화
    Me.Caption = "STRIX 검색 설정"
    
    ' 현재 설정값 로드
    If modSettings.InternalWeight = 0 Then
        modSettings.InitializeSettings
    End If
    
    ' 슬라이더 설정
    With Me.sliderInternal
        .Min = 0
        .Max = 100
        .Value = modSettings.InternalWeight * 100
        .TickFrequency = 10
    End With
    
    ' 레이블 업데이트
    UpdateLabels
    
    ' 시점 콤보박스 설정
    Dim periods As Variant
    periods = modSettings.GetTimePeriodOptions()
    
    Dim i As Integer
    For i = 0 To UBound(periods)
        Me.cboTimePeriod.AddItem periods(i)
    Next i
    
    ' 현재 선택된 시점 표시
    Me.cboTimePeriod.Value = modSettings.TimePeriod
End Sub

Private Sub sliderInternal_Change()
    UpdateLabels
End Sub

Private Sub UpdateLabels()
    Dim internal As Integer
    Dim external As Integer
    
    internal = Me.sliderInternal.Value
    external = 100 - internal
    
    Me.lblInternalValue.Caption = internal & "%"
    Me.lblExternalValue.Caption = external & "%"
    
    ' 배경색 변경으로 비중 시각화
    If internal > 50 Then
        Me.lblInternalValue.ForeColor = RGB(0, 100, 0)
        Me.lblExternalValue.ForeColor = RGB(100, 100, 100)
    ElseIf internal < 50 Then
        Me.lblInternalValue.ForeColor = RGB(100, 100, 100)
        Me.lblExternalValue.ForeColor = RGB(0, 100, 0)
    Else
        Me.lblInternalValue.ForeColor = RGB(0, 0, 0)
        Me.lblExternalValue.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub btnApply_Click()
    ' 설정 적용
    Dim internal As Double
    Dim external As Double
    Dim period As String
    
    internal = Me.sliderInternal.Value / 100
    external = 1 - internal
    period = Me.cboTimePeriod.Value
    
    ' 설정 저장 및 적용
    modSettings.ApplySettings internal, external, period
    
    ' 다이얼로그 닫기
    Unload Me
    
    ' 적용 메시지
    MsgBox "설정이 적용되었습니다." & vbCrLf & vbCrLf & _
           "사내 문서: " & Format(internal * 100, "0") & "%" & vbCrLf & _
           "사외 문서: " & Format(external * 100, "0") & "%" & vbCrLf & _
           "검색 기간: " & period, _
           vbInformation, "설정 완료"
End Sub

Private Sub btnCancel_Click()
    ' 취소
    Unload Me
End Sub

Private Sub btnReset_Click()
    ' 기본값으로 리셋
    Me.sliderInternal.Value = 50
    Me.cboTimePeriod.Value = "전체기간"
    UpdateLabels
End Sub