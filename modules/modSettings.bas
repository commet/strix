Attribute VB_Name = "modSettings"
' STRIX 설정 모듈
Option Explicit

' 전역 설정 변수
Public InternalWeight As Double
Public ExternalWeight As Double
Public TimePeriod As String

' 설정 초기화
Sub InitializeSettings()
    ' 기본값 설정
    InternalWeight = 0.5   ' 사내 50%
    ExternalWeight = 0.5   ' 사외 50%
    TimePeriod = "전체기간"
End Sub

' 설정 버튼 클릭 핸들러
Sub ShowSettingsDialog()
    ' 설정 다이얼로그 표시
    frmSettings.Show
End Sub

' Dashboard에 설정 버튼 추가
Sub CreateSettingsButton()
    Dim ws As Worksheet
    Dim btn As Object
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 기존 버튼 삭제
    On Error Resume Next
    ws.Shapes("btnSettings").Delete
    On Error GoTo 0
    
    ' 설정 버튼 생성 (대화창 버튼 옆에)
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        Range("F5").Left, _
        Range("F5").Top, _
        80, 30)
    
    With btn
        .Name = "btnSettings"
        .TextFrame2.TextRange.Text = "⚙️ 설정"
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.ForeColor.RGB = RGB(100, 100, 100)
        .Line.Visible = msoFalse
        
        ' 클릭 이벤트 연결
        .OnAction = "ShowSettingsDialog"
    End With
    
    ' 현재 설정 상태 표시 영역
    With ws.Range("H5")
        .Value = "현재 설정: 사내 50% / 사외 50%"
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
        .Font.Italic = True
    End With
End Sub

' 설정 적용 함수
Sub ApplySettings(internal As Double, external As Double, period As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 설정 저장
    InternalWeight = internal
    ExternalWeight = external
    TimePeriod = period
    
    ' 상태 업데이트
    Dim statusText As String
    statusText = "현재 설정: 사내 " & Format(internal * 100, "0") & "% / 사외 " & Format(external * 100, "0") & "%"
    
    If period <> "전체기간" Then
        statusText = statusText & " | " & period
    End If
    
    ws.Range("H5").Value = statusText
    
    ' 설정 적용 알림
    ws.Range("B64").Value = "⚙️ 설정이 적용되었습니다: " & statusText
    ws.Range("B64").Font.Color = RGB(0, 100, 200)
    
    ' 실제로는 여기서 API에 파라미터 전달하도록 구현
    ' 현재는 시연용으로 표시만
End Sub

' 시점 옵션 가져오기
Function GetTimePeriodOptions() As Variant
    GetTimePeriodOptions = Array( _
        "전체기간", _
        "최근 1개월", _
        "최근 3개월", _
        "최근 6개월", _
        "2024년", _
        "2024년 하반기", _
        "2024년 상반기" _
    )
End Function