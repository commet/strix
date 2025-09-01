Attribute VB_Name = "modSettingsSimple"
' STRIX 설정 모듈 (간단 버전)
Option Explicit

' 설정 다이얼로그 표시 (InputBox 방식)
Sub ShowSettingsSimple()
    Dim ws As Worksheet
    Dim internalPercent As String
    Dim timePeriod As String
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 현재 설정 표시
    response = MsgBox("STRIX 검색 설정을 변경하시겠습니까?" & vbCrLf & vbCrLf & _
                     "현재 설정:" & vbCrLf & _
                     "• 사내 문서 가중치: 50%" & vbCrLf & _
                     "• 사외 문서 가중치: 50%" & vbCrLf & _
                     "• 검색 기간: 전체기간", _
                     vbYesNo + vbQuestion, "검색 설정")
    
    If response = vbNo Then Exit Sub
    
    ' 가중치 입력
    internalPercent = InputBox("사내 문서 가중치를 입력하세요 (0-100):" & vbCrLf & vbCrLf & _
                               "• 100 = 사내 문서만 검색" & vbCrLf & _
                               "• 50 = 균등하게 검색" & vbCrLf & _
                               "• 0 = 사외 문서만 검색", _
                               "가중치 설정", "50")
    
    If internalPercent = "" Then Exit Sub
    
    ' 유효성 검사
    If Not IsNumeric(internalPercent) Then
        MsgBox "숫자를 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    Dim internalVal As Integer
    internalVal = CInt(internalPercent)
    
    If internalVal < 0 Or internalVal > 100 Then
        MsgBox "0에서 100 사이의 값을 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' 시점 선택
    Dim periods As String
    periods = "1. 전체기간" & vbCrLf & _
              "2. 최근 1개월" & vbCrLf & _
              "3. 최근 3개월" & vbCrLf & _
              "4. 최근 6개월" & vbCrLf & _
              "5. 2024년" & vbCrLf & _
              "6. 2024년 하반기"
    
    timePeriod = InputBox("검색 기간을 선택하세요 (1-6):" & vbCrLf & vbCrLf & periods, _
                         "검색 기간 설정", "1")
    
    If timePeriod = "" Then Exit Sub
    
    ' 시점 텍스트 변환
    Dim periodText As String
    Select Case timePeriod
        Case "1": periodText = "전체기간"
        Case "2": periodText = "최근 1개월"
        Case "3": periodText = "최근 3개월"
        Case "4": periodText = "최근 6개월"
        Case "5": periodText = "2024년"
        Case "6": periodText = "2024년 하반기"
        Case Else: periodText = "전체기간"
    End Select
    
    ' 설정 적용
    ApplySettingsSimple internalVal, 100 - internalVal, periodText
End Sub

' 설정 적용
Sub ApplySettingsSimple(internal As Integer, external As Integer, period As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 상태 업데이트
    Dim statusText As String
    statusText = "설정: 사내 " & internal & "% / 사외 " & external & "%"
    
    If period <> "전체기간" Then
        statusText = statusText & " | " & period
    End If
    
    ' H5 셀에 상태 표시
    ws.Range("H5").Value = statusText
    ws.Range("H5").Font.Size = 9
    ws.Range("H5").Font.Color = RGB(100, 100, 100)
    ws.Range("H5").Font.Italic = True
    
    ' 상태바에 알림
    ws.Range("B64").Value = "⚙️ 설정 적용: " & statusText
    ws.Range("B64").Font.Color = RGB(0, 100, 200)
    
    ' 적용 완료 메시지
    MsgBox "설정이 적용되었습니다!" & vbCrLf & vbCrLf & _
           "• 사내 문서: " & internal & "%" & vbCrLf & _
           "• 사외 문서: " & external & "%" & vbCrLf & _
           "• 검색 기간: " & period & vbCrLf & vbCrLf & _
           "다음 검색부터 적용됩니다.", _
           vbInformation, "설정 완료"
End Sub

' Dashboard에 설정 버튼 추가
Sub AddSettingsButton()
    Dim ws As Worksheet
    Dim btn As Object
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 기존 버튼 삭제
    On Error Resume Next
    ws.Shapes("btnSettings").Delete
    On Error GoTo 0
    
    ' 설정 버튼 생성 (F5 위치에)
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        Range("F5").Left, _
        Range("F5").Top, _
        70, 28)
    
    With btn
        .Name = "btnSettings"
        .TextFrame2.TextRange.Text = "⚙️ 설정"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.ForeColor.RGB = RGB(80, 80, 80)
        .Line.Visible = msoFalse
        
        ' 클릭 이벤트 연결
        .OnAction = "ShowSettingsSimple"
        
        ' 호버 효과를 위한 설정
        .Shadow.Type = msoShadow25
        .Shadow.Visible = msoTrue
        .Shadow.Style = msoShadowStyleOuterShadow
        .Shadow.Blur = 3
        .Shadow.OffsetX = 1
        .Shadow.OffsetY = 1
        .Shadow.ForeColor.RGB = RGB(200, 200, 200)
        .Shadow.Transparency = 0.5
    End With
    
    ' 초기 설정 상태 표시
    With ws.Range("H5")
        .Value = "설정: 사내 50% / 사외 50% | 전체기간"
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
    End With
    
    MsgBox "설정 버튼이 추가되었습니다!" & vbCrLf & _
           "대화창 옆의 [⚙️ 설정] 버튼을 클릭하세요.", _
           vbInformation, "설정 버튼 추가 완료"
End Sub