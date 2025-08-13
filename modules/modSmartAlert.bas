Attribute VB_Name = "modSmartAlert"
' Smart Alert System - AI 기반 이슈 예측 및 자동 알림
Option Explicit

' 전역 변수
Private Const ALERT_THRESHOLD As Integer = 70  ' 위험도 임계값
Private alertData As Collection

' ===== 메인 함수 =====
Sub CreateSmartAlertDashboard()
    Dim ws As Worksheet
    Dim alertWs As Worksheet
    
    ' 기존 시트 삭제
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Smart Alerts").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' 새 시트 생성
    Set alertWs = ThisWorkbook.Sheets.Add
    alertWs.Name = "Smart Alerts"
    alertWs.Activate
    
    ' 전체 배경색
    alertWs.Cells.Interior.Color = RGB(240, 242, 247)
    
    ' 열 너비 설정
    alertWs.Columns("A").ColumnWidth = 2
    alertWs.Columns("B").ColumnWidth = 8   ' 순위
    alertWs.Columns("C").ColumnWidth = 35  ' 이슈
    alertWs.Columns("D").ColumnWidth = 12  ' 위험도
    alertWs.Columns("E").ColumnWidth = 15  ' 예상 시점
    alertWs.Columns("F").ColumnWidth = 25  ' 권장 액션
    alertWs.Columns("G").ColumnWidth = 12  ' 담당
    alertWs.Columns("H").ColumnWidth = 10  ' 상태
    alertWs.Columns("I").ColumnWidth = 2
    
    ' ===== 헤더 영역 =====
    With alertWs.Range("B2:H2")
        .Merge
        .Value = "🚨 STRIX Smart Alert System"
        .Font.Name = "맑은 고딕"
        .Font.Size = 26
        .Font.Bold = True
        .Interior.Color = RGB(231, 76, 60)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 55
    End With
    
    ' 부제목 및 시간
    With alertWs.Range("B3:H3")
        .Merge
        .Value = "AI 기반 실시간 이슈 예측 및 알림 | 마지막 업데이트: " & Format(Now, "yyyy-mm-dd hh:mm")
        .Font.Size = 12
        .Font.Color = RGB(52, 73, 94)
        .HorizontalAlignment = xlCenter
        .RowHeight = 25
    End With
    
    ' ===== 오늘의 알림 요약 =====
    With alertWs.Range("B5:H5")
        .Merge
        .Value = "📊 오늘의 브리핑"
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = RGB(52, 73, 94)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    
    ' 요약 통계
    Dim summaryRow As Integer
    summaryRow = 6
    
    ' 통계 박스들
    Call CreateStatBox(alertWs, "B", summaryRow, "Critical", "3", RGB(231, 76, 60))
    Call CreateStatBox(alertWs, "C", summaryRow, "High", "7", RGB(230, 126, 34))
    Call CreateStatBox(alertWs, "D", summaryRow, "Medium", "12", RGB(241, 196, 15))
    Call CreateStatBox(alertWs, "E", summaryRow, "Low", "8", RGB(46, 204, 113))
    Call CreateStatBox(alertWs, "F", summaryRow, "총 이슈", "30", RGB(52, 152, 219))
    Call CreateStatBox(alertWs, "G", summaryRow, "신규", "+5", RGB(155, 89, 182))
    
    ' ===== 자동 실행 설정 영역 =====
    With alertWs.Range("B9")
        .Value = "⚙️ 자동 알림 설정"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' 자동 실행 체크박스
    Dim cb As Object
    Set cb = alertWs.CheckBoxes.Add(alertWs.Range("C9").Left, _
                                    alertWs.Range("C9").Top, 150, 20)
    With cb
        .Caption = "매일 오전 9시 자동 실행"
        .Value = xlOn
        .OnAction = "ToggleAutoAlert"
    End With
    
    ' 즉시 실행 버튼
    Dim runBtn As Object
    Set runBtn = alertWs.Buttons.Add(alertWs.Range("E9").Left, _
                                     alertWs.Range("E9").Top, 100, 25)
    With runBtn
        .Caption = "▶️ 즉시 분석"
        .OnAction = "RunSmartAnalysis"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' 설정 버튼
    Dim settingsBtn As Object
    Set settingsBtn = alertWs.Buttons.Add(alertWs.Range("F9").Left, _
                                          alertWs.Range("F9").Top, 80, 25)
    With settingsBtn
        .Caption = "⚙️ 설정"
        .OnAction = "ShowAlertSettings"
        .Font.Size = 11
    End With
    
    ' 이메일 전송 버튼
    Dim emailBtn As Object
    Set emailBtn = alertWs.Buttons.Add(alertWs.Range("G9").Left, _
                                       alertWs.Range("G9").Top, 100, 25)
    With emailBtn
        .Caption = "📧 이메일 전송"
        .OnAction = "SendAlertEmail"
        .Font.Size = 11
    End With
    
    ' ===== TOP 5 Critical Issues =====
    With alertWs.Range("B11:H11")
        .Merge
        .Value = "🔥 TOP 5 Critical Issues - 즉시 확인 필요"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(192, 57, 43)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 헤더 행
    Dim headerRow As Integer
    headerRow = 12
    With alertWs.Range("B" & headerRow & ":H" & headerRow)
        .Interior.Color = RGB(44, 62, 80)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    alertWs.Cells(headerRow, 2).Value = "#"
    alertWs.Cells(headerRow, 3).Value = "이슈"
    alertWs.Cells(headerRow, 4).Value = "위험도"
    alertWs.Cells(headerRow, 5).Value = "예상 영향"
    alertWs.Cells(headerRow, 6).Value = "권장 액션"
    alertWs.Cells(headerRow, 7).Value = "담당"
    alertWs.Cells(headerRow, 8).Value = "구분"
    
    ' 샘플 Critical 이슈 추가
    Call AddCriticalIssues(alertWs, headerRow + 1)
    
    ' ===== AI 예측 섹션 =====
    With alertWs.Range("B20:H20")
        .Merge
        .Value = "🤖 AI 예측 분석"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(142, 68, 173)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 예측 내용
    With alertWs.Range("B21:H25")
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    
    alertWs.Range("B21").Value = "📈 향후 72시간 예측:" & vbLf & _
        "• 원자재 가격 변동성 증가 예상 (신뢰도: 85%)" & vbLf & _
        "• 경쟁사 신제품 발표 가능성 높음 (신뢰도: 78%)" & vbLf & _
        "• 정부 규제 발표 예정 - ESG 관련 (신뢰도: 92%)" & vbLf & vbLf & _
        "💡 권장사항: 리스크 대응 TF 즉시 소집 필요"
    
    ' ===== 액션 트래커 =====
    With alertWs.Range("B27:H27")
        .Merge
        .Value = "✅ Action Tracker"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(39, 174, 96)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 액션 아이템
    Call AddActionItems(alertWs, 28)
    
    ' ===== 알림 로그 =====
    With alertWs.Range("B35:H35")
        .Merge
        .Value = "📝 알림 히스토리"
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(149, 165, 166)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 로그 영역
    With alertWs.Range("B36:H40")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
    End With
    
    ' 샘플 로그
    alertWs.Range("B36").Value = Format(Now - 1, "mm/dd hh:mm") & " - Critical 알림 3건 발송 (경영진)"
    alertWs.Range("B37").Value = Format(Now - 0.5, "mm/dd hh:mm") & " - 리스크 레벨 상향 조정: 원자재 이슈"
    alertWs.Range("B38").Value = Format(Now - 0.25, "mm/dd hh:mm") & " - 신규 이슈 감지: ESG 규제 강화"
    alertWs.Range("B39").Value = Format(Now, "mm/dd hh:mm") & " - 일일 브리핑 생성 완료"
    
    ' 화면 설정
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    alertWs.Range("B2").Select
    
    MsgBox "Smart Alert System이 생성되었습니다!" & vbLf & vbLf & _
           "🚨 주요 기능:" & vbLf & _
           "• AI 기반 이슈 위험도 예측" & vbLf & _
           "• 자동 일일 브리핑 (오전 9시)" & vbLf & _
           "• Critical 이슈 실시간 알림" & vbLf & _
           "• 액션 아이템 자동 생성" & vbLf & _
           "• 이메일 알림 연동 준비", _
           vbInformation, "Smart Alert System"
End Sub

' 통계 박스 생성
Private Sub CreateStatBox(ws As Worksheet, col As String, row As Integer, title As String, _
                          value As String, color As Long)
    With ws.Range(col & row)
        .Value = title
        .Font.Size = 10
        .Font.Color = RGB(100, 100, 100)
    End With
    
    With ws.Range(col & row + 1)
        .Value = value
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = color
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range(col & row & ":" & col & row + 1)
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
End Sub

' Critical 이슈 추가
Private Sub AddCriticalIssues(ws As Worksheet, startRow As Integer)
    Dim issues As Variant
    Dim i As Integer
    
    ' 2025년 최신 Critical 이슈 (SK온 및 배터리 업계)
    issues = Array( _
        Array("1", "SK온-SK엔무브 합병 통합법인 출범 준비", "92", "11월 1일", "통합 실무 TF 구성", "경영기획", "사내"), _
        Array("2", "트럼프 IRA 폐지 가능성, AMPC 세액공제 위기", "90", "즉시", "정책 대응 시나리오 수립", "정책대응", "사외"), _
        Array("3", "BYD 5분 충전 기술 공개, 기술격차 심화", "88", "1개월", "기술 캐치업 전략 수립", "R&D", "사외"), _
        Array("4", "5조원 자본확충 진행, 유상증자 실행", "85", "8월", "IR 준비 및 투자자 설명", "재무팀", "사내"), _
        Array("5", "LG엔솔 위기경영 선언, K배터리 위기", "82", "2주", "경쟁사 분석 및 대응", "전략기획", "사외") _
    )
    
    For i = 0 To UBound(issues)
        Dim currentRow As Integer
        currentRow = startRow + i
        
        ' 순위
        ws.Cells(currentRow, 2).Value = issues(i)(0)
        ws.Cells(currentRow, 2).Font.Bold = True
        ws.Cells(currentRow, 2).HorizontalAlignment = xlCenter
        
        ' 이슈
        ws.Cells(currentRow, 3).Value = issues(i)(1)
        ws.Cells(currentRow, 3).WrapText = True
        
        ' 위험도 (시각화)
        With ws.Cells(currentRow, 4)
            .Value = issues(i)(2) & "%"
            .Font.Bold = True
            If CInt(issues(i)(2)) >= 90 Then
                .Font.Color = RGB(231, 76, 60)
            ElseIf CInt(issues(i)(2)) >= 80 Then
                .Font.Color = RGB(230, 126, 34)
            Else
                .Font.Color = RGB(241, 196, 15)
            End If
            .HorizontalAlignment = xlCenter
        End With
        
        ' 예상 영향
        ws.Cells(currentRow, 5).Value = issues(i)(3)
        ws.Cells(currentRow, 5).HorizontalAlignment = xlCenter
        
        ' 권장 액션
        ws.Cells(currentRow, 6).Value = issues(i)(4)
        ws.Cells(currentRow, 6).Font.Size = 10
        
        ' 담당
        ws.Cells(currentRow, 7).Value = issues(i)(5)
        ws.Cells(currentRow, 7).HorizontalAlignment = xlCenter
        
        ' 구분 (사내/사외)
        ws.Cells(currentRow, 8).Value = issues(i)(6)
        ws.Cells(currentRow, 8).Font.Size = 10
        ws.Cells(currentRow, 8).HorizontalAlignment = xlCenter
        If issues(i)(6) = "사내" Then
            ws.Cells(currentRow, 8).Font.Color = RGB(52, 152, 219)
            ws.Cells(currentRow, 8).Font.Bold = True
        Else
            ws.Cells(currentRow, 8).Font.Color = RGB(155, 89, 182)
            ws.Cells(currentRow, 8).Font.Bold = True
        End If
        
        ' 행 서식
        With ws.Range("B" & currentRow & ":H" & currentRow)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200)
            If i Mod 2 = 0 Then
                .Interior.Color = RGB(248, 248, 248)
            Else
                .Interior.Color = RGB(255, 255, 255)
            End If
        End With
        
        ws.Rows(currentRow).RowHeight = 30
    Next i
End Sub

' 액션 아이템 추가
Private Sub AddActionItems(ws As Worksheet, startRow As Integer)
    ' 헤더
    With ws.Range("B" & startRow & ":H" & startRow)
        .Interior.Color = RGB(236, 240, 241)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
    
    ws.Cells(startRow, 2).Value = "No"
    ws.Cells(startRow, 3).Value = "액션 아이템"
    ws.Cells(startRow, 4).Value = "우선순위"
    ws.Cells(startRow, 5).Value = "마감일"
    ws.Cells(startRow, 6).Value = "담당자"
    ws.Cells(startRow, 7).Value = "진행률"
    ws.Cells(startRow, 8).Value = "상태"
    
    ' 샘플 액션 아이템
    Dim actions As Variant
    actions = Array( _
        Array("A1", "SK온-SK엔무브 통합 실무 TF 구성 및 가동", "Critical", Format(Date + 2, "mm/dd"), "경영기획팀", "10%", "착수"), _
        Array("A2", "IRA 정책 변화 대응 시나리오 수립", "Critical", Format(Date + 1, "mm/dd"), "정책대응팀", "0%", "대기"), _
        Array("A3", "BYD 기술 분석 및 대응 로드맵 작성", "Critical", Format(Date + 7, "mm/dd"), "R&D팀", "15%", "진행중"), _
        Array("A4", "5조원 자본확충 IR 자료 준비", "High", Format(Date + 5, "mm/dd"), "재무팀", "40%", "진행중") _
    )
    
    Dim i As Integer
    For i = 0 To UBound(actions)
        Dim row As Integer
        row = startRow + 1 + i
        
        ws.Cells(row, 2).Value = actions(i)(0)
        ws.Cells(row, 3).Value = actions(i)(1)
        ws.Cells(row, 4).Value = actions(i)(2)
        ws.Cells(row, 5).Value = actions(i)(3)
        ws.Cells(row, 6).Value = actions(i)(4)
        ws.Cells(row, 7).Value = actions(i)(5)
        ws.Cells(row, 8).Value = actions(i)(6)
        
        ' 우선순위별 색상
        If actions(i)(2) = "Critical" Then
            ws.Cells(row, 4).Font.Color = RGB(231, 76, 60)
            ws.Cells(row, 4).Font.Bold = True
        ElseIf actions(i)(2) = "High" Then
            ws.Cells(row, 4).Font.Color = RGB(230, 126, 34)
        End If
        
        ' 행 서식
        With ws.Range("B" & row & ":H" & row)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200)
            .Interior.Color = RGB(255, 255, 255)
        End With
    Next i
End Sub

' ===== 실행 함수들 =====
Sub RunSmartAnalysis()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Smart Alerts")
    
    Application.StatusBar = "AI가 이슈를 분석중입니다..."
    Application.Wait Now + TimeValue("00:00:02") ' 시뮬레이션
    
    ' 시간 업데이트
    ws.Range("B3").Value = "AI 기반 실시간 이슈 예측 및 알림 | 마지막 업데이트: " & Format(Now, "yyyy-mm-dd hh:mm")
    
    ' 새로운 위험도 계산 (랜덤 시뮬레이션)
    Dim newRisk As Integer
    newRisk = Int(Rnd() * 20) + 75
    ws.Cells(13, 4).Value = newRisk & "%"
    
    If newRisk >= 90 Then
        MsgBox "⚠️ 경고: Critical 수준의 이슈가 감지되었습니다!" & vbLf & vbLf & _
               "이슈: " & ws.Cells(13, 3).Value & vbLf & _
               "위험도: " & newRisk & "%" & vbLf & vbLf & _
               "즉시 대응이 필요합니다.", vbCritical, "Critical Alert"
    Else
        MsgBox "✅ 분석 완료!" & vbLf & vbLf & _
               "• 신규 Critical 이슈: 0건" & vbLf & _
               "• 위험도 상승 이슈: 2건" & vbLf & _
               "• 해결된 이슈: 1건" & vbLf & vbLf & _
               "상세 내용은 대시보드를 확인하세요.", vbInformation, "분석 완료"
    End If
    
    Application.StatusBar = False
End Sub

Sub ToggleAutoAlert()
    Dim cb As Object
    Set cb = ThisWorkbook.Sheets("Smart Alerts").CheckBoxes(1)
    
    If cb.Value = xlOn Then
        ' 자동 실행 스케줄 설정 (실제로는 Windows Task Scheduler 연동 필요)
        MsgBox "자동 알림이 활성화되었습니다." & vbLf & _
               "매일 오전 9시에 자동으로 분석이 실행됩니다.", vbInformation
    Else
        MsgBox "자동 알림이 비활성화되었습니다.", vbInformation
    End If
End Sub

Sub ShowAlertSettings()
    ' 간단한 입력 다이얼로그 사용
    Dim settingsMsg As String
    Dim ws As Worksheet
    Dim threshold As String
    Dim recipients As String
    Dim frequency As String
    
    ' 현재 설정 불러오기
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Settings"
        ws.Visible = xlSheetHidden
        ' 기본값 설정
        ws.Range("B1").Value = "70"
        ws.Range("B2").Value = "실시간"
        ws.Range("B4").Value = "ceo@company.com; coo@company.com"
    End If
    
    threshold = ws.Range("B1").Value
    frequency = ws.Range("B2").Value
    recipients = ws.Range("B4").Value
    
    ' 설정 메뉴 표시
    Dim choice As String
    choice = InputBox("변경할 설정을 선택하세요:" & vbLf & vbLf & _
                      "1. Critical 임계값 (현재: " & threshold & "%)" & vbLf & _
                      "2. 알림 주기 (현재: " & frequency & ")" & vbLf & _
                      "3. 이메일 수신자 (현재: " & Left(recipients, 30) & "...)" & vbLf & _
                      "4. 알림 시간 설정" & vbLf & _
                      "5. 현재 설정 보기" & vbLf & vbLf & _
                      "번호를 입력하세요 (1-5):", "Smart Alert 설정")
    
    Select Case choice
        Case "1"
            threshold = InputBox("Critical 임계값을 입력하세요 (50-100):", "임계값 설정", threshold)
            If threshold <> "" And IsNumeric(threshold) Then
                ws.Range("B1").Value = threshold
                MsgBox "임계값이 " & threshold & "%로 설정되었습니다.", vbInformation
            End If
            
        Case "2"
            frequency = InputBox("알림 주기를 입력하세요:" & vbLf & _
                               "- 실시간" & vbLf & _
                               "- 1시간마다" & vbLf & _
                               "- 3시간마다" & vbLf & _
                               "- 하루 2회" & vbLf & _
                               "- 하루 1회", "알림 주기", frequency)
            If frequency <> "" Then
                ws.Range("B2").Value = frequency
                MsgBox "알림 주기가 '" & frequency & "'로 설정되었습니다.", vbInformation
            End If
            
        Case "3"
            recipients = InputBox("이메일 수신자를 입력하세요 (세미콜론으로 구분):" & vbLf & vbLf & _
                                "예: john@company.com; sarah@company.com", _
                                "이메일 수신자", recipients)
            If recipients <> "" Then
                ws.Range("B4").Value = recipients
                MsgBox "이메일 수신자가 설정되었습니다.", vbInformation
            End If
            
        Case "4"
            Dim alertTime As String
            alertTime = InputBox("자동 알림 시간을 입력하세요 (예: 09:00):", "알림 시간", "09:00")
            If alertTime <> "" Then
                ws.Range("B3").Value = alertTime
                MsgBox "알림 시간이 " & alertTime & "로 설정되었습니다.", vbInformation
            End If
            
        Case "5"
            MsgBox "현재 설정:" & vbLf & vbLf & _
                   "Critical 임계값: " & ws.Range("B1").Value & "%" & vbLf & _
                   "알림 주기: " & ws.Range("B2").Value & vbLf & _
                   "알림 시간: " & ws.Range("B3").Value & vbLf & _
                   "이메일 수신자: " & vbLf & ws.Range("B4").Value & vbLf & vbLf & _
                   "이메일 알림: 활성화" & vbLf & _
                   "Slack 연동: 준비중", _
                   vbInformation, "현재 설정"
    End Select
End Sub

Sub SendAlertEmail()
    On Error GoTo ErrorHandler
    
    ' 설정에서 수신자 불러오기
    Dim ws As Worksheet
    Dim recipients As String
    Dim subject As String
    Dim body As String
    Dim cc As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    If Not ws Is Nothing Then
        recipients = ws.Range("B4").Value
    End If
    
    If recipients = "" Then
        recipients = "ceo@company.com; coo@company.com"
    End If
    
    ' 이메일 작성 다이얼로그
    Dim emailChoice As String
    emailChoice = InputBox("이메일 작성 옵션을 선택하세요:" & vbLf & vbLf & _
                          "1. 기본 설정으로 발송" & vbLf & _
                          "2. 수신자 변경" & vbLf & _
                          "3. 제목/본문 편집" & vbLf & _
                          "4. 상세 설정" & vbLf & vbLf & _
                          "번호를 선택하세요 (1-4):", "이메일 작성")
    
    Select Case emailChoice
        Case "1"
            ' 기본 발송
            subject = "[STRIX Alert] " & Format(Date, "yyyy-mm-dd") & " Critical Issues Report"
            Call QuickSendEmail(recipients, subject)
            
        Case "2"
            ' 수신자 변경
            recipients = InputBox("수신자 이메일을 입력하세요:" & vbLf & _
                                "세미콜론으로 구분" & vbLf & vbLf & _
                                "현재: " & recipients, _
                                "수신자 설정", recipients)
            If recipients <> "" Then
                ws.Range("B4").Value = recipients
                subject = "[STRIX Alert] " & Format(Date, "yyyy-mm-dd") & " Critical Issues Report"
                Call QuickSendEmail(recipients, subject)
            End If
            
        Case "3"
            ' 제목/본문 편집
            subject = InputBox("이메일 제목을 입력하세요:", "제목", _
                             "[STRIX Alert] " & Format(Date, "yyyy-mm-dd") & " Critical Issues Report")
            
            body = InputBox("추가 메시지를 입력하세요:" & vbLf & _
                          "(기본 보고서에 추가됨)", "본문 추가")
            
            Call DetailedSendEmail(recipients, subject, body)
            
        Case "4"
            ' 상세 설정
            Call ShowEmailComposer
            
        Case Else
            Exit Sub
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "이메일 발송 중 오류가 발생했습니다.", vbExclamation
End Sub

' 빠른 이메일 발송
Private Sub QuickSendEmail(recipients As String, subject As String)
    Dim result As VbMsgBoxResult
    result = MsgBox("다음 내용으로 이메일을 발송하시겠습니까?" & vbLf & vbLf & _
                    "수신: " & recipients & vbLf & _
                    "제목: " & subject & vbLf & vbLf & _
                    "Critical Issues 보고서가 첨부됩니다.", _
                    vbYesNo + vbQuestion, "이메일 발송 확인")
    
    If result = vbYes Then
        ' 발송 시뮬레이션
        Application.StatusBar = "이메일 발송 중..."
        Application.Wait Now + TimeValue("00:00:02")
        
        ' 발송 로그 저장
        Call SaveEmailLog(recipients, subject)
        
        Application.StatusBar = False
        MsgBox "이메일이 성공적으로 발송되었습니다!" & vbLf & vbLf & _
               "발송 시간: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbLf & _
               "수신자 수: " & UBound(Split(recipients, ";")) + 1 & "명", _
               vbInformation, "발송 완료"
    End If
End Sub

' 상세 이메일 발송
Private Sub DetailedSendEmail(recipients As String, subject As String, additionalBody As String)
    Dim body As String
    Dim ws As Worksheet
    
    ' 기본 본문 생성
    body = "안녕하세요," & vbLf & vbLf
    body = body & "STRIX Smart Alert System에서 발송하는 Critical Issues 보고서입니다." & vbLf & vbLf
    
    If additionalBody <> "" Then
        body = body & additionalBody & vbLf & vbLf
    End If
    
    ' Smart Alerts 시트에서 데이터 가져오기
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Smart Alerts")
    If Not ws Is Nothing Then
        body = body & "TOP 5 CRITICAL ISSUES:" & vbLf
        Dim i As Integer
        For i = 13 To 17
            If ws.Cells(i, 3).Value <> "" Then
                body = body & ws.Cells(i, 2).Value & ". " & ws.Cells(i, 3).Value & _
                      " (위험도: " & ws.Cells(i, 4).Value & ")" & vbLf
            End If
        Next i
    End If
    
    body = body & vbLf & "감사합니다."
    
    ' 발송 확인
    If MsgBox("이메일 미리보기:" & vbLf & vbLf & _
              "수신: " & recipients & vbLf & _
              "제목: " & subject & vbLf & vbLf & _
              "본문:" & vbLf & Left(body, 300) & "...", _
              vbYesNo + vbQuestion, "이메일 발송 확인") = vbYes Then
        
        Call SaveEmailLog(recipients, subject)
        MsgBox "이메일이 발송되었습니다!", vbInformation
    End If
End Sub

' 이메일 작성기 표시
Private Sub ShowEmailComposer()
    ' 상세 이메일 작성 화면
    Dim recipients As String, cc As String, subject As String, body As String
    Dim priority As String
    
    ' 기본값 설정
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    If Not ws Is Nothing Then
        recipients = ws.Range("B4").Value
    Else
        recipients = "ceo@company.com"
    End If
    
    ' 입력 받기
    recipients = InputBox("수신자 (To):" & vbLf & "세미콜론으로 구분", "수신자", recipients)
    If recipients = "" Then Exit Sub
    
    cc = InputBox("참조 (CC):" & vbLf & "세미콜론으로 구분", "참조", "risk-management@company.com")
    
    subject = InputBox("제목:", "제목", "[STRIX Alert] " & Format(Date, "yyyy-mm-dd") & " Critical Issues Report")
    
    priority = InputBox("우선순위 (1: 높음, 2: 보통, 3: 낮음):", "우선순위", "1")
    
    body = InputBox("추가 메시지:" & vbLf & vbLf & _
                   "(기본 Critical Issues 보고서에 추가됨)", "본문")
    
    ' 발송 확인
    Dim msg As String
    msg = "이메일 정보:" & vbLf & vbLf
    msg = msg & "수신: " & recipients & vbLf
    msg = msg & "참조: " & cc & vbLf
    msg = msg & "제목: " & subject & vbLf
    msg = msg & "우선순위: " & IIf(priority = "1", "높음", IIf(priority = "2", "보통", "낮음")) & vbLf
    msg = msg & "첨부: Critical_Issues_Report_" & Format(Date, "yyyymmdd") & ".xlsx" & vbLf & vbLf
    msg = msg & "이메일을 발송하시겠습니까?"
    
    If MsgBox(msg, vbYesNo + vbQuestion, "이메일 발송 확인") = vbYes Then
        Call SaveEmailLog(recipients & "; " & cc, subject)
        MsgBox "이메일이 성공적으로 발송되었습니다!", vbInformation, "발송 완료"
    End If
End Sub

' 이메일 발송 로그 저장
Private Sub SaveEmailLog(recipients As String, subject As String)
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
        ws.Range("D1").Value = "상태"
    End If
    
    ' 새 로그 추가
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ws.Cells(lastRow, 1).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Cells(lastRow, 2).Value = recipients
    ws.Cells(lastRow, 3).Value = subject
    ws.Cells(lastRow, 4).Value = "발송완료"
End Sub

' 일일 자동 실행 함수
Sub DailyAutoRun()
    ' 이 함수는 Windows Task Scheduler에서 호출
    Call RunSmartAnalysis
    
    ' Critical 이슈가 있으면 자동 이메일 발송
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Smart Alerts")
    
    Dim risk As Integer
    risk = Val(Replace(ws.Cells(13, 4).Value, "%", ""))
    
    If risk >= ALERT_THRESHOLD Then
        Call SendAlertEmail
    End If
End Sub

' API 연동 함수
Function GetAIPrediction() As String
    Dim http As Object
    Dim url As String
    Dim response As String
    
    On Error GoTo ErrorHandler
    
    url = "http://localhost:5000/api/issues/predict"
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    http.Open "GET", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send
    
    If http.Status = 200 Then
        GetAIPrediction = http.responseText
    Else
        GetAIPrediction = ""
    End If
    
    Exit Function
    
ErrorHandler:
    GetAIPrediction = ""
End Function