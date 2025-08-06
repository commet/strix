Attribute VB_Name = "modDashboardEnhanced"
' Enhanced Dashboard with Reference Display
Option Explicit

Sub CreateEnhancedDashboard()
    Dim ws As Worksheet
    Dim btn As Object
    Dim shp As Shape
    
    ' 기존 Dashboard 삭제
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Dashboard").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' 새 Dashboard 생성
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Dashboard"
    ws.Activate
    
    ' 전체 배경색
    ws.Cells.Interior.Color = RGB(250, 250, 250)
    
    ' 열 너비 설정
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C:F").ColumnWidth = 20
    ws.Columns("G").ColumnWidth = 2
    
    ' 행 높이 설정
    ws.Rows("1").RowHeight = 10
    ws.Rows("2").RowHeight = 40
    
    ' ===== 1. 헤더 =====
    With ws.Range("B2:F2")
        .Merge
        .Value = "STRIX Intelligence Dashboard v2.0"
        .Font.Name = "맑은 고딕"
        .Font.Size = 22
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' 부제목
    With ws.Range("B3:F3")
        .Merge
        .Value = "AI 기반 문서 검색 시스템 (레퍼런스 포함)"
        .Font.Size = 12
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .RowHeight = 25
    End With
    
    ' ===== 2. 질문 입력 영역 =====
    ws.Range("B5").Value = "질문:"
    ws.Range("B5").Font.Bold = True
    ws.Range("B5").Font.Size = 12
    
    With ws.Range("C5:F5")
        .Merge
        .Name = "QuestionInput"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .Font.Size = 11
        .Value = "여기에 질문을 입력하세요"
        .Font.Color = RGB(150, 150, 150)
    End With
    
    ' ===== 3. 메인 버튼 =====
    Set btn = ws.Buttons.Add(ws.Range("B7").Left, ws.Range("B7").Top, 120, 35)
    With btn
        .Caption = "🔍 검색하기"
        .OnAction = "RunSearchWithSources"  ' 소스 포함 검색
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("C7").Left + 40, ws.Range("C7").Top, 120, 35)
    With btn
        .Caption = "💬 대화창"
        .OnAction = "ShowSTRIXDialog"
        .Font.Size = 12
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("E7").Left, ws.Range("E7").Top, 120, 35)
    With btn
        .Caption = "🔄 초기화"
        .OnAction = "ClearAll"
        .Font.Size = 12
    End With
    
    ' ===== 4. 답변 영역 =====
    With ws.Range("B9:F9")
        .Merge
        .Value = "📋 답변"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(46, 204, 113)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 답변 표시 영역 (크기 축소)
    With ws.Range("B10:F20")
        .Merge
        .Name = "AnswerArea"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .Font.Size = 11
        .Value = "답변이 여기에 표시됩니다..."
        .Font.Color = RGB(150, 150, 150)
    End With
    
    ' ===== 5. 레퍼런스 영역 =====
    With ws.Range("B22:F22")
        .Merge
        .Value = "📚 참고 문서"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(52, 152, 219)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 레퍼런스 표시 영역
    With ws.Range("B23:F35")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .Font.Size = 10
    End With
    
    ' 레퍼런스 헤더
    ws.Range("B23").Value = "번호"
    ws.Range("C23").Value = "제목"
    ws.Range("D23").Value = "조직/출처"
    ws.Range("E23").Value = "날짜"
    ws.Range("F23").Value = "유형"
    With ws.Range("B23:F23")
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' ===== 6. 빠른 질문 섹션 =====
    ws.Range("B37").Value = "빠른 질문:"
    ws.Range("B37").Font.Bold = True
    ws.Range("B37").Font.Size = 12
    
    ' 빠른 질문 버튼들
    Set btn = ws.Buttons.Add(ws.Range("B38").Left, ws.Range("B38").Top, 180, 30)
    With btn
        .Caption = "전고체 배터리 개발 현황"
        .OnAction = "QuickQuestion1"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("D38").Left, ws.Range("D38").Top, 180, 30)
    With btn
        .Caption = "최근 배터리 시장 동향"
        .OnAction = "QuickQuestion2"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("B39").Left, ws.Range("B39").Top + 10, 180, 30)
    With btn
        .Caption = "ESG 규제 현황"
        .OnAction = "QuickQuestion3"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("D39").Left, ws.Range("D39").Top + 10, 180, 30)
    With btn
        .Caption = "경쟁사 기술 동향"
        .OnAction = "QuickQuestion4"
        .Font.Size = 11
    End With
    
    ' ===== 7. 상태바 =====
    With ws.Range("B41:F41")
        .Merge
        .Name = "StatusBar"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Value = "✅ 준비 완료 - 레퍼런스 기능 활성화"
        .HorizontalAlignment = xlCenter
        .Font.Size = 10
        .Font.Color = RGB(0, 150, 0)
    End With
    
    ' ===== 8. 범례 =====
    With ws.Range("B43:F43")
        .Merge
        .Value = "💡 Tip: 답변의 [1], [2] 번호는 아래 참고 문서와 매칭됩니다"
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' 화면 설정
    ws.Range("B5").Select
    ActiveWindow.Zoom = 90
    ActiveWindow.DisplayGridlines = False
    
    MsgBox "Enhanced STRIX Dashboard가 생성되었습니다!" & vbLf & vbLf & _
           "주요 기능:" & vbLf & _
           "- 답변에 참고 문서 번호 표시 [1], [2]..." & vbLf & _
           "- 각 문서의 상세 정보 확인 가능" & vbLf & _
           "- 내부/외부 문서 구분 표시" & vbLf & vbLf & _
           "API 서버 실행: py api_server_with_sources.py", _
           vbInformation, "STRIX v2.0"
End Sub

' 전체 초기화
Sub ClearAll()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 질문 초기화
    ws.Range("C5").Value = "여기에 질문을 입력하세요"
    ws.Range("C5").Font.Color = RGB(150, 150, 150)
    
    ' 답변 초기화
    ws.Range("B10").Value = "답변이 여기에 표시됩니다..."
    ws.Range("B10").Font.Color = RGB(150, 150, 150)
    
    ' 레퍼런스 초기화
    ws.Range("B24:F35").Clear
    With ws.Range("B24:F35")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    
    ' 헤더 다시 생성
    ws.Range("B23").Value = "번호"
    ws.Range("C23").Value = "제목"
    ws.Range("D23").Value = "조직/출처"
    ws.Range("E23").Value = "날짜"
    ws.Range("F23").Value = "유형"
    With ws.Range("B23:F23")
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 상태 업데이트
    ws.Range("B41").Value = "✅ 초기화 완료"
End Sub

' 빠른 질문들 (레퍼런스 포함 검색 사용)
Sub QuickQuestion1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("C5").Value = "전고체 배터리 개발 현황은?"
    ws.Range("C5").Font.Color = RGB(0, 0, 0)
    Call RunSearchWithSources
End Sub

Sub QuickQuestion2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("C5").Value = "최근 배터리 시장 동향은?"
    ws.Range("C5").Font.Color = RGB(0, 0, 0)
    Call RunSearchWithSources
End Sub

Sub QuickQuestion3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("C5").Value = "ESG 규제 현황과 대응 방안은?"
    ws.Range("C5").Font.Color = RGB(0, 0, 0)
    Call RunSearchWithSources
End Sub

Sub QuickQuestion4()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("C5").Value = "경쟁사의 기술 개발 동향은?"
    ws.Range("C5").Font.Color = RGB(0, 0, 0)
    Call RunSearchWithSources
End Sub