Option Explicit

' STRIX Dashboard 자동 생성 매크로
' 이 매크로를 실행하면 완벽한 STRIX Dashboard가 자동으로 생성됩니다

Sub CreateSTRIXDashboard()
    Dim ws As Worksheet
    Dim btn As Object
    Dim shp As Shape
    
    ' Dashboard 시트 생성 또는 초기화
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("STRIX Dashboard")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "STRIX Dashboard"
    Else
        ' 기존 내용 모두 삭제
        ws.Cells.Clear
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If
    On Error GoTo 0
    
    ' 시트 활성화
    ws.Activate
    
    ' ==== 1. 전체 레이아웃 설정 ====
    With ws
        ' 배경색 설정
        .Cells.Interior.Color = RGB(245, 245, 245)
        
        ' 열 너비 설정
        .Columns("A").ColumnWidth = 2
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 2
        .Columns("H").ColumnWidth = 20
        .Columns("I").ColumnWidth = 20
        .Columns("J").ColumnWidth = 20
        .Columns("K").ColumnWidth = 2
        
        ' 행 높이 설정
        .Rows("1").RowHeight = 10
        .Rows("2").RowHeight = 40
        .Rows("3").RowHeight = 10
        .Rows("4").RowHeight = 30
        .Rows("5").RowHeight = 10
    End With
    
    ' ==== 2. 헤더 영역 ====
    With ws.Range("B2:J2")
        .Merge
        .Value = "STRIX Intelligence Dashboard"
        .Font.Name = "Segoe UI"
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(41, 128, 185)  ' 파란색
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' 부제목
    With ws.Range("B4:J4")
        .Merge
        .Value = "AI 기반 문서 검색 및 인텔리전스 시스템"
        .Font.Name = "맑은 고딕"
        .Font.Size = 12
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
    End With
    
    ' ==== 3. 메인 컨트롤 패널 ====
    ' 패널 배경
    With ws.Range("B6:F11")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    
    ' 패널 제목
    With ws.Range("B6:F6")
        .Merge
        .Value = "🔍 검색 및 분석"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(52, 152, 219)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 질문 입력 레이블
    ws.Range("B8").Value = "질문:"
    ws.Range("B8").Font.Bold = True
    
    ' 질문 입력 영역
    With ws.Range("C8:F8")
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Name = "QuestionInput"
        .Value = "여기에 질문을 입력하세요..."
        .Font.Color = RGB(150, 150, 150)
        .Font.Italic = True
    End With
    
    ' ==== 4. 버튼 생성 ====
    ' STRIX 대화창 버튼
    Set btn = ws.Buttons.Add(ws.Range("B10").Left, ws.Range("B10").Top, 100, 30)
    With btn
        .Name = "btnSTRIXDialog"
        .Caption = "💬 STRIX 대화창"
        .OnAction = "ShowSTRIXDialog"
    End With
    
    ' 검색 실행 버튼
    Set btn = ws.Buttons.Add(ws.Range("C10").Left + 10, ws.Range("C10").Top, 100, 30)
    With btn
        .Name = "btnSearch"
        .Caption = "🔎 검색 실행"
        .OnAction = "ExecuteSearch"
    End With
    
    ' 선택 분석 버튼
    Set btn = ws.Buttons.Add(ws.Range("D10").Left + 20, ws.Range("D10").Top, 100, 30)
    With btn
        .Name = "btnAnalyze"
        .Caption = "📊 선택 분석"
        .OnAction = "AskAboutSelection"
    End With
    
    ' 문서 업로드 버튼
    Set btn = ws.Buttons.Add(ws.Range("E10").Left + 30, ws.Range("E10").Top, 100, 30)
    With btn
        .Name = "btnUpload"
        .Caption = "📤 문서 업로드"
        .OnAction = "BulkUploadDocuments"
    End With
    
    ' ==== 5. 답변 표시 영역 ====
    With ws.Range("B13:F25")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    
    ' 답변 영역 제목
    With ws.Range("B13:F13")
        .Merge
        .Value = "📝 답변 및 분석 결과"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(46, 204, 113)  ' 녹색
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 답변 표시 영역
    With ws.Range("B15:F24")
        .Merge
        .Name = "AnswerDisplay"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Value = "답변이 여기에 표시됩니다..."
        .Font.Color = RGB(150, 150, 150)
        .Font.Italic = True
    End With
    
    ' ==== 6. 검색 기록 패널 ====
    With ws.Range("H6:J20")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    
    ' 검색 기록 제목
    With ws.Range("H6:J6")
        .Merge
        .Value = "📋 최근 검색 기록"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(155, 89, 182)  ' 보라색
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 검색 기록 헤더
    ws.Range("H8").Value = "시간"
    ws.Range("I8").Value = "질문"
    ws.Range("J8").Value = "결과"
    With ws.Range("H8:J8")
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 새로고침 버튼
    Set btn = ws.Buttons.Add(ws.Range("H22").Left, ws.Range("H22").Top, 260, 25)
    With btn
        .Name = "btnRefreshHistory"
        .Caption = "🔄 검색 기록 새로고침"
        .OnAction = "ShowRecentSearches"
    End With
    
    ' ==== 7. 상태 표시 영역 ====
    With ws.Range("B27:J28")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    
    ' 상태 표시
    With ws.Range("B27:J27")
        .Merge
        .Name = "StatusBar"
        .Value = "✅ 준비 완료 - API 서버가 실행 중인지 확인하세요 (http://localhost:5000)"
        .Font.Size = 11
        .Font.Color = RGB(46, 204, 113)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 버전 정보
    With ws.Range("B28:J28")
        .Merge
        .Value = "STRIX v1.0 | AI-Powered Intelligence System | © 2024"
        .Font.Size = 9
        .Font.Color = RGB(150, 150, 150)
        .HorizontalAlignment = xlCenter
    End With
    
    ' ==== 8. 빠른 질문 템플릿 ====
    With ws.Range("H24:J30")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    
    ' 템플릿 제목
    With ws.Range("H24:J24")
        .Merge
        .Value = "💡 빠른 질문"
        .Font.Size = 12
        .Font.Bold = True
        .Interior.Color = RGB(241, 196, 15)  ' 노란색
        .HorizontalAlignment = xlCenter
    End With
    
    ' 템플릿 버튼들
    Set btn = ws.Buttons.Add(ws.Range("H26").Left, ws.Range("H26").Top, 260, 20)
    With btn
        .Name = "btnTemplate1"
        .Caption = "전고체 배터리 개발 현황"
        .OnAction = "QuickQuestion1"
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("H27").Left, ws.Range("H27").Top + 5, 260, 20)
    With btn
        .Name = "btnTemplate2"
        .Caption = "최근 배터리 시장 동향"
        .OnAction = "QuickQuestion2"
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("H28").Left, ws.Range("H28").Top + 10, 260, 20)
    With btn
        .Name = "btnTemplate3"
        .Caption = "경쟁사 기술 개발 현황"
        .OnAction = "QuickQuestion3"
    End With
    
    ' 화면 보기 설정
    ws.Range("B2").Select
    ActiveWindow.Zoom = 100
    
    ' 완료 메시지
    MsgBox "STRIX Dashboard가 성공적으로 생성되었습니다!" & vbCrLf & vbCrLf & _
           "사용 전 확인사항:" & vbCrLf & _
           "1. API 서버 실행: py api_server.py" & vbCrLf & _
           "2. VBA 모듈이 모두 import 되었는지 확인" & vbCrLf & _
           "3. 참조 설정이 완료되었는지 확인", vbInformation, "STRIX Dashboard"
    
End Sub

' 검색 실행 함수
Sub ExecuteSearch()
    Dim question As String
    Dim answer As String
    
    ' 질문 입력 영역에서 텍스트 가져오기
    question = ThisWorkbook.Sheets("STRIX Dashboard").Range("QuestionInput").Value
    
    If question = "" Or question = "여기에 질문을 입력하세요..." Then
        MsgBox "질문을 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' 상태 표시
    ThisWorkbook.Sheets("STRIX Dashboard").Range("StatusBar").Value = "🔄 검색 중..."
    
    ' STRIX에 질문
    answer = AskSTRIX(question)
    
    ' 답변 표시
    With ThisWorkbook.Sheets("STRIX Dashboard").Range("AnswerDisplay")
        .Value = answer
        .Font.Color = RGB(0, 0, 0)
        .Font.Italic = False
    End With
    
    ' 상태 업데이트
    ThisWorkbook.Sheets("STRIX Dashboard").Range("StatusBar").Value = "✅ 검색 완료 - " & Now()
    
End Sub

' 빠른 질문 템플릿 함수들
Sub QuickQuestion1()
    ThisWorkbook.Sheets("STRIX Dashboard").Range("QuestionInput").Value = "전고체 배터리 개발 현황은?"
    ExecuteSearch
End Sub

Sub QuickQuestion2()
    ThisWorkbook.Sheets("STRIX Dashboard").Range("QuestionInput").Value = "최근 배터리 시장 동향은?"
    ExecuteSearch
End Sub

Sub QuickQuestion3()
    ThisWorkbook.Sheets("STRIX Dashboard").Range("QuestionInput").Value = "경쟁사의 기술 개발 현황은?"
    ExecuteSearch
End Sub

' 질문 입력 영역 클릭 시 placeholder 제거
Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Intersect(Target, Range("QuestionInput")) Is Nothing Then
        If Range("QuestionInput").Value = "여기에 질문을 입력하세요..." Then
            Range("QuestionInput").Value = ""
            Range("QuestionInput").Font.Color = RGB(0, 0, 0)
            Range("QuestionInput").Font.Italic = False
        End If
    End If
End Sub