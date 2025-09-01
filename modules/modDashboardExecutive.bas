Attribute VB_Name = "modDashboardExecutive"
' Executive Dashboard for STRIX - Enhanced for Demo
Option Explicit

Private Const INTERNAL_COLOR As Long = 12611584  ' RGB(255, 192, 192) - 연한 빨간색
Private Const EXTERNAL_COLOR As Long = 13421619  ' RGB(179, 204, 255) - 연한 파란색
Private Const INTERNAL_ACCENT As Long = 255      ' RGB(255, 0, 0) - 빨간색
Private Const EXTERNAL_ACCENT As Long = 12611584 ' RGB(0, 112, 192) - 파란색

Sub CreateExecutiveDashboard()
    Dim ws As Worksheet
    Dim btn As Object
    Dim shp As Shape
    Dim dd As DropDown
    
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
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C:D").ColumnWidth = 18
    ws.Columns("E:F").ColumnWidth = 15
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 2
    
    ' ===== 1. 헤더 =====
    With ws.Range("B2:H2")
        .Merge
        .Value = "STRIX Executive Intelligence Dashboard"
        .Font.Name = "맑은 고딕"
        .Font.Size = 24
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 45
    End With
    
    ' 부제목
    With ws.Range("B3:H3")
        .Merge
        .Value = "AI 기반 통합 정보 분석 시스템"
        .Font.Size = 13
        .Font.Color = RGB(80, 80, 80)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' ===== 2. 검색 설정 영역 =====
    ' 질문 입력 배경 강조
    With ws.Range("B5:H6")
        .Interior.Color = RGB(245, 250, 255)  ' 연한 파란색 배경
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(68, 114, 196)  ' 진한 파란색 테두리
    End With
    
    ' 질문 레이블
    ws.Range("B5").Value = "질문:"
    ws.Range("B5").Font.Bold = True
    ws.Range("B5").Font.Size = 14
    ws.Range("B5").Font.Color = RGB(68, 114, 196)
    
    ' 질문 입력 필드
    With ws.Range("C5:H6")
        .Merge
        .Name = "QuestionInput"
        .Interior.Color = RGB(255, 250, 205)  ' 밝은 노란색 배경
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(68, 114, 196)
        .Font.Size = 14
        .Font.Bold = False
        .Value = "여기에 질문을 입력하세요"
        .Font.Color = RGB(0, 0, 0)  ' 검은색 텍스트
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 35
    End With
    
    ' 행 높이 조정
    ws.Rows("5:6").RowHeight = 25
    
    ' ===== 3. 가중치 조절 슬라이더 영역 =====
    ws.Range("B8").Value = "정보 소스 가중치:"
    ws.Range("B8").Font.Bold = True
    ws.Range("B8").Font.Size = 11
    
    ' 사내 문서 레이블
    ws.Range("C8").Value = "사내"
    ws.Range("C8").Font.Color = INTERNAL_ACCENT
    ws.Range("C8").Font.Bold = True
    ws.Range("C8").HorizontalAlignment = xlCenter
    
    ' 슬라이더 배경
    With ws.Range("D8:E8")
        .Merge
        .Name = "SliderArea"
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
        .RowHeight = 25
    End With
    
    ' 사내 가중치 바 (빨간색)
    Dim totalWidth As Double
    totalWidth = ws.Range("D8:E8").Width - 4
    
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        ws.Range("D8").Left + 2, _
        ws.Range("D8").Top + 5, _
        totalWidth * 0.5, 15)
    With shp
        .Name = "InternalWeightBar"
        .Fill.ForeColor.RGB = RGB(255, 100, 100)
        .Line.Visible = msoFalse
    End With
    
    ' 사외 가중치 바 (파란색)
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        ws.Range("D8").Left + 2 + (totalWidth * 0.5), _
        ws.Range("D8").Top + 5, _
        totalWidth * 0.5, 15)
    With shp
        .Name = "ExternalWeightBar"
        .Fill.ForeColor.RGB = RGB(100, 150, 255)
        .Line.Visible = msoFalse
    End With
    
    ws.Range("F8").Clear  ' F8 영역 비우기
    
    ' 사외 문서 레이블
    ws.Range("F8").Value = "사외"
    ws.Range("F8").Font.Color = EXTERNAL_ACCENT
    ws.Range("F8").Font.Bold = True
    ws.Range("F8").HorizontalAlignment = xlCenter
    
    ' 가중치 퍼센트 표시
    ws.Range("G8").Value = "50% / 50%"
    ws.Range("G8").Name = "WeightDisplay"
    ws.Range("G8").Font.Size = 11
    ws.Range("G8").HorizontalAlignment = xlCenter
    
    ' ===== 4. 검색 기간 선택 =====
    ws.Range("B9").Value = "검색 기간:"
    ws.Range("B9").Font.Bold = True
    ws.Range("B9").Font.Size = 11
    
    ' 기간 드롭다운
    With ws.Range("C9")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            Formula1:="최근 1개월,최근 3개월,최근 6개월,최근 1년,전체 기간"
        .Value = "최근 3개월"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 11
    End With
    
    ' 문서 유형 필터
    ws.Range("E9").Value = "문서 유형:"
    ws.Range("E9").Font.Bold = True
    ws.Range("E9").Font.Size = 11
    
    With ws.Range("F9")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            Formula1:="전체,보고서,회의록,뉴스,분석자료"
        .Value = "전체"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 11
    End With
    
    ' ===== 5. 메인 버튼 =====
    Set btn = ws.Buttons.Add(ws.Range("B11").Left, ws.Range("B11").Top, 120, 40)
    With btn
        .Caption = "AI 분석 실행"
        .OnAction = "ExecutiveRAGSearch"
        .Font.Size = 13
        .Font.Bold = True
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("D11").Left, ws.Range("D11").Top, 120, 40)
    With btn
        .Caption = "가중치 조절"
        .OnAction = "AdjustWeights"
        .Font.Size = 12
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("F11").Left, ws.Range("F11").Top, 120, 40)
    With btn
        .Caption = "초기화"
        .OnAction = "ResetDashboard"
        .Font.Size = 12
    End With
    
    ' ===== 6. 검색 진행 상태 표시 영역 =====
    With ws.Range("B13:H13")
        .Merge
        .Name = "SearchProgress"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
        .Font.Bold = True
        .Value = "준비 완료"
        .Font.Color = RGB(0, 150, 0)
    End With
    
    ' ===== 7. 답변 영역 =====
    With ws.Range("B15:H15")
        .Merge
        .Value = "AI 분석 결과"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(46, 204, 113)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 답변 표시 영역
    With ws.Range("B16:H26")
        .Merge
        .Name = "AnswerArea"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .Font.Size = 11
        .Value = "AI 분석 결과가 여기에 표시됩니다..."
        .Font.Color = RGB(150, 150, 150)
    End With
    
    ' ===== 8. 참고 문서 영역 =====
    With ws.Range("B27:H27")
        .Merge
        .Value = "참고 문서 (AI가 참조한 문서)"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(52, 152, 219)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 참고 문서 테이블 헤더
    ws.Range("B28").Value = "번호"
    ws.Range("C28").Value = "제목"
    ws.Range("D28").Value = "조직/출처"
    ws.Range("E28").Value = "날짜"
    ws.Range("F28").Value = "유형"
    ws.Range("G28").Value = "문서유형"
    ws.Range("H28").Value = "관련도"
    
    With ws.Range("B28:H28")
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    ' 참고 문서 영역 서식
    With ws.Range("B29:H50")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .Font.Size = 10
    End With
    
    ' ===== 9. 빠른 질문 =====
    ws.Range("B52").Value = "빠른 질문:"
    ws.Range("B52").Font.Bold = True
    ws.Range("B52").Font.Size = 12
    
    ' 일반적인 질문 버튼들
    Set btn = ws.Buttons.Add(ws.Range("B53").Left, ws.Range("B53").Top, 180, 30)
    With btn
        .Caption = "전고체 배터리 개발 현황"
        .OnAction = "QuickQuestion1"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("D53").Left + 20, ws.Range("D53").Top, 180, 30)
    With btn
        .Caption = "최근 배터리 시장 동향"
        .OnAction = "QuickQuestion2"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("F53").Left + 40, ws.Range("F53").Top, 180, 30)
    With btn
        .Caption = "경쟁사 기술 동향"
        .OnAction = "QuickQuestion3"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("B54").Left, ws.Range("B54").Top + 10, 180, 30)
    With btn
        .Caption = "ESG 규제 현황"
        .OnAction = "QuickQuestion4"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("D54").Left + 20, ws.Range("D54").Top + 10, 180, 30)
    With btn
        .Caption = "원자재 가격 동향"
        .OnAction = "QuickQuestion5"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("F54").Left + 40, ws.Range("F54").Top + 10, 180, 30)
    With btn
        .Caption = "글로벌 정책 변화"
        .OnAction = "QuickQuestion6"
        .Font.Size = 11
    End With
    
    ' 하단 여백
    
    ' 화면 설정
    ws.Range("B5").Select
    ActiveWindow.Zoom = 85
    ActiveWindow.DisplayGridlines = False
    
    MsgBox "Executive Dashboard가 생성되었습니다!" & Chr(10) & Chr(10) & _
           "주요 기능:" & Chr(10) & _
           "- 정보 소스 가중치 동적 조절" & Chr(10) & _
           "- 검색 기간 및 문서 유형 필터링" & Chr(10) & _
           "- AI 분석 결과 및 참고 문서 표시" & Chr(10) & _
           "- 관련도 기반 문서 순위 표시", _
           vbInformation, "STRIX Executive Dashboard"
End Sub

' 가중치 조절 함수
Sub AdjustWeights()
    Dim ws As Worksheet
    Dim internalWeight As Integer
    Dim externalWeight As Integer
    Dim internalBar As Shape
    Dim externalBar As Shape
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 현재 가중치 가져오기
    Dim weightText As String
    weightText = ws.Range("WeightDisplay").Value
    internalWeight = Val(Split(weightText, "/")(0))
    
    ' 가중치 10% 단위로 조정
    internalWeight = internalWeight - 10
    If internalWeight < 10 Then
        internalWeight = 90
    End If
    externalWeight = 100 - internalWeight
    
    ' 바 크기 조절
    Set internalBar = ws.Shapes("InternalWeightBar")
    Set externalBar = ws.Shapes("ExternalWeightBar")
    
    Dim totalWidth As Double
    totalWidth = ws.Range("D8:E8").Width - 4
    
    internalBar.Width = totalWidth * (internalWeight / 100)
    externalBar.Width = totalWidth * (externalWeight / 100)
    externalBar.Left = ws.Range("D8").Left + 2 + internalBar.Width
    
    ' 표시 업데이트
    ws.Range("WeightDisplay").Value = internalWeight & "% / " & externalWeight & "%"
    
    ' 상태 메시지
    If externalWeight > 50 Then
        ws.Range("SearchProgress").Value = "사외 정보 중심 분석 모드"
        ws.Range("SearchProgress").Font.Color = EXTERNAL_ACCENT
    ElseIf internalWeight > 50 Then
        ws.Range("SearchProgress").Value = "사내 정보 중심 분석 모드"
        ws.Range("SearchProgress").Font.Color = INTERNAL_ACCENT
    Else
        ws.Range("SearchProgress").Value = "균형 분석 모드"
        ws.Range("SearchProgress").Font.Color = RGB(0, 150, 0)
    End If
End Sub

' Executive RAG 검색 실행
Sub ExecutiveRAGSearch()
    Dim ws As Worksheet
    Dim question As String
    Dim searchProgress As Range
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Set searchProgress = ws.Range("SearchProgress")
    question = ws.Range("C5").Value
    
    If question = "" Or question = "질문을 입력하세요" Then
        MsgBox "질문을 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' 검색 진행 상태 애니메이션
    For i = 1 To 3
        searchProgress.Value = "AI 분석 중" & String(i, ".")
        searchProgress.Font.Color = RGB(255, 140, 0)
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
    Next i
    
    searchProgress.Value = "문서 검색 및 분석 중..."
    DoEvents
    
    ' 실제 RAG API 호출
    Call modRAGAPI.RunRAGSearchWithSources
    
    ' 답변 포맷팅 (참조 번호 강조)
    Call FormatAnswerReferences
    
    ' 참고 문서 색상 코딩
    Call ColorCodeDocuments
    
    ' 완료 메시지
    searchProgress.Value = "분석 완료 - " & Format(Now, "hh:mm:ss")
    searchProgress.Font.Color = RGB(0, 150, 0)
End Sub

' 답변의 참조 번호 [1], [2] 등을 파란색 Bold로 포맷팅
Sub FormatAnswerReferences()
    Dim ws As Worksheet
    Dim answerCell As Range
    Dim answerText As String
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Set answerCell = ws.Range("AnswerArea")
    answerText = answerCell.Value
    
    ' 정규식으로 [숫자] 패턴 찾아서 포맷팅
    For i = 1 To 30
        Dim refPattern As String
        refPattern = "[" & i & "]"
        If InStr(answerText, refPattern) > 0 Then
            ' 참조 번호를 찾으면 해당 부분 포맷팅
            ' VBA에서는 부분 포맷팅이 제한적이므로 전체 텍스트 색상 변경
            answerCell.Font.Color = RGB(0, 0, 0)
        End If
    Next i
End Sub

' 참고 문서 색상 코딩
Sub ColorCodeDocuments()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    lastRow = 50 ' 최대 표시 행
    
    For i = 29 To lastRow
        If ws.Cells(i, 6).Value = "사내" Or ws.Cells(i, 6).Value = "internal" Then
            ws.Cells(i, 6).Interior.Color = INTERNAL_COLOR
            ws.Cells(i, 6).Font.Color = INTERNAL_ACCENT
            ws.Cells(i, 6).Font.Bold = True
        ElseIf ws.Cells(i, 6).Value = "사외" Or ws.Cells(i, 6).Value = "external" Then
            ws.Cells(i, 6).Interior.Color = EXTERNAL_COLOR
            ws.Cells(i, 6).Font.Color = EXTERNAL_ACCENT
            ws.Cells(i, 6).Font.Bold = True
        End If
        
        ' 문서 유형 추가
        If ws.Cells(i, 6).Value = "사내" Or ws.Cells(i, 6).Value = "internal" Then
            ' 제목에서 문서 유형 추론
            Dim title As String
            title = ws.Cells(i, 3).Value
            If InStr(title, "보고") > 0 Then
                ws.Cells(i, 7).Value = "보고서"
            ElseIf InStr(title, "회의") > 0 Then
                ws.Cells(i, 7).Value = "회의록"
            ElseIf InStr(title, "분석") > 0 Then
                ws.Cells(i, 7).Value = "분석자료"
            ElseIf InStr(title, "전략") > 0 Then
                ws.Cells(i, 7).Value = "전략문서"
            Else
                ws.Cells(i, 7).Value = "일반문서"
            End If
        ElseIf ws.Cells(i, 6).Value = "사외" Or ws.Cells(i, 6).Value = "external" Then
            If InStr(ws.Cells(i, 3).Value, "뉴스") > 0 Or InStr(ws.Cells(i, 3).Value, "속보") > 0 Then
                ws.Cells(i, 7).Value = "뉴스"
            ElseIf InStr(ws.Cells(i, 3).Value, "리포트") > 0 Then
                ws.Cells(i, 7).Value = "리포트"
            Else
                ws.Cells(i, 7).Value = "외부자료"
            End If
        End If
        
        ' 관련도 추출 및 표시
        If InStr(title, "(") > 0 And InStr(title, "%)") > 0 Then
            Dim startPos As Integer, endPos As Integer
            startPos = InStrRev(title, "(")
            endPos = InStr(startPos, title, "%)")
            If startPos > 0 And endPos > 0 Then
                Dim relevance As String
                relevance = Mid(title, startPos + 1, endPos - startPos)
                ws.Cells(i, 8).Value = relevance
                ws.Cells(i, 8).HorizontalAlignment = xlCenter
                ' 제목에서 관련도 제거
                ws.Cells(i, 3).Value = Left(title, startPos - 2)
            End If
        End If
    Next i
End Sub

' 빠른 질문들
Sub QuickQuestion1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").Value = "전고체 배터리 기술 개발 현황과 상용화 전망은?"
    ws.Range("QuestionInput").Font.Color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

Sub QuickQuestion2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").Value = "최근 글로벌 배터리 시장 동향과 주요 이슈는?"
    ws.Range("QuestionInput").Font.Color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

Sub QuickQuestion3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").Value = "CATL, BYD 등 주요 경쟁사의 기술 개발 동향은?"
    ws.Range("QuestionInput").Font.Color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

Sub QuickQuestion4()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").Value = "ESG 및 탄소중립 규제가 배터리 산업에 미치는 영향은?"
    ws.Range("QuestionInput").Font.Color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

Sub QuickQuestion5()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").Value = "리튬, 니켈 등 주요 원자재 가격 동향과 전망은?"
    ws.Range("QuestionInput").Font.Color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

Sub QuickQuestion6()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").Value = "미국 IRA, 유럽 CBAM 등 글로벌 정책 변화의 영향은?"
    ws.Range("QuestionInput").Font.Color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub


' Dashboard 초기화
Sub ResetDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 질문 초기화
    ws.Range("C5").Value = "여기에 질문을 입력하세요"
    ws.Range("C5").Font.Color = RGB(0, 0, 0)  ' 검은색
    
    ' 답변 초기화
    ws.Range("AnswerArea").Value = "AI 분석 결과가 여기에 표시됩니다..."
    ws.Range("AnswerArea").Font.Color = RGB(100, 100, 100)  ' 진한 회색
    
    ' 참고 문서 초기화
    ws.Range("B29:H50").Clear
    With ws.Range("B29:H50")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    
    ' 가중치 초기화 (50:50)
    Dim internalBar As Shape
    Dim externalBar As Shape
    Set internalBar = ws.Shapes("InternalWeightBar")
    Set externalBar = ws.Shapes("ExternalWeightBar")
    
    internalBar.Width = totalWidth * 0.5
    externalBar.Width = totalWidth * 0.5
    externalBar.Left = ws.Range("D8").Left + 2 + internalBar.Width
    
    ws.Range("WeightDisplay").Value = "50% / 50%"
    
    ' 상태 초기화
    ws.Range("SearchProgress").Value = "준비 완료"
    ws.Range("SearchProgress").Font.Color = RGB(0, 150, 0)
    
    ' 기간 초기화
    ws.Range("C8").Value = "최근 3개월"
    ws.Range("F8").Value = "전체"
End Sub