' Executive Dashboard for STRIX - Enhanced with Larger AI Results Display
Option Explicit

Private Const INTERNAL_COLOR As Long = 12611584  ' RGB(255, 192, 192) - 연한 빨간색
Private Const EXTERNAL_COLOR As Long = 13421619  ' RGB(179, 204, 255) - 연한 파란색
Private Const INTERNAL_ACCENT As Long = 255      ' RGB(255, 0, 0) - 빨간색
Private Const EXTERNAL_ACCENT As Long = 12611584 ' RGB(0, 112, 192) - 파란색

Public Sub CreateExecutiveDashboard()
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
    ws.Cells.Interior.color = RGB(250, 250, 250)
    
    ' 열 너비 설정 (AI 결과 창을 위해 확장)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C:D").ColumnWidth = 20
    ws.Columns("E:F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 15
    ws.Columns("H").ColumnWidth = 12
    ws.Columns("I:J").ColumnWidth = 15
    ws.Columns("K").ColumnWidth = 2
    
    ' ===== 1. 헤더 =====
    With ws.Range("B2:J2")
        .Merge
        .value = "STRIX Executive Intelligence Dashboard"
        .Font.Name = "맑은 고딕"
        .Font.Size = 24
        .Font.Bold = True
        .Interior.color = RGB(68, 114, 196)
        .Font.color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 45
    End With
    
    ' 부제목
    With ws.Range("B3:J3")
        .Merge
        .value = "AI 기반 통합 정보 분석 시스템"
        .Font.Size = 13
        .Font.color = RGB(80, 80, 80)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' ===== 2. 검색 설정 영역 =====
    ' 질문 입력 배경 강조
    With ws.Range("B5:J6")
        .Interior.color = RGB(245, 250, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.color = RGB(68, 114, 196)
    End With
    
    ' 질문 레이블
    ws.Range("B5").value = "질문:"
    ws.Range("B5").Font.Bold = True
    ws.Range("B5").Font.Size = 14
    ws.Range("B5").Font.color = RGB(68, 114, 196)
    
    ' 질문 입력 필드
    With ws.Range("C5:J6")
        .Merge
        .Name = "QuestionInput"
        .Interior.color = RGB(255, 250, 205)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.color = RGB(68, 114, 196)
        .Font.Size = 14
        .Font.Bold = False
        .value = "여기에 질문을 입력하세요"
        .Font.color = RGB(0, 0, 0)
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 35
    End With
    
    ws.Rows("5:6").RowHeight = 25
    
    ' ===== 3. 가중치 조절 슬라이더 영역 =====
    ws.Range("B8").value = "정보 소스 가중치:"
    ws.Range("B8").Font.Bold = True
    ws.Range("B8").Font.Size = 11
    
    ' 사내 문서 레이블
    ws.Range("C8").value = "사내"
    ws.Range("C8").Font.color = INTERNAL_ACCENT
    ws.Range("C8").Font.Bold = True
    ws.Range("C8").HorizontalAlignment = xlCenter
    
    ' 슬라이더 배경
    With ws.Range("D8:E8")
        .Merge
        .Name = "SliderArea"
        .Interior.color = RGB(240, 240, 240)
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
    
    ' 사외 문서 레이블
    ws.Range("F8").value = "사외"
    ws.Range("F8").Font.color = EXTERNAL_ACCENT
    ws.Range("F8").Font.Bold = True
    ws.Range("F8").HorizontalAlignment = xlCenter
    
    ' 가중치 퍼센트 표시
    ws.Range("G8").value = "50% / 50%"
    ws.Range("G8").Name = "WeightDisplay"
    ws.Range("G8").Font.Size = 11
    ws.Range("G8").HorizontalAlignment = xlCenter
    
    ' 검색 기간 선택
    ws.Range("H8").value = "기간:"
    ws.Range("H8").Font.Bold = True
    ws.Range("H8").Font.Size = 11
    
    With ws.Range("I8")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            Formula1:="최근 1개월,최근 3개월,최근 6개월,최근 1년,전체 기간"
        .value = "최근 3개월"
        .Interior.color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 11
    End With
    
    ' ===== 4. 메인 버튼 =====
    Set btn = ws.Buttons.Add(ws.Range("B10").Left, ws.Range("B10").Top, 120, 40)
    With btn
        .Caption = "AI 분석 실행"
        .OnAction = "ExecutiveRAGSearch"
        .Font.Size = 13
        .Font.Bold = True
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("D10").Left, ws.Range("D10").Top, 120, 40)
    With btn
        .Caption = "가중치 조절"
        .OnAction = "AdjustWeights"
        .Font.Size = 12
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("F10").Left, ws.Range("F10").Top, 120, 40)
    With btn
        .Caption = "초기화"
        .OnAction = "ResetDashboard"
        .Font.Size = 12
    End With
    
    ' 문서 유형 필터
    ws.Range("H10").value = "문서유형:"
    ws.Range("H10").Font.Bold = True
    ws.Range("H10").Font.Size = 11
    
    With ws.Range("I10")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            Formula1:="전체,보고서,회의록,뉴스,분석자료"
        .value = "전체"
        .Interior.color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 11
    End With
    
    ' ===== 5. 검색 진행 상태 표시 영역 =====
    With ws.Range("B12:J12")
        .Merge
        .Name = "SearchProgress"
        .Interior.color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
        .Font.Bold = True
        .value = "준비 완료"
        .Font.color = RGB(0, 150, 0)
    End With
    
    ' ===== 6. AI 분석 결과 영역 (크게 확대) =====
    With ws.Range("B14:J14")
        .Merge
        .value = "AI 분석 결과"
        .Font.Bold = True
        .Font.Size = 16
        .Interior.color = RGB(46, 204, 113)
        .Font.color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .RowHeight = 30
    End With
    
    ' AI 답변 표시 영역 (대폭 확대)
    With ws.Range("B15:J30")
        .Merge
        .Name = "AnswerArea"
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.color = RGB(46, 204, 113)
        .Font.Size = 12
        .value = "AI 분석 결과가 여기에 표시됩니다..." & vbNewLine & vbNewLine & _
                "• 질문을 입력하고 'AI 분석 실행' 버튼을 클릭하세요" & vbNewLine & _
                "• 가중치 조절로 사내/사외 정보 비중을 조정할 수 있습니다" & vbNewLine & _
                "• 참고 문서는 아래 테이블에 관련도 순으로 표시됩니다"
        .Font.color = RGB(150, 150, 150)
        .RowHeight = 25
    End With
    
    ' ===== 7. 참고 문서 영역 =====
    With ws.Range("B32:J32")
        .Merge
        .value = "참고 문서 (AI가 참조한 문서)"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.color = RGB(52, 152, 219)
        .Font.color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .RowHeight = 25
    End With
    
    ' 참고 문서 테이블 헤더
    ws.Range("B33").value = "번호"
    ws.Range("C33:D33").Merge
    ws.Range("C33").value = "제목"
    ws.Range("E33").value = "조직/출처"
    ws.Range("F33").value = "날짜"
    ws.Range("G33").value = "유형"
    ws.Range("H33").value = "문서유형"
    ws.Range("I33").value = "관련도"
    ws.Range("J33").value = "요약"
    
    With ws.Range("B33:J33")
        .Font.Bold = True
        .Font.Size = 11
        .Interior.color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 25
    End With
    
    ' 참고 문서 데이터 영역 서식
    Dim docRow As Integer
    For docRow = 34 To 53
        ' 번호 열
        ws.Cells(docRow, 2).HorizontalAlignment = xlCenter
        ws.Cells(docRow, 2).Borders.LineStyle = xlContinuous
        
        ' 제목 열 (병합)
        ws.Range(ws.Cells(docRow, 3), ws.Cells(docRow, 4)).Merge
        ws.Range(ws.Cells(docRow, 3), ws.Cells(docRow, 4)).WrapText = True
        ws.Range(ws.Cells(docRow, 3), ws.Cells(docRow, 4)).Borders.LineStyle = xlContinuous
        
        ' 나머지 열들
        ws.Cells(docRow, 5).HorizontalAlignment = xlCenter
        ws.Cells(docRow, 5).Borders.LineStyle = xlContinuous
        
        ws.Cells(docRow, 6).HorizontalAlignment = xlCenter
        ws.Cells(docRow, 6).Borders.LineStyle = xlContinuous
        
        ws.Cells(docRow, 7).HorizontalAlignment = xlCenter
        ws.Cells(docRow, 7).Borders.LineStyle = xlContinuous
        
        ws.Cells(docRow, 8).HorizontalAlignment = xlCenter
        ws.Cells(docRow, 8).Borders.LineStyle = xlContinuous
        
        ws.Cells(docRow, 9).HorizontalAlignment = xlCenter
        ws.Cells(docRow, 9).Borders.LineStyle = xlContinuous
        
        ws.Cells(docRow, 10).WrapText = True
        ws.Cells(docRow, 10).Borders.LineStyle = xlContinuous
    Next docRow
    
    With ws.Range("B34:J53")
        .Interior.color = RGB(255, 255, 255)
        .Font.Size = 10
        .RowHeight = 20
    End With
    
    ' ===== 8. 빠른 질문 =====
    ws.Range("B55").value = "빠른 질문:"
    ws.Range("B55").Font.Bold = True
    ws.Range("B55").Font.Size = 12
    
    ' 빠른 질문 버튼들
    Set btn = ws.Buttons.Add(ws.Range("B56").Left, ws.Range("B56").Top, 200, 30)
    With btn
        .Caption = "전고체 배터리 개발 현황"
        .OnAction = "QuickQuestion1"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("D56").Left + 20, ws.Range("D56").Top, 200, 30)
    With btn
        .Caption = "최근 배터리 시장 동향"
        .OnAction = "QuickQuestion2"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("F56").Left + 40, ws.Range("F56").Top, 200, 30)
    With btn
        .Caption = "경쟁사 기술 동향"
        .OnAction = "QuickQuestion3"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("H56").Left + 60, ws.Range("H56").Top, 200, 30)
    With btn
        .Caption = "ESG 규제 현황"
        .OnAction = "QuickQuestion4"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("B57").Left, ws.Range("B57").Top + 10, 200, 30)
    With btn
        .Caption = "원자재 가격 동향"
        .OnAction = "QuickQuestion5"
        .Font.Size = 11
    End With
    
    Set btn = ws.Buttons.Add(ws.Range("D57").Left + 20, ws.Range("D57").Top + 10, 200, 30)
    With btn
        .Caption = "글로벌 정책 변화"
        .OnAction = "QuickQuestion6"
        .Font.Size = 11
    End With
    
    ' 화면 설정
    ws.Range("B5").Select
    ActiveWindow.Zoom = 80
    ActiveWindow.DisplayGridlines = False
    
    MsgBox "Executive Dashboard가 생성되었습니다!" & Chr(10) & Chr(10) & _
           "주요 기능:" & Chr(10) & _
           "- AI 분석 결과가 큰 창에 표시됩니다" & Chr(10) & _
           "- 참고 문서가 테이블 형태로 정리됩니다" & Chr(10) & _
           "- 정보 소스 가중치 동적 조절" & Chr(10) & _
           "- 관련도 기반 문서 순위 표시", _
           vbInformation, "STRIX Executive Dashboard"
End Sub

' 가중치 조절 함수
Public Sub AdjustWeights()
    Dim ws As Worksheet
    Dim internalWeight As Integer
    Dim externalWeight As Integer
    Dim internalBar As Shape
    Dim externalBar As Shape
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 현재 가중치 가져오기
    Dim weightText As String
    weightText = ws.Range("WeightDisplay").value
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
    ws.Range("WeightDisplay").value = internalWeight & "% / " & externalWeight & "%"
    
    ' 상태 메시지
    If externalWeight > 50 Then
        ws.Range("SearchProgress").value = "사외 정보 중심 분석 모드"
        ws.Range("SearchProgress").Font.color = EXTERNAL_ACCENT
    ElseIf internalWeight > 50 Then
        ws.Range("SearchProgress").value = "사내 정보 중심 분석 모드"
        ws.Range("SearchProgress").Font.color = INTERNAL_ACCENT
    Else
        ws.Range("SearchProgress").value = "균형 분석 모드"
        ws.Range("SearchProgress").Font.color = RGB(0, 150, 0)
    End If
End Sub

' Executive RAG 검색 실행
Public Sub ExecutiveRAGSearch()
    Dim ws As Worksheet
    Dim question As String
    Dim searchProgress As Range
    Dim answerArea As Range
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Set searchProgress = ws.Range("SearchProgress")
    Set answerArea = ws.Range("AnswerArea")
    question = ws.Range("QuestionInput").value
    
    If question = "" Or question = "여기에 질문을 입력하세요" Then
        MsgBox "질문을 입력해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' STRIX Langchain RAG 분석 상태 표시
    searchProgress.value = "🔍 STRIX Langchain RAG 초기화 중..."
    searchProgress.Font.color = RGB(0, 100, 200)
    DoEvents
    Application.Wait Now + TimeValue("0:00:01")
    
    searchProgress.value = "📊 벡터 데이터베이스 검색 중... (Supabase pgvector)"
    DoEvents
    Application.Wait Now + TimeValue("0:00:01")
    
    searchProgress.value = "🤖 LLM 모델로 답변 생성 중... (GPT-4 Turbo)"
    searchProgress.Font.color = RGB(255, 140, 0)
    DoEvents
    Application.Wait Now + TimeValue("0:00:01")
    
    searchProgress.value = "📝 참고 문서 정리 및 검증 중..."
    DoEvents
    Application.Wait Now + TimeValue("0:00:01")
    
    ' Enhanced 시뮬레이션 답변 생성 (실제 데이터 기반)
    Dim simulatedAnswer As String
    simulatedAnswer = modEnhancedRAGSimulation.GenerateEnhancedAnswer(question)
    
    ' 답변 표시
    answerArea.value = simulatedAnswer
    answerArea.Font.color = RGB(0, 0, 0)
    answerArea.Font.Size = 12
    
    ' 답변 포맷팅
    Call FormatAnswerDisplay
    
    ' 참고 문서 테이블 채우기
    Call PopulateReferenceDocuments
    
    ' 완료 메시지
    searchProgress.value = "분석 완료 - " & Format(Now, "hh:mm:ss")
    searchProgress.Font.color = RGB(0, 150, 0)
End Sub

' AI 답변 표시 포맷팅
Public Sub FormatAnswerDisplay()
    Dim ws As Worksheet
    Dim answerArea As Range
    Dim answerText As String
    Dim i As Integer
    Dim startPos As Integer
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Set answerArea = ws.Range("AnswerArea")
    
    ' 답변 영역 스타일 개선
    With answerArea
        .Font.Name = "맑은 고딕"
        .Font.Size = 12
        .Font.color = RGB(0, 0, 0)
        .Interior.color = RGB(255, 255, 255)
        .Interior.Pattern = xlSolid
    End With
    
    ' 답변 내용이 있으면 포맷팅
    If answerArea.value <> "" And answerArea.value <> "AI 분석 결과가 여기에 표시됩니다..." Then
        answerText = answerArea.value
        
        ' 글머리 기호 처리
        answerText = Replace(answerText, "•", "◆")
        answerText = Replace(answerText, "-", "▶")
        
        answerArea.value = answerText
        
        ' 레퍼런스 번호 [1], [2] 등을 파란색 Bold로 포맷팅
        For i = 1 To 50
            Dim refPattern As String
            refPattern = "[" & i & "]"
            startPos = 1
            
            Do While InStr(startPos, answerText, refPattern) > 0
                Dim foundPos As Integer
                foundPos = InStr(startPos, answerText, refPattern)
                
                ' 찾은 위치의 텍스트를 파란색 Bold로 변경
                With answerArea.Characters(foundPos, Len(refPattern))
                    .Font.color = RGB(0, 112, 192)  ' 파란색
                    .Font.Bold = True
                    .Font.Size = 13  ' 약간 크게
                End With
                
                startPos = foundPos + Len(refPattern)
            Loop
        Next i
        
        ' 섹션 제목 (◆로 시작하는 줄) Bold 처리
        Dim lines() As String
        lines = Split(answerText, vbNewLine)
        Dim currentPos As Integer
        currentPos = 1
        
        For i = 0 To UBound(lines)
            If Left(lines(i), 1) = "◆" Then
                ' 해당 줄을 Bold로
                With answerArea.Characters(currentPos, Len(lines(i)))
                    .Font.Bold = True
                    .Font.Size = 13
                End With
            End If
            currentPos = currentPos + Len(lines(i)) + 2  ' vbNewLine 길이
        Next i
    End If
End Sub

' 참고 문서 테이블 채우기
Public Sub PopulateReferenceDocuments()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim i As Integer
    Dim docs As Collection
    Dim doc As Variant
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    startRow = 34
    
    ' 기존 데이터 지우기
    ws.Range("B34:J53").ClearContents
    
    ' 질문 내용을 가져와서 문서 유형 결정
    Dim questionText As String
    questionText = ws.Range("QuestionInput").value
    
    ' Enhanced 시뮬레이션에서 질문 유형에 맞는 문서 데이터 가져오기
    Set docs = modEnhancedRAGSimulation.GenerateReferenceDocuments(questionText)
    
    ' 데이터 입력
    For i = 1 To docs.Count
        If i > 20 Then Exit For ' 최대 20개 문서 표시
        
        Set doc = docs(i)
        
        ws.Cells(startRow + i - 1, 2).value = doc("num")  ' 번호
        ws.Cells(startRow + i - 1, 3).value = doc("title")  ' 제목
        ws.Cells(startRow + i - 1, 5).value = doc("org")  ' 조직/출처
        ws.Cells(startRow + i - 1, 6).value = doc("date")  ' 날짜
        ws.Cells(startRow + i - 1, 7).value = doc("type")  ' 유형
        ws.Cells(startRow + i - 1, 8).value = doc("docType")  ' 문서유형
        ws.Cells(startRow + i - 1, 9).value = doc("relevance")  ' 관련도
        ws.Cells(startRow + i - 1, 10).value = doc("summary") ' 요약
        
        ' 유형별 색상 코딩
        If doc("type") = "사내" Then
            ws.Cells(startRow + i - 1, 7).Interior.color = INTERNAL_COLOR
            ws.Cells(startRow + i - 1, 7).Font.color = INTERNAL_ACCENT
            ws.Cells(startRow + i - 1, 7).Font.Bold = True
        ElseIf doc("type") = "사외" Then
            ws.Cells(startRow + i - 1, 7).Interior.color = EXTERNAL_COLOR
            ws.Cells(startRow + i - 1, 7).Font.color = EXTERNAL_ACCENT
            ws.Cells(startRow + i - 1, 7).Font.Bold = True
        End If
        
        ' 관련도에 따른 색상 그라데이션
        Dim relevance As Integer
        relevance = Val(Replace(doc("relevance"), "%", ""))
        If relevance >= 90 Then
            ws.Cells(startRow + i - 1, 9).Font.color = RGB(0, 150, 0)
            ws.Cells(startRow + i - 1, 9).Font.Bold = True
        ElseIf relevance >= 80 Then
            ws.Cells(startRow + i - 1, 9).Font.color = RGB(0, 100, 200)
        ElseIf relevance >= 70 Then
            ws.Cells(startRow + i - 1, 9).Font.color = RGB(255, 140, 0)
        End If
        
        ' 행 포맷팅
        With ws.Range(ws.Cells(startRow + i - 1, 2), ws.Cells(startRow + i - 1, 10))
            If i Mod 2 = 0 Then
                .Interior.color = RGB(248, 248, 248)
            End If
        End With
    Next i
End Sub

' 빠른 질문들
Public Sub QuickQuestion1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").value = "전고체 배터리 기술 개발 현황과 상용화 전망은?"
    ws.Range("QuestionInput").Font.color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

Public Sub QuickQuestion2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").value = "최근 글로벌 배터리 시장 동향과 주요 이슈는?"
    ws.Range("QuestionInput").Font.color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

Public Sub QuickQuestion3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").value = "CATL, BYD 등 주요 경쟁사의 기술 개발 동향은?"
    ws.Range("QuestionInput").Font.color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

Public Sub QuickQuestion4()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").value = "ESG 및 탄소중립 규제가 배터리 산업에 미치는 영향은?"
    ws.Range("QuestionInput").Font.color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

Public Sub QuickQuestion5()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").value = "리튬, 니켈 등 주요 원자재 가격 동향과 전망은?"
    ws.Range("QuestionInput").Font.color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

Public Sub QuickQuestion6()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.Range("QuestionInput").value = "미국 IRA, 유럽 CBAM 등 글로벌 정책 변화의 영향은?"
    ws.Range("QuestionInput").Font.color = RGB(0, 0, 0)
    Call ExecutiveRAGSearch
End Sub

' Dashboard 초기화
Public Sub ResetDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 질문 초기화
    ws.Range("QuestionInput").value = "여기에 질문을 입력하세요"
    ws.Range("QuestionInput").Font.color = RGB(0, 0, 0)
    
    ' 답변 초기화
    ws.Range("AnswerArea").value = "AI 분석 결과가 여기에 표시됩니다..." & vbNewLine & vbNewLine & _
            "• 질문을 입력하고 'AI 분석 실행' 버튼을 클릭하세요" & vbNewLine & _
            "• 가중치 조절로 사내/사외 정보 비중을 조정할 수 있습니다" & vbNewLine & _
            "• 참고 문서는 아래 테이블에 관련도 순으로 표시됩니다"
    ws.Range("AnswerArea").Font.color = RGB(150, 150, 150)
    
    ' 참고 문서 초기화
    ws.Range("B34:J53").ClearContents
    
    ' 가중치 초기화 (50:50)
    Dim internalBar As Shape
    Dim externalBar As Shape
    Dim totalWidth As Double
    
    Set internalBar = ws.Shapes("InternalWeightBar")
    Set externalBar = ws.Shapes("ExternalWeightBar")
    totalWidth = ws.Range("D8:E8").Width - 4
    
    internalBar.Width = totalWidth * 0.5
    externalBar.Width = totalWidth * 0.5
    externalBar.Left = ws.Range("D8").Left + 2 + internalBar.Width
    
    ws.Range("WeightDisplay").value = "50% / 50%"
    
    ' 상태 초기화
    ws.Range("SearchProgress").value = "준비 완료"
    ws.Range("SearchProgress").Font.color = RGB(0, 150, 0)
    
    ' 기간 및 문서유형 초기화
    ws.Range("I8").value = "최근 3개월"
    ws.Range("I10").value = "전체"
    
    MsgBox "Dashboard가 초기화되었습니다.", vbInformation
End Sub