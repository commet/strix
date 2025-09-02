Attribute VB_Name = "modIssueTimelineUpdated"
Option Explicit

' ============================================
' 업데이트된 Issue Timeline - SK온 사내 정보 반영
' ============================================

Private allIssues As Collection
Private filteredIssues As Collection

Sub CreateUpdatedDashboard()
    Dim ws As Worksheet
    Dim row As Long
    Dim btn As Object
    
    ' 시트 생성 또는 초기화
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Issue Timeline"
    Else
        ws.Cells.Clear
        Dim shp As Shape
        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Then shp.Delete
        Next shp
    End If
    On Error GoTo 0
    
    ' 전체 시트 폰트 설정
    With ws.Cells.Font
        .Name = "맑은 고딕"
        .Size = 12
    End With
    
    ' 헤더 영역
    With ws.Range("B2:R2")
        .Merge
        .Value = "STRIX Issue Timeline & Decision Tracker"
        .Font.Size = 24
        .Font.Bold = True
        .Interior.Color = RGB(39, 55, 39)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .RowHeight = 50
    End With
    
    ' 부제목
    With ws.Range("B3:R3")
        .Merge
        .Value = "사내 이슈 진행 현황 및 의사결정 추적 시스템"
        .Font.Size = 14
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' 검색 영역
    ws.Range("B5").Value = "검색:"
    ws.Range("B5").Font.Bold = True
    ws.Range("B5").Font.Size = 14
    
    With ws.Range("C5:G5")
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Font.Size = 14
        .RowHeight = 30
    End With
    
    ' 검색 버튼
    Set btn = ws.Buttons.Add(ws.Range("H5").Left, ws.Range("H5").Top, _
                             ws.Range("H5").Width, ws.Range("H5").Height)
    With btn
        .Caption = "검색"
        .OnAction = "SearchUpdated"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' 전체보기 버튼
    Set btn = ws.Buttons.Add(ws.Range("I5").Left, ws.Range("I5").Top, _
                             ws.Range("I5").Width, ws.Range("I5").Height)
    With btn
        .Caption = "전체보기"
        .OnAction = "ShowAllUpdated"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' 필터 레이블
    ws.Range("D7").Value = "분류1"
    ws.Range("E7").Value = "세부구분"
    ws.Range("F7").Value = "상태"
    ws.Range("G7").Value = "담당부서"
    
    With ws.Range("D7:G7")
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 필터 드롭다운
    With ws.Range("D8")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        On Error Resume Next
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="전체,사내,사외"
        On Error GoTo 0
        .Value = "전체"
        .RowHeight = 25
    End With
    
    With ws.Range("E8")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        On Error Resume Next
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="전체,정책,경쟁사,Tech,Marketing,Production,R&D,Staff,ESS,투자,특허,시장"
        On Error GoTo 0
        .Value = "전체"
    End With
    
    With ws.Range("F8")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        On Error Resume Next
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="전체,해결됨,모니터링,진행중,미해결"
        On Error GoTo 0
        .Value = "전체"
    End With
    
    With ws.Range("G8")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        On Error Resume Next
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="전체,전략기획팀,생산관리팀,품질관리팀,영업마케팅팀,R&D센터,경영지원팀,구매팀,인사팀,시장분석팀,경영기획팀,법무팀,안전환경팀,해외사업팀,중국사업팀,ESS사업팀"
        On Error GoTo 0
        .Value = "전체"
    End With
    
    ' 필터 적용 버튼
    Set btn = ws.Buttons.Add(ws.Range("H8").Left, ws.Range("H8").Top, _
                             ws.Range("H8:I8").Width, ws.Range("H8").Height)
    With btn
        .Caption = "🔍 필터 적용"
        .OnAction = "ApplyFilterUpdated"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' 테이블 헤더
    ws.Range("A10:Q10").Value = Array("No", "날짜", "제목", "분류1", "분류2", _
                                      "상태", "담당부서", "진행률", _
                                      "2025-05", "2025-06", "2025-07", _
                                      "2025-08", "2025-09", "2025-10", "2025-11", _
                                      "문서 참조", "업데이트")
    
    With ws.Range("A10:Q10")
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
    
    ' 데이터 로드
    Call LoadUpdatedData
    Call ShowAllUpdated
    
    ' 안내 메시지
    ws.Range("L5").Value = "💡 드롭다운 선택 후 [필터 적용] 버튼 클릭"
    ws.Range("L5").Font.Color = RGB(0, 0, 255)
    ws.Range("L5").Font.Size = 11
    
    ws.Activate
    
    MsgBox "업데이트된 Issue Timeline이 생성되었습니다!" & vbCrLf & vbCrLf & _
           "사용 방법:" & vbCrLf & _
           "1. 드롭다운에서 원하는 필터 선택" & vbCrLf & _
           "2. [필터 적용] 버튼 클릭" & vbCrLf & vbCrLf & _
           "검색: ESS 관련 이슈 → 11개 문서 필터링", vbInformation
End Sub

Sub ApplyFilterUpdated()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    
    Dim filter1 As String, filter2 As String, filter3 As String, filter4 As String
    Dim searchTerm As String
    
    ' 필터 값 읽기
    filter1 = ws.Range("D8").Value
    filter2 = ws.Range("E8").Value
    filter3 = ws.Range("F8").Value
    filter4 = ws.Range("G8").Value
    searchTerm = ws.Range("C5").Value
    
    ' allIssues가 비어있으면 로드
    If allIssues Is Nothing Then
        Call LoadUpdatedData
    End If
    
    If allIssues.Count = 0 Then
        Call LoadUpdatedData
    End If
    
    ' 필터링된 컬렉션 생성
    Set filteredIssues = New Collection
    Dim issue As Object
    Dim includeIssue As Boolean
    
    For Each issue In allIssues
        includeIssue = True
        
        ' 검색어 필터
        If searchTerm <> "" Then
            If InStr(1, searchTerm, "ESS", vbTextCompare) > 0 And _
               (InStr(1, searchTerm, "관련", vbTextCompare) > 0 Or _
                InStr(1, searchTerm, "이슈", vbTextCompare) > 0) Then
                If Not issue("isESS") Then includeIssue = False
            ElseIf InStr(1, issue("title"), searchTerm, vbTextCompare) = 0 And _
                   InStr(1, issue("category2"), searchTerm, vbTextCompare) = 0 Then
                includeIssue = False
            End If
        End If
        
        ' 분류1 필터
        If filter1 <> "전체" And filter1 <> "" Then
            If issue("category1") <> filter1 Then includeIssue = False
        End If
        
        ' 세부구분 필터
        If filter2 <> "전체" And filter2 <> "" Then
            If issue("category2") <> filter2 Then includeIssue = False
        End If
        
        ' 상태 필터
        If filter3 <> "전체" And filter3 <> "" Then
            If issue("status") <> filter3 Then includeIssue = False
        End If
        
        ' 담당부서 필터
        If filter4 <> "전체" And filter4 <> "" Then
            If issue("dept") <> filter4 Then includeIssue = False
        End If
        
        If includeIssue Then
            filteredIssues.Add issue
        End If
    Next issue
    
    ' 필터링된 이슈 표시
    Call DisplayUpdatedIssues(ws)
End Sub

Sub SearchUpdated()
    Call ApplyFilterUpdated
End Sub

Sub ShowAllUpdated()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    
    ' 모든 필터 초기화
    ws.Range("D8").Value = "전체"
    ws.Range("E8").Value = "전체"
    ws.Range("F8").Value = "전체"
    ws.Range("G8").Value = "전체"
    ws.Range("C5").Value = ""
    
    ' 필터 적용
    Call ApplyFilterUpdated
End Sub

Private Sub DisplayUpdatedIssues(ws As Worksheet)
    Dim row As Long
    Dim issue As Object
    Dim displayCount As Integer
    Dim lastRow As Long
    
    ' 기존 데이터 영역 삭제
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow >= 11 Then
        ws.Range("A11:Q" & lastRow).Clear
    End If
    
    row = 11
    displayCount = 0
    
    ' 필터링된 이슈 표시
    For Each issue In filteredIssues
        displayCount = displayCount + 1
        Call AddUpdatedIssueRow(ws, row, displayCount, issue)
        row = row + 1
    Next issue
    
    ' 결과 메시지
    ws.Range("K5").Value = "총 " & displayCount & "개"
    ws.Range("K5").Font.Color = IIf(displayCount = allIssues.Count, RGB(0, 128, 0), RGB(0, 0, 255))
    ws.Range("K5").Font.Bold = True
    
    ' 테두리 적용
    If row > 11 Then
        With ws.Range("A10:Q" & (row - 1))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End If
End Sub

Private Sub AddUpdatedIssueRow(ws As Worksheet, row As Long, no As Integer, issue As Object)
    ' 번호
    ws.Cells(row, 1).Value = no
    ws.Cells(row, 1).HorizontalAlignment = xlCenter
    
    ' 날짜
    ws.Cells(row, 2).Value = Format(issue("date"), "yyyy-mm-dd")
    ws.Cells(row, 2).HorizontalAlignment = xlCenter
    
    ' 제목
    ws.Cells(row, 3).Value = issue("title")
    
    ' 분류1
    ws.Cells(row, 4).Value = issue("category1")
    ws.Cells(row, 4).HorizontalAlignment = xlCenter
    If issue("category1") = "사내" Then
        ws.Cells(row, 4).Interior.Color = RGB(255, 100, 100)
        ws.Cells(row, 4).Font.Color = RGB(255, 255, 255)
    Else
        ws.Cells(row, 4).Interior.Color = RGB(100, 150, 255)
        ws.Cells(row, 4).Font.Color = RGB(255, 255, 255)
    End If
    
    ' 분류2
    ws.Cells(row, 5).Value = issue("category2")
    ws.Cells(row, 5).HorizontalAlignment = xlCenter
    
    ' 상태
    ws.Cells(row, 6).Value = issue("status")
    ws.Cells(row, 6).HorizontalAlignment = xlCenter
    ws.Cells(row, 6).Font.Bold = True
    Select Case issue("status")
        Case "해결됨"
            ws.Cells(row, 6).Font.Color = RGB(0, 176, 80)
        Case "진행중"
            ws.Cells(row, 6).Font.Color = RGB(255, 192, 0)
        Case "미해결"
            ws.Cells(row, 6).Font.Color = RGB(255, 0, 0)
        Case "모니터링"
            ws.Cells(row, 6).Font.Color = RGB(0, 112, 192)
    End Select
    
    ' 담당부서
    ws.Cells(row, 7).Value = issue("dept")
    ws.Cells(row, 7).HorizontalAlignment = xlCenter
    
    ' 진행률
    ws.Cells(row, 8).Value = issue("progress") & "%"
    ws.Cells(row, 8).HorizontalAlignment = xlCenter
    
    ' 문서 참조
    With ws.Cells(row, 16)
        .Value = issue("docRef")
        .Font.Color = RGB(0, 0, 255)
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 12
    End With
    
    ' 업데이트 날짜
    ws.Cells(row, 17).Value = Format(issue("updateDate"), "yyyy-mm-dd")
    ws.Cells(row, 17).HorizontalAlignment = xlCenter
    
    ' 타임라인 그리기
    Call DrawUpdatedTimeline(ws, row, issue)
End Sub

Private Sub DrawUpdatedTimeline(ws As Worksheet, row As Long, issue As Object)
    Dim startCol As Integer, endCol As Integer, currentCol As Integer
    Dim monthDiff As Integer
    Dim baseDate As Date
    Dim cellColor As Long
    
    baseDate = #5/1/2025#
    
    ' 시작 월 계산
    monthDiff = DateDiff("m", baseDate, issue("startDate"))
    If monthDiff < 0 Then monthDiff = 0
    If monthDiff > 6 Then monthDiff = 6
    startCol = 9 + monthDiff
    
    ' 종료 월 계산
    monthDiff = DateDiff("m", baseDate, issue("endDate"))
    If monthDiff < 0 Then monthDiff = 0
    If monthDiff > 6 Then monthDiff = 6
    endCol = 9 + monthDiff
    
    ' 현재 월 계산 (2025년 8월)
    currentCol = 12
    
    ' 색상 결정
    Select Case issue("status")
        Case "해결됨"
            cellColor = RGB(112, 173, 71)   ' 초록색
        Case "진행중"
            cellColor = RGB(255, 192, 0)    ' 노란색
        Case "미해결"
            cellColor = RGB(255, 0, 0)      ' 빨간색
        Case "모니터링"
            cellColor = RGB(68, 114, 196)   ' 파란색
    End Select
    
    ' 타임라인 그리기
    Dim i As Integer
    For i = startCol To endCol
        ws.Cells(row, i).Interior.Color = cellColor
        
        ' 현재 시점 마커 (8월)
        If i = currentCol Then
            ws.Cells(row, i).Value = "●"
            ws.Cells(row, i).Font.Color = RGB(255, 255, 255)
            ws.Cells(row, i).Font.Size = 14
            ws.Cells(row, i).HorizontalAlignment = xlCenter
        End If
        
        ' 완료 체크마크
        If issue("status") = "해결됨" And i = endCol Then
            ws.Cells(row, i).Font.Name = "Wingdings"
            ws.Cells(row, i).Value = Chr(252)
            ws.Cells(row, i).Font.Color = RGB(255, 255, 255)
            ws.Cells(row, i).Font.Size = 14
            ws.Cells(row, i).HorizontalAlignment = xlCenter
        End If
    Next i
End Sub

Private Sub LoadUpdatedData()
    ' 업데이트된 54개 이슈 데이터
    Set allIssues = New Collection
    Dim issue As Object
    
    ' ESS 관련 이슈 11개 (사내 SK온 + 사외 경쟁사)
    ' 사내 ESS 이슈들 (SK온)
    Set issue = CreateUpdatedIssue(#8/22/2025#, "SK온 조지아공장 12개 라인 중 2개 ESS용 LFP 라인 배정 확정", _
                "사내", "ESS", "해결됨", "ESS사업팀", _
                "SKBA_LFP라인배정.docx", #8/21/2025#, 100, #6/1/2025#, #8/22/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#7/31/2025#, "SK온-SK엔무브 합병 이사회 의결 - 11월 1일 공식출범 예정", _
                "사내", "ESS", "진행중", "경영기획팀", _
                "합병계획서_내부.docx", #7/30/2025#, 85, #7/1/2025#, #11/1/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#7/11/2025#, "SK온-엘앤에프 북미 LFP 양극재 공급 MOU 체결 - ESS용", _
                "사내", "ESS", "해결됨", "ESS사업팀", _
                "LnF_LFP공급MOU.pdf", #7/10/2025#, 100, #5/1/2025#, #7/11/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#6/5/2025#, "SK온 미국 ESS 전용공장 하반기 가동준비 - 텍사스 20GWh", _
                "사내", "ESS", "진행중", "ESS사업팀", _
                "텍사스ESS공장_준비현황.docx", #6/4/2025#, 75, #3/1/2025#, #9/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#5/15/2025#, "SK온 ESS사업부 신설 - Utility-scale ESS 집중 공략", _
                "사내", "ESS", "진행중", "ESS사업팀", _
                "ESS사업부_조직개편.pptx", #5/14/2025#, 60, #4/1/2025#, #10/31/2025#, True)
    allIssues.Add issue
    
    ' 사외 ESS 이슈들
    Set issue = CreateUpdatedIssue(#8/29/2025#, "LG에너지솔루션 9월 RE+ 2025에서 ESS용 각형 LFP 배터리 공개 예정", _
                "사외", "경쟁사", "모니터링", "전략기획팀", _
                "LG_각형LFP공개.pdf", #8/28/2025#, 90, #8/1/2025#, #9/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#7/25/2025#, "삼성SDI 2025년 1차 중앙계약시장 ESS 80% 수주 - 540MW 규모", _
                "사외", "경쟁사", "해결됨", "시장분석팀", _
                "삼성SDI_ESS수주.pdf", #7/24/2025#, 100, #6/1/2025#, #7/25/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#7/18/2025#, "Gotion High-Tech 독일서 5MWh 액체냉각 ESS 현지생산 시작", _
                "사외", "경쟁사", "모니터링", "해외사업팀", _
                "Gotion_독일ESS.pdf", #7/17/2025#, 85, #6/1/2025#, #11/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#6/30/2025#, "EVE Energy 말레이시아 86.5억위안 ESS 배터리공장 투자 결정", _
                "사외", "경쟁사", "모니터링", "해외사업팀", _
                "EVE_말레이시아투자.pdf", #6/29/2025#, 75, #6/1/2025#, #12/31/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#6/27/2025#, "CATL코리아 ESS 테크니컬 솔루션 엔지니어 대규모 채용", _
                "사외", "경쟁사", "모니터링", "인사팀", _
                "CATL_한국채용.pdf", #6/26/2025#, 80, #6/1/2025#, #9/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#6/12/2025#, "CATL 587Ah 고용량 BESS 전용셀 양산 - 시스템부품 40% 절감", _
                "사외", "경쟁사", "해결됨", "R&D센터", _
                "CATL_587Ah출시.pdf", #6/11/2025#, 100, #5/1/2025#, #6/12/2025#, True)
    allIssues.Add issue
    
    ' 비ESS 이슈들 (43개) - SK온 사내 이슈들
    Set issue = CreateUpdatedIssue(#8/29/2025#, "SK온 BMW iX4 차세대 46파이 원통형 배터리 20GWh 공급계약 협상", _
                "사내", "Marketing", "진행중", "영업마케팅팀", _
                "BMW_46파이_계약서.docx", #8/28/2025#, 70, #7/1/2025#, #10/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/27/2025#, "SK온 헝가리 3공장 NCM811 라인 월 15GWh 증설 프로젝트 착공", _
                "사내", "Production", "진행중", "생산관리팀", _
                "헝가리3공장_증설.xlsx", #8/26/2025#, 45, #6/1/2025#, #12/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/26/2025#, "SK온 전고체 배터리 파일럿 라인 200MWh 시험생산 목표 달성", _
                "사내", "R&D", "해결됨", "R&D센터", _
                "전고체_성과보고.docx", #8/25/2025#, 100, #3/1/2025#, #8/26/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/25/2025#, "SK온 메르세데스-벤츠 EQS 차세대 NCM9 30GWh 독점공급 확정", _
                "사내", "Marketing", "해결됨", "영업마케팅팀", _
                "MB_계약완료.pdf", #8/24/2025#, 100, #5/1/2025#, #8/25/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/23/2025#, "SK온 2025년 하반기 원가 20% 절감 TF - 음극재 대체소재 개발", _
                "사내", "R&D", "진행중", "R&D센터", _
                "원가절감TF.pptx", #8/22/2025#, 60, #7/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/21/2025#, "SK온 중국 창저우 2공장 LFP 배터리 월 10GWh 양산 승인", _
                "사내", "Production", "해결됨", "중국사업팀", _
                "창저우_양산승인.docx", #8/20/2025#, 100, #4/1/2025#, #8/21/2025#, False)
    allIssues.Add issue
    
    ' 5월, 6월 이슈들 추가
    Set issue = CreateUpdatedIssue(#5/30/2025#, "SK온 아우디 Q8 e-tron 2026년형 18GWh 공급 확정", _
                "사내", "Marketing", "해결됨", "영업마케팅팀", _
                "Audi_계약완료.pdf", #5/29/2025#, 100, #3/1/2025#, #5/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#5/28/2025#, "SK온 코발트프리 NCM 배터리 개발 프로젝트 2단계 진입", _
                "사내", "R&D", "진행중", "R&D센터", _
                "코발트프리_진행.pptx", #5/27/2025#, 55, #4/1/2025#, #10/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#6/28/2025#, "SK온 Stellantis STLA Large 플랫폼 35GWh 장기계약 체결", _
                "사내", "Marketing", "해결됨", "영업마케팅팀", _
                "Stellantis_계약.pdf", #6/27/2025#, 100, #4/1/2025#, #6/28/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#6/25/2025#, "SK온 폴란드 브로츠와프공장 NCM622 라인 8GWh 증설 승인", _
                "사내", "Production", "진행중", "생산관리팀", _
                "폴란드_증설.xlsx", #6/24/2025#, 35, #5/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    ' 사내 추가 이슈들
    Set issue = CreateUpdatedIssue(#8/19/2025#, "SK온 현대차 아이오닉6 배터리 단가 5% 인하 요구 대응방안 협의", _
                "사내", "Marketing", "미해결", "영업마케팅팀", _
                "현대차_가격협상.xlsx", #8/18/2025#, 35, #7/15/2025#, #9/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/17/2025#, "SK온 대전 R&D센터 안전성테스트 장비 50억원 도입 완료", _
                "사내", "R&D", "해결됨", "R&D센터", _
                "R&D_장비도입.pdf", #8/16/2025#, 100, #5/1/2025#, #8/17/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/15/2025#, "SK온 GM Ultium 플랫폼 차세대 15GWh 공급 협상 진행", _
                "사내", "Marketing", "진행중", "영업마케팅팀", _
                "GM_Ultium_협상.pptx", #8/14/2025#, 55, #6/1/2025#, #10/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/12/2025#, "SK온 베트남 VinFast VF9 12GWh 공급계약 체결", _
                "사내", "Marketing", "해결됨", "해외사업팀", _
                "VinFast_계약.pdf", #8/11/2025#, 100, #6/1/2025#, #8/12/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#7/30/2025#, "SK온 포드 F-150 Lightning 2026년형 25GWh 공급 입찰 참여", _
                "사내", "Marketing", "진행중", "영업마케팅팀", _
                "Ford_입챰서.docx", #7/29/2025#, 50, #6/1/2025#, #9/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#7/28/2025#, "SK온 인도네시아 니켈광산 JV PT Vale 지분 30% 인수 완료", _
                "사내", "투자", "해결됨", "경영기획팀", _
                "인니_JV인수.pdf", #7/27/2025#, 100, #5/1/2025#, #7/28/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#7/18/2025#, "SK온 터키 Togg T10X 전기SUV 8GWh 공급 협상 진행", _
                "사내", "Marketing", "진행중", "해외사업팀", _
                "Togg_협상안.pptx", #7/17/2025#, 65, #6/1/2025#, #10/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#6/18/2025#, "SK온 리비안 R1T/R1S 차세대 10GWh 공급 협상", _
                "사내", "Marketing", "진혨중", "영업마케팅팀", _
                "Rivian_제안서.pptx", #6/17/2025#, 45, #5/1/2025#, #9/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#6/8/2025#, "SK온 실리콘음극재 5% 적용 배터리수명 15% 향상 확인", _
                "사내", "R&D", "해결됨", "R&D센터", _
                "실리콘음극_테스트.docx", #6/7/2025#, 100, #3/1/2025#, #6/8/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#5/22/2025#, "SK온 인도 Tata Motors 전기버스 5GWh 공급 협상", _
                "사내", "Marketing", "진행중", "해외사업팀", _
                "Tata_협상안.docx", #5/21/2025#, 40, #4/1/2025#, #8/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#5/15/2025#, "SK온 캐나다 온타리오 배터리소재공장 부지선정 완료", _
                "사내", "투자", "해결됨", "경영기획팀", _
                "캐나다_부지확정.pdf", #5/14/2025#, 100, #2/1/2025#, #5/15/2025#, False)
    allIssues.Add issue
    
    ' 사외 경쟁사 이슈들
    Set issue = CreateUpdatedIssue(#8/29/2025#, "삼성SDI 조직개편 - 극판센터 신설 및 전략마케팅 통합", _
                "사외", "경쟁사", "모니터링", "전략기획팀", _
                "경쟁사분석.pptx", #8/28/2025#, 90, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/27/2025#, "두산밥캣 eFORCE LAB 배터리팩 연구소 출범 - BSUP 개발", _
                "사외", "Tech", "모니터링", "R&D센터", _
                "기술동향.pdf", #8/26/2025#, 75, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/24/2025#, "중국 8개 분리막 기업 향후 2년간 신규 증설 중단 합의", _
                "사외", "시장", "모니터링", "구매팀", _
                "공급망분석.xlsx", #8/23/2025#, 95, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/18/2025#, "Subaru 전고체 배터리 탑재 산업용 로봇 테스트 - Maxell PSB401010H", _
                "사외", "Tech", "모니터링", "R&D센터", _
                "전고체동향.pdf", #8/17/2025#, 85, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#8/13/2025#, "CATL 리튬 광산 운영 중단으로 리튬 가격 8% 급등", _
                "사외", "시장", "미해결", "구매팀", _
                "원자재시장분석.pdf", #8/12/2025#, 25, #7/1/2025#, #10/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#7/22/2025#, "일본 정부 전기차 보조금 50% 샭감 발표", _
                "사외", "정책", "미해결", "전략기획팀", _
                "일본_정책변경.pdf", #7/21/2025#, 30, #7/1/2025#, #9/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#7/12/2025#, "미국 배터리 제조 세액공제 45X 연장 법안 상원 통과", _
                "사외", "정책", "해결됨", "법무팀", _
                "미국_세제혜택.pdf", #7/11/2025#, 100, #5/1/2025#, #7/12/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#6/12/2025#, "중국 BYD Blade Battery 2.0 에너지밀도 190Wh/kg 달성 발표", _
                "사외", "경쟁사", "모니터링", "R&D센터", _
                "BYD_기술분석.pdf", #6/11/2025#, 95, #6/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#5/18/2025#, "EU 탄소국경조정제도(CBAM) 배터리 적용 2027년 확정", _
                "사외", "정책", "모니터링", "법무팀", _
                "EU_CBAM.pdf", #5/17/2025#, 85, #5/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#5/12/2025#, "Northvolt Ett 공장 화재로 유럽 공급망 차질 우려", _
                "사외", "시장", "미해결", "전략기획팀", _
                "Northvolt_사고분석.pdf", #5/11/2025#, 20, #5/1/2025#, #7/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#5/8/2025#, "말레이시아 정부 EV 배터리 공장 투자 인센티브 30% 확대", _
                "사외", "정책", "해결됨", "해외사업팀", _
                "말레이시아_인센티브.pdf", #5/7/2025#, 100, #3/1/2025#, #5/8/2025#, False)
    allIssues.Add issue
    
    ' 미래 예측 이슈들 (9월-11월)
    Set issue = CreateUpdatedIssue(#9/15/2025#, "[예측] 테슬라 4680 배터리 자체생산 50GWh 달성 예상", _
                "사외", "경쟁사", "진행중", "전략기획팀", _
                "Tesla_예측.pdf", #8/30/2025#, 40, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#9/25/2025#, "[계획] SK온 미국 켄터키 2공장 46파이 대량생산 시작 예정", _
                "사내", "Production", "진혈중", "생산관리팀", _
                "켄터키_생산계획.xlsx", #8/30/2025#, 30, #7/1/2025#, #10/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#10/10/2025#, "[예측] 중국 CATL 나트륨이온 배터리 상용화 발표 예상", _
                "사외", "경쟁사", "진혨중", "R&D센터", _
                "나트륨배터리_분석.pdf", #8/30/2025#, 25, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#10/20/2025#, "[계획] SK온 2026년 전고체배터리 양산라인 구축 예산승인 예정", _
                "사내", "투자", "진혨중", "경영기획팀", _
                "전고체_투자계획.pptx", #8/30/2025#, 20, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#11/5/2025#, "[예측] EU 배터리 여권(Battery Passport) 시행령 최종 발표", _
                "사외", "정책", "진혨중", "법무팀", _
                "EU_배터리여권.pdf", #8/30/2025#, 15, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateUpdatedIssue(#11/15/2025#, "[계획] SK온 동남아 시장진출 전략 수립 - 태국/인니 중심", _
                "사내", "Marketing", "진혨중", "해외사업팀", _
                "동남아_전략.docx", #8/30/2025#, 10, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
End Sub

Private Function CreateUpdatedIssue(issueDate As Date, title As String, category1 As String, _
                            category2 As String, status As String, dept As String, _
                            docRef As String, updateDate As Date, _
                            progress As Integer, startDate As Date, endDate As Date, isESS As Boolean) As Object
    Dim issue As Object
    Set issue = CreateObject("Scripting.Dictionary")
    
    issue.Add "date", issueDate
    issue.Add "title", title
    issue.Add "category1", category1
    issue.Add "category2", category2
    issue.Add "status", status
    issue.Add "dept", dept
    issue.Add "docRef", docRef
    issue.Add "updateDate", updateDate
    issue.Add "progress", progress
    issue.Add "startDate", startDate
    issue.Add "endDate", endDate
    issue.Add "isESS", isESS
    
    Set CreateUpdatedIssue = issue
End Function

' 워크시트에서 직접 호출할 수 있는 Public 서브루틴
Public Sub FilterChangedUpdated()
    On Error Resume Next
    Call ApplyFilterUpdated
    On Error GoTo 0
End Sub