Attribute VB_Name = "modIssueTimeline"
' Issue Timeline Dashboard Module
Option Explicit

' 이슈 타임라인 대시보드 생성
Sub CreateIssueTimelineDashboard()
    Dim ws As Worksheet
    Dim timelineWs As Worksheet
    
    ' 기존 시트 삭제
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Issue Timeline").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' 새 시트 생성
    Set timelineWs = ThisWorkbook.Sheets.Add
    timelineWs.Name = "Issue Timeline"
    timelineWs.Activate
    
    ' 전체 배경색
    timelineWs.Cells.Interior.Color = RGB(245, 245, 245)
    
    ' 열 너비 설정
    timelineWs.Columns("A").ColumnWidth = 2
    timelineWs.Columns("B").ColumnWidth = 15  ' 날짜
    timelineWs.Columns("C").ColumnWidth = 40  ' 이슈 제목
    timelineWs.Columns("D").ColumnWidth = 12  ' 카테고리
    timelineWs.Columns("E").ColumnWidth = 12  ' 상태
    timelineWs.Columns("F").ColumnWidth = 15  ' 부서
    timelineWs.Columns("G:K").ColumnWidth = 20 ' 타임라인
    timelineWs.Columns("L").ColumnWidth = 2
    
    ' 헤더 영역
    With timelineWs.Range("B2:K2")
        .Merge
        .Value = "STRIX Issue Timeline & Decision Tracker"
        .Font.Name = "맑은 고딕"
        .Font.Size = 24
        .Font.Bold = True
        .Interior.Color = RGB(41, 128, 185)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 50
    End With
    
    ' 부제목
    With timelineWs.Range("B3:K3")
        .Merge
        .Value = "사내 이슈 진행 현황 및 의사결정 추적 시스템"
        .Font.Size = 14
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' 필터 영역
    timelineWs.Range("B5").Value = "필터:"
    timelineWs.Range("B5").Font.Bold = True
    timelineWs.Range("B5").Font.Size = 12
    
    ' 카테고리 필터
    Dim categoryBtn As Object
    Set categoryBtn = timelineWs.DropDowns.Add(timelineWs.Range("C5").Left, _
                                               timelineWs.Range("C5").Top, 100, 20)
    With categoryBtn
        .AddItem "전체"
        .AddItem "전략"
        .AddItem "기술"
        .AddItem "리스크"
        .AddItem "경쟁사"
        .AddItem "정책"
        .Value = 1
        .OnAction = "FilterIssuesByCategory"
    End With
    
    ' 상태 필터
    Dim statusBtn As Object
    Set statusBtn = timelineWs.DropDowns.Add(timelineWs.Range("D5").Left, _
                                            timelineWs.Range("D5").Top, 100, 20)
    With statusBtn
        .AddItem "전체"
        .AddItem "미해결"
        .AddItem "진행중"
        .AddItem "해결됨"
        .AddItem "모니터링"
        .Value = 1
        .OnAction = "FilterIssuesByStatus"
    End With
    
    ' 기간 필터
    With timelineWs.Range("E5")
        .Value = "최근 3개월"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    ' 새로고침 버튼
    Dim refreshBtn As Object
    Set refreshBtn = timelineWs.Buttons.Add(timelineWs.Range("F5").Left, _
                                           timelineWs.Range("F5").Top, 80, 25)
    With refreshBtn
        .Caption = "새로고침"
        .OnAction = "RefreshIssueTimeline"
        .Font.Size = 11
    End With
    
    ' AI 분석 버튼
    Dim aiBtn As Object
    Set aiBtn = timelineWs.Buttons.Add(timelineWs.Range("G5").Left, _
                                       timelineWs.Range("G5").Top, 80, 25)
    With aiBtn
        .Caption = "AI 분석"
        .OnAction = "RunIssueAIAnalysis"
        .Font.Size = 11
    End With
    
    ' 필터 초기화 버튼
    Dim resetBtn As Object
    Set resetBtn = timelineWs.Buttons.Add(timelineWs.Range("H5").Left, _
                                          timelineWs.Range("H5").Top, 80, 25)
    With resetBtn
        .Caption = "필터 초기화"
        .OnAction = "ResetAllFilters"
        .Font.Size = 11
    End With
    
    ' 타임라인 영역 헤더
    With timelineWs.Range("B8:K8")
        .Interior.Color = RGB(52, 73, 94)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    timelineWs.Range("B8").Value = "최초 언급"
    timelineWs.Range("C8").Value = "이슈 제목"
    timelineWs.Range("D8").Value = "카테고리"
    timelineWs.Range("E8").Value = "상태"
    timelineWs.Range("F8").Value = "담당부서"
    
    ' 타임라인 월 헤더 (동적 생성)
    Dim currentMonth As Date
    Dim col As Integer
    currentMonth = DateSerial(Year(Date), Month(Date) - 2, 1) ' 3개월 전부터
    
    For col = 7 To 11
        timelineWs.Cells(8, col).Value = Format(currentMonth, "yyyy-MM")
        currentMonth = DateAdd("m", 1, currentMonth)
    Next col
    
    ' 샘플 이슈 데이터
    Call AddSampleIssues(timelineWs)
    
    ' 범례
    With timelineWs.Range("B45:K46")
        .Interior.Color = RGB(236, 240, 241)
        .Borders.LineStyle = xlContinuous
    End With
    
    timelineWs.Range("B45").Value = "상태:"
    timelineWs.Range("C45").Value = "● 미해결"
    timelineWs.Range("C45").Font.Color = RGB(231, 76, 60)
    timelineWs.Range("D45").Value = "● 진행중"
    timelineWs.Range("D45").Font.Color = RGB(241, 196, 15)
    timelineWs.Range("E45").Value = "● 해결됨"
    timelineWs.Range("E45").Font.Color = RGB(46, 204, 113)
    timelineWs.Range("F45").Value = "● 모니터링"
    timelineWs.Range("F45").Font.Color = RGB(52, 152, 219)
    
    ' 범례 - 타임라인 마커
    timelineWs.Range("B46").Value = "마커:"
    timelineWs.Range("C46").Value = "● 시작/이벤트"
    timelineWs.Range("D46").Value = "▲ 진행/계획"
    timelineWs.Range("E46").Value = "☑ 완료"
    
    ' 화면 설정
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    timelineWs.Range("B8").Select
    
    MsgBox "Issue Timeline Dashboard가 생성되었습니다!" & Chr(10) & Chr(10) & _
           "주요 기능:" & Chr(10) & _
           "- 이슈별 타임라인 시각화" & Chr(10) & _
           "- 상태별/카테고리별 필터링" & Chr(10) & _
           "- 실시간 데이터 연동" & Chr(10) & _
           "- AI 분석 및 예측", _
           vbInformation, "STRIX Issue Tracker"
End Sub

' 샘플 이슈 추가
Private Sub AddSampleIssues(ws As Worksheet)
    Dim row As Integer
    row = 9
    
    ' 이슈 1: SK온 적자 문제
    ws.Cells(row, 2).Value = "2024-01-05"
    ws.Cells(row, 3).Value = "SK온 연속 적자 및 재무구조 개선 필요"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "재무팀"
    Call DrawTimelineBar(ws, row, 7, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 7, "●", "최초 보고")
    Call AddTimelineMarker(ws, row, 9, "▲", "구조조정 착수")
    
    ' 이슈 2: 전고체 배터리 개발
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-15"
    ws.Cells(row, 3).Value = "전고체 배터리 양산 기술 개발 및 파일럿 라인 구축"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 7, 9, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 7, "●", "최초 언급")
    Call AddTimelineMarker(ws, row, 8, "▲", "파일럿 계획")
    
    ' 이슈 3: CATL 시장점유율 확대
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-10"
    ws.Cells(row, 3).Value = "CATL 점유율 37.9% 달성, 대응 전략 수립 필요"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 7, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 7, "●", "이슈 제기")
    
    ' 이슈 4: 미국 IRA 세액공제 축소
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-20"
    ws.Cells(row, 3).Value = "IRA AMPC 세액공제 2401억→385억원 급감"
    ws.Cells(row, 4).Value = "정책"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "경영지원"
    Call DrawTimelineBar(ws, row, 7, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 7, "●", "정책 변경")
    
    ' 이슈 5: 원자재 가격 변동
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-25"
    ws.Cells(row, 3).Value = "리튬 가격 하락에 따른 배터리 단가 인하 압박"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "구매팀"
    Call DrawTimelineBar(ws, row, 7, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 7, "●", "리스크 식별")
    
    ' 이슈 6: 생산능력 확대
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-01"
    ws.Cells(row, 3).Value = "2024년 글로벌 생산능력 152GWh로 70% 확대"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "생산관리"
    Call DrawTimelineBar(ws, row, 8, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 8, "●", "확대 결정")
    
    ' 이슈 7: BYD 배터리 자체생산
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-05"
    ws.Cells(row, 3).Value = "BYD 수직계열화 전략으로 가격경쟁력 강화"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 8, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 8, "●", "경쟁사 분석")
    
    ' 이슈 8: 테슬라 공급계약
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-10"
    ws.Cells(row, 3).Value = "테슬라 모델3/Y 배터리 공급 물량 9.6% 증가"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "영업팀"
    Call DrawTimelineBar(ws, row, 8, 9, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 8, "●", "계약 체결")
    Call AddTimelineMarker(ws, row, 9, "☑", "공급 시작")
    
    ' 이슈 9: 전기차 캐즘
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-15"
    ws.Cells(row, 3).Value = "전기차 수요 둔화로 인한 시장 성장률 하락"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "마케팅"
    Call DrawTimelineBar(ws, row, 8, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 8, "●", "시장분석")
    
    ' 이슈 10: 헝가리 공장 증설
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-20"
    ws.Cells(row, 3).Value = "헝가리 제3공장 증설 완료 및 가동 준비"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "해외사업"
    Call DrawTimelineBar(ws, row, 8, 9, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 9, "☑", "증설 완료")
    
    ' 이슈 11: K배터리 점유율 하락
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-01"
    ws.Cells(row, 3).Value = "K배터리 3사 글로벌 점유율 18.4%로 4.7%p 하락"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 9, "●", "실적 발표")
    
    ' 이슈 12: 46파이 원통형 배터리
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-05"
    ws.Cells(row, 3).Value = "46파이 원통형 배터리 파일럿 라인 구축"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 9, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 9, "●", "착공")
    
    ' 이슈 13: ESG 경영 강화
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-10"
    ws.Cells(row, 3).Value = "폐배터리 재활용 체계 구축 및 ESG 대응"
    ws.Cells(row, 4).Value = "정책"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "ESG팀"
    Call DrawTimelineBar(ws, row, 9, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 9, "●", "체계 구축")
    
    ' 이슈 14: 중국시장 경쟁 심화
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-15"
    ws.Cells(row, 3).Value = "중국 내수시장 CATL-BYD 양강구도 고착화"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "중국사업"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 9, "●", "시장 분석")
    
    ' 이슈 15: 인력 구조조정
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-20"
    ws.Cells(row, 3).Value = "해외사업장 중심 인력감축 및 무급휴직 시행"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "인사팀"
    Call DrawTimelineBar(ws, row, 9, 10, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 10, "☑", "조정 완료")
    
    ' 이슈 16: 각형 배터리 개발
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-25"
    ws.Cells(row, 3).Value = "각형 배터리 개발 완료, 3대 폼팩터 라인업 구축"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 9, 10, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 10, "☑", "개발 완료")
    
    ' 이슈 17: 4분기 흑자전환 목표
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-01"
    ws.Cells(row, 3).Value = "2024년 4분기 흑자전환 목표 수립 및 추진"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "경영기획"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "목표 수립")
    
    ' 이슈 18: 광물자원 확보
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-05"
    ws.Cells(row, 3).Value = "핵심 광물자원 장기 공급계약 체결 추진"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "구매팀"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "협상 시작")
    
    ' 이슈 19: 북미시장 확대
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-10"
    ws.Cells(row, 3).Value = "미국 조지아 공장 2단계 증설 검토"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "해외사업"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "타당성 검토")
    
    ' 이슈 20: 인터배터리 2025 참가
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-15"
    ws.Cells(row, 3).Value = "인터배터리 2025 차세대 기술 전시 준비"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "마케팅"
    Call DrawTimelineBar(ws, row, 10, 10, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 10, "☑", "전시 완료")
    
    ' 2025년 최신 이슈 추가
    
    ' 이슈 21: SK온-SK엔무브 합병
    row = row + 1
    ws.Cells(row, 2).Value = "2025-07-30"
    ws.Cells(row, 3).Value = "SK온-SK엔무브 합병 결정, 11월 1일 통합법인 출범"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "경영기획"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "합병 결의")
    Call AddTimelineMarker(ws, row, 11, "▲", "11월 출범예정")
    
    ' 이슈 22: 트럼프 IRA 정책 변경 위기
    row = row + 1
    ws.Cells(row, 2).Value = "2025-01-20"
    ws.Cells(row, 3).Value = "트럼프 2기 IRA 폐지 가능성, AMPC 축소 우려"
    ws.Cells(row, 4).Value = "정책"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "정책대응"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 11, "●", "정책 불확실성")
    
    ' 이슈 23: BYD 글로벌 1위 도약
    row = row + 1
    ws.Cells(row, 2).Value = "2025-01-15"
    ws.Cells(row, 3).Value = "BYD 전기차 판매 테슬라 추월, 점유율 15.7% 달성"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "시장분석"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 11, "●", "시장 역전")
    
    ' 이슈 24: LG엔솔 위기경영 선언
    row = row + 1
    ws.Cells(row, 2).Value = "2025-01-20"
    ws.Cells(row, 3).Value = "LG에너지솔루션 위기경영 선언, 투자계획 전면 재검토"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 11, "●", "경쟁사 동향")
    
    ' 이슈 25: 5조원 자본확충 추진
    row = row + 1
    ws.Cells(row, 2).Value = "2025-07-30"
    ws.Cells(row, 3).Value = "SK이노-SK온 5조원 규모 자본확충 진행"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "재무팀"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "유상증자 착수")
    
    ' 이슈 26: 니켈 중심 배터리 전환
    row = row + 1
    ws.Cells(row, 2).Value = "2025-01-10"
    ws.Cells(row, 3).Value = "코발트 프리 배터리 가속화, 니켈 비중 확대"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 11, "●", "기술 전환")
    
    ' 이슈 27: 2030년 EBITDA 20조원 목표
    row = row + 1
    ws.Cells(row, 2).Value = "2025-07-30"
    ws.Cells(row, 3).Value = "SK이노베이션 2030년 EBITDA 20조원 달성 목표 수립"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "경영기획"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "중장기 목표")
    
    ' 이슈 28: LFP 배터리 수요 증가
    row = row + 1
    ws.Cells(row, 2).Value = "2025-01-25"
    ws.Cells(row, 3).Value = "테슬라 LFP 배터리 비중 65% 확대, 시장 판도 변화"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "기술전략"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 11, "●", "시장 트렌드")
    
    ' 이슈 29: BYD 5분 충전기술 공개
    row = row + 1
    ws.Cells(row, 2).Value = "2025-01-12"
    ws.Cells(row, 3).Value = "BYD 5분 충전 400km 주행 기술 공개, 게임체인저 등장"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 11, "●", "기술 격차")
    
    ' 이슈 30: Q1 실적 AMPC 의존도 심화
    row = row + 1
    ws.Cells(row, 2).Value = "2025-04-07"
    ws.Cells(row, 3).Value = "LG엔솔 Q1 AMPC 제외시 830억 영업손실, 의존도 심화"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "재무팀"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 11, "●", "실적 우려")
    
    ' 행 서식 적용
    Dim i As Integer
    For i = 9 To row
        With ws.Range(ws.Cells(i, 2), ws.Cells(i, 11))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200)
            If i Mod 2 = 0 Then
                .Interior.Color = RGB(248, 248, 248)
            End If
        End With
        ws.Rows(i).RowHeight = 25
    Next i
End Sub

' 타임라인 바 그리기
Private Sub DrawTimelineBar(ws As Worksheet, row As Integer, startCol As Integer, endCol As Integer, barColor As Long)
    Dim cell As Range
    For Each cell In ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol))
        cell.Interior.Color = barColor
        cell.Interior.Pattern = xlSolid
    Next cell
End Sub

' 타임라인 마커 추가
Private Sub AddTimelineMarker(ws As Worksheet, row As Integer, col As Integer, marker As String, tooltip As String)
    With ws.Cells(row, col)
        .Value = marker
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .AddComment tooltip
        .Comment.Visible = False
    End With
End Sub

' 필터링 함수들
Sub FilterIssuesByCategory()
    Dim ws As Worksheet
    Dim categoryDropDown As DropDown
    Dim selectedCategory As String
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    Set categoryDropDown = ws.DropDowns(1)
    
    selectedCategory = categoryDropDown.List(categoryDropDown.Value)
    
    ' 모든 행 보이기
    ws.Rows("9:50").Hidden = False
    
    ' 선택된 카테고리가 "전체"가 아니면 필터링
    If selectedCategory <> "전체" Then
        For i = 9 To 50
            If ws.Cells(i, 4).Value <> "" Then
                If ws.Cells(i, 4).Value <> selectedCategory Then
                    ws.Rows(i).Hidden = True
                End If
            End If
        Next i
    End If
    
    MsgBox selectedCategory & " 카테고리 필터가 적용되었습니다.", vbInformation
End Sub

Sub FilterIssuesByStatus()
    Dim ws As Worksheet
    Dim statusDropDown As DropDown
    Dim selectedStatus As String
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    Set statusDropDown = ws.DropDowns(2)
    
    selectedStatus = statusDropDown.List(statusDropDown.Value)
    
    ' 모든 행 보이기
    ws.Rows("9:50").Hidden = False
    
    ' 선택된 상태가 "전체"가 아니면 필터링
    If selectedStatus <> "전체" Then
        For i = 9 To 50
            If ws.Cells(i, 5).Value <> "" Then
                If ws.Cells(i, 5).Value <> selectedStatus Then
                    ws.Rows(i).Hidden = True
                End If
            End If
        Next i
    End If
    
    MsgBox selectedStatus & " 상태 필터가 적용되었습니다.", vbInformation
End Sub

' 모든 필터 초기화
Sub ResetAllFilters()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    ' 모든 행 보이기
    ws.Rows("9:50").Hidden = False
    
    ' 드롭다운을 "전체"로 설정
    ws.DropDowns(1).Value = 1
    ws.DropDowns(2).Value = 1
    
    MsgBox "모든 필터가 초기화되었습니다.", vbInformation
End Sub

' 새로고침
Sub RefreshIssueTimeline()
    On Error GoTo ErrorHandler
    
    ' UpdateIssueTimeline 함수 호출 (modIssueAPI에 정의됨)
    Call UpdateIssueTimeline
    Exit Sub
    
ErrorHandler:
    MsgBox "타임라인 새로고침 중 오류: " & Err.Description, vbCritical
End Sub

' AI 분석 실행
Sub RunIssueAIAnalysis()
    Dim http As Object
    Dim url As String
    Dim responseText As String
    
    On Error GoTo ErrorHandler
    
    ' 상태 표시
    Application.StatusBar = "AI가 미해결 이슈를 분석중입니다..."
    
    ' API 호출
    url = "http://localhost:5000/api/issues/predict"
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send "{}"
    
    If http.Status = 200 Then
        MsgBox "AI 분석이 완료되었습니다!" & Chr(10) & _
               "각 이슈의 예측 정보가 업데이트되었습니다." & Chr(10) & Chr(10) & _
               "타임라인에서 이슈를 클릭하여 AI 예측을 확인하세요.", _
               vbInformation, "AI 분석 완료"
        
        ' 타임라인 새로고침
        Call RefreshIssueTimeline
    Else
        MsgBox "AI 분석 중 오류가 발생했습니다." & Chr(10) & _
               "API 서버 상태를 확인해주세요.", vbExclamation
    End If
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "AI 분석 실행 중 오류: " & Err.Description, vbCritical
End Sub