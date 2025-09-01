Attribute VB_Name = "modIssueTimeline"
' Issue Timeline Dashboard Module - Complete Version with All Issues
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
    timelineWs.Columns("C").ColumnWidth = 45  ' 이슈 제목 (더 넓게)
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
    
    ' 필터 영역 라벨
    With timelineWs.Range("B5")
        .Value = "필터:"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' 키워드 검색 영역
    With timelineWs.Range("B6")
        .Value = "키워드:"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' 키워드 입력 필드
    With timelineWs.Range("C6:D6")
        .Merge
        .Value = ""
        .Interior.Color = RGB(255, 255, 224)  ' Light yellow
        .Borders.LineStyle = xlContinuous
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .Name = "KeywordInput"  ' Named range for easy reference
    End With
    
    ' 카테고리 필터 (범위 참조 방식)
    ' 숨겨진 영역에 리스트 생성
    With timelineWs.Range("Z1:Z7")
        .Clear
        .Cells(1, 1).Value = "전체"
        .Cells(2, 1).Value = "전략"
        .Cells(3, 1).Value = "기술"
        .Cells(4, 1).Value = "리스크"
        .Cells(5, 1).Value = "경쟁사"
        .Cells(6, 1).Value = "정책"
        .Cells(7, 1).Value = "ESS"
        .Name = "CategoryList"
    End With
    
    With timelineWs.Range("C5")
        On Error Resume Next
        .Validation.Delete
        On Error GoTo 0
        On Error Resume Next
        .Validation.Add Type:=xlValidateList, _
            Formula1:="=$Z$1:$Z$7"
        On Error GoTo 0
        .Value = "전체"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' 상태 필터 (범위 참조 방식)
    ' 숨겨진 영역에 리스트 생성
    With timelineWs.Range("AA1:AA5")
        .Clear
        .Cells(1, 1).Value = "전체"
        .Cells(2, 1).Value = "미해결"
        .Cells(3, 1).Value = "진행중"
        .Cells(4, 1).Value = "해결됨"
        .Cells(5, 1).Value = "모니터링"
        .Name = "StatusList"
    End With
    
    With timelineWs.Range("D5")
        On Error Resume Next
        .Validation.Delete
        On Error GoTo 0
        On Error Resume Next
        .Validation.Add Type:=xlValidateList, _
            Formula1:="=$AA$1:$AA$5"
        On Error GoTo 0
        .Value = "전체"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' 기간 필터
    With timelineWs.Range("E5")
        .Value = "최근 3개월"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' 키워드 검색 버튼
    Dim searchBtn As Object
    Set searchBtn = timelineWs.Buttons.Add(timelineWs.Range("E6").Left, _
                                          timelineWs.Range("E6").Top, 80, 25)
    With searchBtn
        .Caption = "검색"
        .OnAction = "SearchIssuesByKeyword"
        .Font.Size = 11
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
        .RowHeight = 30
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
    
    ' 샘플 이슈 데이터 - 완전한 버전
    Call AddAllIssues(timelineWs)
    
    ' 범례
    With timelineWs.Range("B50:K51")
        .Interior.Color = RGB(236, 240, 241)
        .Borders.LineStyle = xlContinuous
    End With
    
    timelineWs.Range("B50").Value = "상태:"
    timelineWs.Range("C50").Value = "● 미해결"
    timelineWs.Range("C50").Font.Color = RGB(231, 76, 60)
    timelineWs.Range("D50").Value = "● 진행중"
    timelineWs.Range("D50").Font.Color = RGB(241, 196, 15)
    timelineWs.Range("E50").Value = "● 해결됨"
    timelineWs.Range("E50").Font.Color = RGB(46, 204, 113)
    timelineWs.Range("F50").Value = "● 모니터링"
    timelineWs.Range("F50").Font.Color = RGB(52, 152, 219)
    
    ' 범례 - 타임라인 마커
    timelineWs.Range("B51").Value = "마커:"
    timelineWs.Range("C51").Value = "● 시작/이벤트"
    timelineWs.Range("D51").Value = "▲ 진행/계획"
    timelineWs.Range("E51").Value = "■ 완료"
    
    ' 화면 설정
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    timelineWs.Range("B8").Select
    
    MsgBox "Issue Timeline Dashboard가 생성되었습니다!" & Chr(10) & Chr(10) & _
           "주요 기능:" & Chr(10) & _
           "- 키워드 검색 (ESS 등)" & Chr(10) & _
           "- 이슈별 타임라인 시각화" & Chr(10) & _
           "- 상태별/카테고리별 필터링" & Chr(10) & _
           "- 실시간 데이터 연동" & Chr(10) & _
           "- AI 분석 및 예측", _
           vbInformation, "STRIX Issue Tracker"
End Sub

' 모든 이슈 추가 (ESS 포함 전체)
Private Sub AddAllIssues(ws As Worksheet)
    Dim row As Integer
    row = 9
    
    ' 일반 이슈들 (기존)
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
    
    ' ESS 관련 이슈들
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-18"
    ws.Cells(row, 3).Value = "ESS 사업 진출 전략 수립 - 글로벌 시장 진입"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 7, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 7, "●", "전략 수립")
    
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-20"
    ws.Cells(row, 3).Value = "ESS용 LFP 배터리 기술개발 착수"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 7, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 7, "●", "개발 시작")
    
    ' 이슈 5: CATL 시장점유율
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-10"
    ws.Cells(row, 3).Value = "CATL 점유율 37.9% 달성, 대응 전략 수립 필요"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 7, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 7, "●", "이슈 제기")
    
    ' ESS 화재 리스크
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-25"
    ws.Cells(row, 3).Value = "ESS 화재 리스크 관리 방안 수립"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "경영지원"
    Call DrawTimelineBar(ws, row, 7, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 7, "●", "리스크 분석")
    
    ' 이슈 7: IRA 정책
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-20"
    ws.Cells(row, 3).Value = "IRA AMPC 세액공제 2401억→385억원 급감"
    ws.Cells(row, 4).Value = "정책"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "경영지원"
    Call DrawTimelineBar(ws, row, 7, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 7, "●", "정책 변경")
    
    ' ESS 글로벌 영업
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-01"
    ws.Cells(row, 3).Value = "ESS 글로벌 영업 전략 - 미국/유럽 시장 공략"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "영업마케팅"
    Call DrawTimelineBar(ws, row, 8, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 8, "●", "영업 시작")
    
    ' 이슈 9: 원자재 가격
    row = row + 1
    ws.Cells(row, 2).Value = "2024-01-25"
    ws.Cells(row, 3).Value = "리튬 가격 하락에 따른 배터리 단가 인하 압박"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "구매팀"
    Call DrawTimelineBar(ws, row, 7, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 7, "●", "리스크 식별")
    
    ' 이슈 10: 생산능력 확대
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-01"
    ws.Cells(row, 3).Value = "2024년 글로벌 생산능력 152GWh로 70% 확대"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "생산관리"
    Call DrawTimelineBar(ws, row, 8, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 8, "●", "확대 결정")
    
    ' ESS 미국 IRA 혜택
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-05"
    ws.Cells(row, 3).Value = "미국 ESS 투자세액공제 40% 확대 발표"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "정책대응"
    Call DrawTimelineBar(ws, row, 8, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 8, "●", "정책 발표")
    
    ' 이슈 12: BYD 수직계열화
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-05"
    ws.Cells(row, 3).Value = "BYD 수직계열화 전략으로 가격경쟁력 강화"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 8, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 8, "●", "경쟁사 분석")
    
    ' ESS CATL 사우디 프로젝트
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-08"
    ws.Cells(row, 3).Value = "CATL 사우디 10GWh ESS 프로젝트 수주"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "시장분석"
    Call DrawTimelineBar(ws, row, 8, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 8, "●", "경쟁사 수주")
    
    ' 이슈 14: 테슬라 공급계약
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-10"
    ws.Cells(row, 3).Value = "테슬라 모델3/Y 배터리 공급 물량 9.6% 증가"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "영업팀"
    Call DrawTimelineBar(ws, row, 8, 9, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 8, "●", "계약 체결")
    Call AddTimelineMarker(ws, row, 9, "■", "공급 시작")
    
    ' ESS 시장 전망
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-12"
    ws.Cells(row, 3).Value = "글로벌 ESS 시장 2030년 1,200억 달러 전망"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 8, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 8, "●", "시장 분석")
    
    ' 이슈 16: 전기차 캐즘
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-15"
    ws.Cells(row, 3).Value = "전기차 수요 둔화로 인한 시장 성장률 하락"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "마케팅"
    Call DrawTimelineBar(ws, row, 8, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 8, "●", "시장분석")
    
    ' 이슈 17: 헝가리 공장
    row = row + 1
    ws.Cells(row, 2).Value = "2024-02-20"
    ws.Cells(row, 3).Value = "헝가리 제3공장 증설 완료 및 가동 준비"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "해외사업"
    Call DrawTimelineBar(ws, row, 8, 9, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 9, "■", "증설 완료")
    
    ' ESS 인터배터리 전시
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-01"
    ws.Cells(row, 3).Value = "인터배터리 2024 ESS 솔루션 전시 준비"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "마케팅"
    Call DrawTimelineBar(ws, row, 9, 10, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 10, "■", "전시 완료")
    
    ' 이슈 19: K배터리 점유율
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-01"
    ws.Cells(row, 3).Value = "K배터리 3사 글로벌 점유율 18.4%로 4.7%p 하락"
    ws.Cells(row, 4).Value = "리스크"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 9, "●", "실적 발표")
    
    ' 이슈 20: 46파이 원통형
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-05"
    ws.Cells(row, 3).Value = "46파이 원통형 배터리 파일럿 라인 구축"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 9, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 9, "●", "착공")
    
    ' ESS 안전성 인증
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-08"
    ws.Cells(row, 3).Value = "ESS UL9540A 안전성 인증 획득 추진"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "품질관리"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 9, "●", "인증 추진")
    
    ' 이슈 22: ESG 경영
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-10"
    ws.Cells(row, 3).Value = "폐배터리 재활용 체계 구축 및 ESG 대응"
    ws.Cells(row, 4).Value = "정책"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "ESG팀"
    Call DrawTimelineBar(ws, row, 9, 10, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 9, "●", "체계 구축")
    
    ' 이슈 23: 중국시장 경쟁
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-15"
    ws.Cells(row, 3).Value = "중국 내수시장 CATL-BYD 양강구도 고착화"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "중국사업"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 9, "●", "시장 분석")
    
    ' 이슈 24: 인력 구조조정
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-20"
    ws.Cells(row, 3).Value = "해외사업장 중심 인력감축 및 무급휴직 시행"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "인사팀"
    Call DrawTimelineBar(ws, row, 9, 10, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 10, "■", "조정 완료")
    
    ' ESS 국내 실증사업
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-22"
    ws.Cells(row, 3).Value = "한전 ESS 실증사업 100MWh 참여"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "국내영업"
    Call DrawTimelineBar(ws, row, 9, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 9, "●", "사업 참여")
    
    ' 이슈 26: 각형 배터리
    row = row + 1
    ws.Cells(row, 2).Value = "2024-03-25"
    ws.Cells(row, 3).Value = "각형 배터리 개발 완료, 3대 폼팩터 라인업 구축"
    ws.Cells(row, 4).Value = "기술"
    ws.Cells(row, 5).Value = "해결됨"
    ws.Cells(row, 5).Font.Color = RGB(46, 204, 113)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 9, 10, RGB(46, 204, 113))
    Call AddTimelineMarker(ws, row, 10, "■", "개발 완료")
    
    ' 이슈 27: 4분기 흑자전환
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-01"
    ws.Cells(row, 3).Value = "2024년 4분기 흑자전환 목표 수립 및 추진"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "경영기획"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "목표 수립")
    
    ' ESS 호주시장 진출
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-03"
    ws.Cells(row, 3).Value = "호주 ESS 시장 진출 - 태양광 연계 프로젝트"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "해외영업"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "시장 진출")
    
    ' 이슈 29: 광물자원 확보
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-05"
    ws.Cells(row, 3).Value = "핵심 광물자원 장기 공급계약 체결 추진"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "구매팀"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "협상 시작")
    
    ' 이슈 30: 북미시장 확대
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-10"
    ws.Cells(row, 3).Value = "미국 조지아 공장 2단계 증설 검토"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "해외사업"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "타당성 검토")
    
    ' ESS AI 에너지 관리
    row = row + 1
    ws.Cells(row, 2).Value = "2024-04-12"
    ws.Cells(row, 3).Value = "ESS AI 기반 에너지 관리 시스템 개발"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "개발 착수")
    
    ' 2025년 최신 이슈들
    ' 이슈 32: SK온-SK엔무브 합병
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
    
    ' ESS 트럼프 정책
    row = row + 1
    ws.Cells(row, 2).Value = "2025-01-20"
    ws.Cells(row, 3).Value = "트럼프 2기 ESS IRA 정책 변경 가능성"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "정책대응"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 11, "●", "정책 불확실성")
    
    ' 이슈 34: BYD 글로벌 1위
    row = row + 1
    ws.Cells(row, 2).Value = "2025-01-15"
    ws.Cells(row, 3).Value = "BYD 전기차 판매 테슬라 추월, 점유율 15.7% 달성"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "시장분석"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 11, "●", "시장 역전")
    
    ' ESS 나트륨이온 배터리
    row = row + 1
    ws.Cells(row, 2).Value = "2025-02-01"
    ws.Cells(row, 3).Value = "ESS용 나트륨이온 배터리 개발 착수"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 11, "●", "차세대 기술")
    
    ' 이슈 36: LG엔솔 위기경영
    row = row + 1
    ws.Cells(row, 2).Value = "2025-01-20"
    ws.Cells(row, 3).Value = "LG에너지솔루션 위기경영 선언, 투자계획 전면 재검토"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "모니터링"
    ws.Cells(row, 5).Font.Color = RGB(52, 152, 219)
    ws.Cells(row, 6).Value = "전략기획"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(52, 152, 219))
    Call AddTimelineMarker(ws, row, 11, "●", "경쟁사 동향")
    
    ' ESS 5GWh 목표
    row = row + 1
    ws.Cells(row, 2).Value = "2025-02-05"
    ws.Cells(row, 3).Value = "2025년 ESS 사업 5GWh 수주 목표 수립"
    ws.Cells(row, 4).Value = "ESS"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "영업마케팅"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 11, "●", "목표 수립")
    
    ' 이슈 38: 5조원 자본확충
    row = row + 1
    ws.Cells(row, 2).Value = "2025-07-30"
    ws.Cells(row, 3).Value = "SK이노-SK온 5조원 규모 자본확충 진행"
    ws.Cells(row, 4).Value = "전략"
    ws.Cells(row, 5).Value = "진행중"
    ws.Cells(row, 5).Font.Color = RGB(241, 196, 15)
    ws.Cells(row, 6).Value = "재무팀"
    Call DrawTimelineBar(ws, row, 10, 11, RGB(241, 196, 15))
    Call AddTimelineMarker(ws, row, 10, "●", "유상증자 착수")
    
    ' 이슈 39: BYD 5분 충전
    row = row + 1
    ws.Cells(row, 2).Value = "2025-01-12"
    ws.Cells(row, 3).Value = "BYD 5분 충전 400km 주행 기술 공개, 게임체인저 등장"
    ws.Cells(row, 4).Value = "경쟁사"
    ws.Cells(row, 5).Value = "미해결"
    ws.Cells(row, 5).Font.Color = RGB(231, 76, 60)
    ws.Cells(row, 6).Value = "R&D"
    Call DrawTimelineBar(ws, row, 11, 11, RGB(231, 76, 60))
    Call AddTimelineMarker(ws, row, 11, "●", "기술 격차")
    
    ' 이슈 40: Q1 실적 우려
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

' 타임라인 바 그리기 - 핵심 함수
Private Sub DrawTimelineBar(ws As Worksheet, row As Integer, startCol As Integer, endCol As Integer, barColor As Long)
    Dim cell As Range
    For Each cell In ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol))
        cell.Interior.Color = barColor
        cell.Interior.Pattern = xlSolid
    Next cell
End Sub

' 타임라인 마커 추가 - 핵심 함수
Private Sub AddTimelineMarker(ws As Worksheet, row As Integer, col As Integer, marker As String, tooltip As String)
    With ws.Cells(row, col)
        .Value = marker
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        On Error Resume Next
        .AddComment tooltip
        .Comment.Visible = False
        On Error GoTo 0
    End With
End Sub

' 키워드 검색 함수
Sub SearchIssuesByKeyword()
    Dim ws As Worksheet
    Dim keyword As String
    Dim i As Integer
    Dim issueTitle As String
    Dim found As Boolean
    Dim foundCount As Integer
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    keyword = Trim(ws.Range("C6").Value)
    
    ' 모든 행 보이기 및 색상 초기화
    ws.Rows("9:50").Hidden = False
    foundCount = 0
    
    For i = 9 To 50
        If ws.Cells(i, 3).Value <> "" Then
            ' 색상 초기화
            ws.Cells(i, 3).Font.Color = RGB(0, 0, 0)
            ws.Cells(i, 3).Font.Bold = False
            
            ' 키워드가 입력된 경우 검색
            If keyword <> "" Then
                issueTitle = ws.Cells(i, 3).Value
                If InStr(1, issueTitle, keyword, vbTextCompare) > 0 Then
                    ' 키워드 매칭 - 빨간색으로 강조
                    ws.Cells(i, 3).Font.Color = RGB(255, 0, 0)
                    ws.Cells(i, 3).Font.Bold = True
                    foundCount = foundCount + 1
                Else
                    ' 매칭되지 않은 행 숨기기
                    ws.Rows(i).Hidden = True
                End If
            End If
        End If
    Next i
    
    ' 결과 메시지
    If keyword <> "" Then
        If foundCount > 0 Then
            MsgBox "'" & keyword & "' 키워드가 포함된 " & foundCount & "개 이슈를 찾았습니다.", vbInformation
        Else
            MsgBox "'" & keyword & "' 키워드가 포함된 이슈가 없습니다.", vbExclamation
            ' 키워드가 없으면 모든 행 다시 표시
            ws.Rows("9:50").Hidden = False
        End If
    Else
        MsgBox "검색할 키워드를 입력해주세요.", vbExclamation
    End If
End Sub

' 필터링 함수들 (Data Validation 사용)
Sub FilterIssuesByCategory()
    Dim ws As Worksheet
    Dim selectedCategory As String
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    selectedCategory = ws.Range("C5").Value
    
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
    
    Application.StatusBar = selectedCategory & " 카테고리 필터 적용됨"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
End Sub

Sub FilterIssuesByStatus()
    Dim ws As Worksheet
    Dim selectedStatus As String
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    selectedStatus = ws.Range("D5").Value
    
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
    
    Application.StatusBar = selectedStatus & " 상태 필터 적용됨"
    Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
End Sub

' 상태바 초기화
Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

' 모든 필터 초기화
Sub ResetAllFilters()
    Dim ws As Worksheet
    Dim i As Integer
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    ' 모든 행 보이기
    ws.Rows("9:50").Hidden = False
    
    ' Data Validation 필드를 "전체"로 설정
    ws.Range("C5").Value = "전체"
    ws.Range("D5").Value = "전체"
    
    ' 키워드 필드 초기화
    ws.Range("C6").Value = ""
    
    ' 이슈 제목 색상 초기화
    For i = 9 To 50
        If ws.Cells(i, 3).Value <> "" Then
            ws.Cells(i, 3).Font.Color = RGB(0, 0, 0)
            ws.Cells(i, 3).Font.Bold = False
        End If
    Next i
    
    MsgBox "모든 필터가 초기화되었습니다.", vbInformation
End Sub

' 새로고침
Sub RefreshIssueTimeline()
    On Error GoTo ErrorHandler
    
    ' 타임라인 재생성
    Call CreateIssueTimelineDashboard
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
    url = "http://localhost:5001/api/issues/predict"
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json; charset=utf-8"
    http.send "{}"
    
    If http.Status = 200 Then
        MsgBox "AI 분석이 완료되었습니다!" & Chr(10) & Chr(10) & _
               "분석 결과가 Smart Alerts 탭으로 전송되었습니다." & Chr(10) & _
               "Smart Alerts 탭에서 상세한 예측 및 경고 사항을 확인하세요." & Chr(10) & Chr(10) & _
               "✓ 리스크 예측" & Chr(10) & _
               "✓ 우선순위 분석" & Chr(10) & _
               "✓ 의사결정 추천", _
               vbInformation, "AI 분석 완료 - Smart Alerts"
        
        ' Smart Alerts 탭으로 이동
        On Error Resume Next
        ThisWorkbook.Sheets("Smart Alerts").Activate
        On Error GoTo 0
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