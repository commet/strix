Attribute VB_Name = "modIssueTimelineFinalV8"
Option Explicit

' 최종 버전 V8 - 필터 자동 작동 및 열 순서 수정
Private allIssues As Collection
Private filteredIssues As Collection

Sub CreateFinalDashboardV8()
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
        .Font.Name = "맑은 고딕"
        .Font.Size = 24
        .Font.Bold = True
        .Interior.Color = RGB(39, 55, 39)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
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
    
    ' 검색 영역 (더 크게)
    ws.Range("B5").Value = "검색:"
    ws.Range("B5").Font.Bold = True
    ws.Range("B5").Font.Size = 14
    
    With ws.Range("C5:G5")
        .Merge
        .Name = "SearchBox"
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlMedium
        .Font.Size = 14
        .RowHeight = 30
    End With
    
    ' 검색 버튼
    Set btn = ws.Buttons.Add(ws.Range("H5").Left, ws.Range("H5").Top, _
                             ws.Range("H5").Width, ws.Range("H5").Height)
    With btn
        .Caption = "검색"
        .OnAction = "SearchFinalV8"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' 전체보기 버튼
    Set btn = ws.Buttons.Add(ws.Range("I5").Left, ws.Range("I5").Top, _
                             ws.Range("I5").Width, ws.Range("I5").Height)
    With btn
        .Caption = "전체보기"
        .OnAction = "ShowAllFinalV8"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' 새로고침 버튼
    Set btn = ws.Buttons.Add(ws.Range("J5").Left, ws.Range("J5").Top, _
                             ws.Range("J5").Width, ws.Range("J5").Height)
    With btn
        .Caption = "새로고침"
        .OnAction = "RefreshFinalV8"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' 결과 표시
    ws.Range("K5").Font.Color = RGB(0, 0, 255)
    ws.Range("K5").Font.Size = 14
    ws.Range("K5").Font.Bold = True
    
    ' 필터 라벨 (7행) - 순서 변경: 상태와 담당부서 위치 교체
    ws.Range("D7").Value = "분류1"
    ws.Range("E7").Value = "세부구분"
    ws.Range("F7").Value = "상태"      ' 상태를 먼저
    ws.Range("G7").Value = "담당부서"  ' 담당부서를 나중에
    
    With ws.Range("D7:G7")
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 필터 드롭다운 (8행)
    ' 분류1 필터
    With ws.Range("D8")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(100, 100, 100)
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        On Error Resume Next
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            Formula1:="전체,사내,사외"
        On Error GoTo 0
        .Value = "전체"
        .RowHeight = 25
    End With
    
    ' 세부구분 필터
    With ws.Range("E8")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(100, 100, 100)
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        On Error Resume Next
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            Formula1:="전체,정책,경쟁사,Tech,Marketing,Production,R&D,Staff,ESS,투자,특허,시장"
        On Error GoTo 0
        .Value = "전체"
    End With
    
    ' 상태 필터 (F8로 변경)
    With ws.Range("F8")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(100, 100, 100)
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        On Error Resume Next
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            Formula1:="전체,해결됨,모니터링,진행중,미해결"
        On Error GoTo 0
        .Value = "전체"
    End With
    
    ' 담당부서 필터 (G8로 변경)
    With ws.Range("G8")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(100, 100, 100)
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        On Error Resume Next
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
            Formula1:="전체,전략기획팀,생산관리팀,품질관리팀,영업마케팅팀,R&D센터,경영지원팀,구매팀,인사팀,시장분석팀,경영기획팀,법무팀,안전환경팀,해외사업팀,중국사업팀,ESS사업팀"
        On Error GoTo 0
        .Value = "전체"
    End With
    
    ' 타임라인 월 헤더 (10행)
    Dim monthNames As Variant
    Dim i As Integer
    monthNames = Array("2025-05", "2025-06", "2025-07", "2025-08", "2025-09", "2025-10", "2025-11")
    
    For i = 0 To UBound(monthNames)
        ws.Cells(10, 9 + i).Value = monthNames(i)
        ws.Cells(10, 9 + i).HorizontalAlignment = xlCenter
        ws.Cells(10, 9 + i).Font.Bold = True
        ws.Cells(10, 9 + i).Font.Size = 12
        ws.Cells(10, 9 + i).Interior.Color = RGB(220, 220, 220)
        ws.Cells(10, 9 + i).Borders.LineStyle = xlContinuous
    Next i
    
    ' 테이블 헤더 (10행) - 순서 변경: 상태와 담당부서 위치 교체
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
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 데이터 로드
    Call LoadAllIssuesFinalV8
    
    ' 전체 이슈 표시
    Call ApplyFiltersV8(ws)
    
    ' 컬럼 너비 조정
    ws.Columns("A").ColumnWidth = 5   ' No
    ws.Columns("B").ColumnWidth = 12  ' 날짜
    ws.Columns("C").ColumnWidth = 75  ' 제목
    ws.Columns("D:E").ColumnWidth = 10  ' 분류
    ws.Columns("F").ColumnWidth = 10  ' 상태
    ws.Columns("G").ColumnWidth = 15  ' 담당부서
    ws.Columns("H").ColumnWidth = 8   ' 진행률
    ws.Columns("I:O").ColumnWidth = 10  ' 타임라인
    ws.Columns("P").ColumnWidth = 20  ' 문서 참조
    ws.Columns("Q").ColumnWidth = 12  ' 업데이트
    
    ' 행 높이 조정
    ws.Rows("11:70").RowHeight = 20
    
    ws.Activate
    ws.Range("C5").Select
    
    MsgBox "Issue Timeline Dashboard가 생성되었습니다!" & vbCrLf & _
           "총 54개 이슈가 로드되었습니다." & vbCrLf & vbCrLf & _
           "※ 필터는 드롭다운 선택 시 자동 적용됩니다.", vbInformation
End Sub

Private Sub LoadAllIssuesFinalV8()
    Set allIssues = New Collection
    Dim issue As Object
    
    ' 2025년 8월 이슈들
    Set issue = CreateIssueFinalV8(#8/29/2025#, "LG에너지솔루션, 9월 각형 배터리 공개 - RE+ 2025 전시회", _
                "사외", "경쟁사", "모니터링", "전략기획팀", _
                "분석보고서.pdf", #8/28/2025#, 85, #8/1/2025#, #11/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/28/2025#, "BMW iX4 2026년형 46시리즈 원통형 배터리 20GWh 공급계약 협상", _
                "사내", "Marketing", "진행중", "영업마케팅팀", _
                "BMW_계약서_초안.docx", #8/27/2025#, 70, #7/1/2025#, #10/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/29/2025#, "삼성SDI 조직개편 - 극판센터 신설 및 전략마케팅 통합", _
                "사외", "경쟁사", "모니터링", "전략기획팀", _
                "경쟁사분석.pptx", #8/28/2025#, 90, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/27/2025#, "헝가리 3공장 NCM811 라인 월 15GWh 증설 프로젝트 착공", _
                "사내", "Production", "진행중", "생산관리팀", _
                "헝가리3공장_증설.xlsx", #8/26/2025#, 45, #6/1/2025#, #12/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/13/2025#, "CATL 리튬 광산 운영 중단으로 리튬 가격 8% 급등", _
                "사외", "시장", "미해결", "구매팀", _
                "원자재시장분석.pdf", #8/12/2025#, 25, #7/1/2025#, #10/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/26/2025#, "전고체 배터리 파일럿 라인 월 200MWh 시험생산 목표 달성", _
                "사내", "R&D", "해결됨", "R&D센터", _
                "전고체_성과보고.docx", #8/25/2025#, 100, #3/1/2025#, #8/26/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/27/2025#, "두산밥캣 eFORCE LAB 배터리팩 연구소 출범 - BSUP 개발", _
                "사외", "Tech", "모니터링", "R&D센터", _
                "기술동향.pdf", #8/26/2025#, 75, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/25/2025#, "메르세데스 벤츠 EQS 2027년형 NCM9 배터리 30GWh 독점공급 확정", _
                "사내", "Marketing", "해결됨", "영업마케팅팀", _
                "MB_계약완료.pdf", #8/24/2025#, 100, #5/1/2025#, #8/25/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/24/2025#, "중국 8개 분리막 기업 향후 2년간 신규 증설 중단 합의", _
                "사외", "시장", "모니터링", "구매팀", _
                "공급망분석.xlsx", #8/23/2025#, 95, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/23/2025#, "2025년 하반기 원가 20% 절감 TF - 음극재 대체소재 개발", _
                "사내", "R&D", "진행중", "R&D센터", _
                "원가절감계획.pptx", #8/22/2025#, 60, #7/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/22/2025#, "SK온-에코프로 폐배터리 재활용 협력 및 블랙파우더 공급계약", _
                "사외", "ESS", "모니터링", "ESS사업팀", _
                "ESG동향.pdf", #8/21/2025#, 85, #8/1/2025#, #11/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/21/2025#, "중국 창저우 2공장 LFP 배터리 월 10GWh 양산 승인", _
                "사내", "Production", "해결됨", "중국사업팀", _
                "창저우_양산승인.docx", #8/20/2025#, 100, #4/1/2025#, #8/21/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/20/2025#, "SK온 조지아 공장 12개 라인 중 2개 ESS용 LFP 라인 배정", _
                "사외", "경쟁사", "모니터링", "전략기획팀", _
                "경쟁사분석.pdf", #8/19/2025#, 90, #8/1/2025#, #11/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/19/2025#, "현대차 아이오닉6 2026년형 배터리 단가 5% 인하 요구 대응", _
                "사내", "Marketing", "미해결", "영업마케팅팀", _
                "현대차_협상안.xlsx", #8/18/2025#, 35, #7/15/2025#, #9/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/18/2025#, "Subaru 전고체 배터리 탑재 산업용 로봇 테스트 - Maxell PSB401010H", _
                "사외", "Tech", "모니터링", "R&D센터", _
                "기술벤치마킹.pdf", #8/17/2025#, 70, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/15/2025#, "폴란드 1공장 3번 라인 화재사고 - 생산 차질 복구 계획", _
                "사내", "Production", "미해결", "안전환경팀", _
                "사고보고서.docx", #8/14/2025#, 20, #8/15/2025#, #10/15/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/14/2025#, "BMW 미국에서 중국산 배터리 도입 검토 - 82% 관세에도 경제성 판단", _
                "사외", "정책", "모니터링", "영업마케팅팀", _
                "시장분석.pdf", #8/13/2025#, 85, #8/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/12/2025#, "GM Ultium 플랫폼 Validation 기간 6개월→3개월 단축 협의 완료", _
                "사내", "R&D", "해결됨", "R&D센터", _
                "GM_협의결과.pdf", #8/11/2025#, 100, #5/1/2025#, #8/12/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/10/2025#, "사내 HR 주요 변경사항 업데이트 - 성과평가 제도 개편", _
                "사내", "Staff", "해결됨", "인사팀", _
                "HR공지.docx", #8/9/2025#, 100, #6/1/2025#, #8/10/2025#, False)
    allIssues.Add issue
    
    ' 2025년 7월 이슈들
    Set issue = CreateIssueFinalV8(#7/31/2025#, "SK온-SK엔무브 합병 - 11월 1일 통합법인 출범", _
                "사외", "경쟁사", "진행중", "전략기획팀", _
                "M&A분석.pdf", #7/30/2025#, 70, #6/1/2025#, #11/1/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/30/2025#, "3분기 실적 영업이익 3.2조원 달성 - 이사회 4분기 투자 15조원 승인", _
                "사내", "Staff", "해결됨", "경영기획팀", _
                "이사회자료.pptx", #7/29/2025#, 100, #7/1/2025#, #7/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/29/2025#, "CATL 중국 첫 전기 관광선 '위젠 77' 선박용 배터리 공급", _
                "사외", "경쟁사", "모니터링", "시장분석팀", _
                "신시장분석.pdf", #7/28/2025#, 85, #7/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/28/2025#, "테슬라 4680 배터리 연간 50GWh 공급계약 최종 협상 진행", _
                "사내", "Marketing", "진행중", "영업마케팅팀", _
                "테슬라_계약서.docx", #7/27/2025#, 85, #5/1/2025#, #9/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/25/2025#, "계열사 합병 시너지 3조원 도출 - 중복 R&D 통합 및 구매력 강화", _
                "사내", "Staff", "진행중", "경영기획팀", _
                "PMI보고서.xlsx", #7/24/2025#, 65, #5/15/2025#, #10/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/24/2025#, "LG엔솔 Sunwoda 특허침해 소송 승소 - 독일 판매금지", _
                "사외", "특허", "해결됨", "법무팀", _
                "특허분석.pdf", #7/23/2025#, 100, #5/1/2025#, #7/24/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/22/2025#, "인도 첸나이 배터리 공장 2026년 상반기 착공 최종 승인", _
                "사내", "투자", "해결됨", "해외사업팀", _
                "인도투자안.pptx", #7/21/2025#, 100, #3/1/2025#, #7/22/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/20/2025#, "울산 2공장 건식코팅 라인 화재 - 월 2GWh 생산차질 발생", _
                "사내", "Production", "미해결", "안전환경팀", _
                "사고보고서.pdf", #7/19/2025#, 30, #7/15/2025#, #9/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/18/2025#, "현대차-SK온 아이오닉5 액침냉각 배터리 선행개발 시험", _
                "사외", "Tech", "진행중", "R&D센터", _
                "기술동향.pdf", #7/17/2025#, 70, #6/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/18/2025#, "LG에너지솔루션, 미국 뉴멕시코주 600MWh급 ESS 수주", _
                "사외", "ESS", "모니터링", "ESS사업팀", _
                "ESS시장분석.pdf", #7/17/2025#, 95, #7/1/2025#, #11/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/17/2025#, "2025 상반기 독일 전기차 판매량 35% 증가 - 248,726대", _
                "사외", "시장", "모니터링", "해외사업팀", _
                "시장보고서.xlsx", #7/16/2025#, 95, #7/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/17/2025#, "기아 EV5 내수 모델에 CATL 삼원계 배터리 탑재 방침", _
                "사외", "경쟁사", "모니터링", "전략기획팀", _
                "경쟁사동향.pdf", #7/16/2025#, 90, #7/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/15/2025#, "현대기아차 리튬 연간 20만톤 공동구매 MOU 체결", _
                "사내", "Marketing", "해결됨", "구매팀", _
                "MOU계약서.pdf", #7/14/2025#, 100, #5/1/2025#, #7/15/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/10/2025#, "중국 공신부 차관 미팅 - 상하이 3공장 증설 허가 획득", _
                "사내", "Staff", "해결됨", "중국사업팀", _
                "미팅결과.docx", #7/9/2025#, 100, #5/1/2025#, #7/10/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/8/2025#, "Gotion High-Tech, 독일 괴팅겐 5MWh 액체냉각 ESS 현지생산", _
                "사외", "ESS", "모니터링", "ESS사업팀", _
                "ESS경쟁분석.pdf", #7/7/2025#, 85, #7/1/2025#, #11/30/2025#, True)
    allIssues.Add issue
    
    ' 2025년 6월 이슈들
    Set issue = CreateIssueFinalV8(#6/30/2025#, "EVE Energy 말레이시아 ESS 배터리 공장 86.5억위안 투자", _
                "사외", "투자", "모니터링", "전략기획팀", _
                "투자분석.pdf", #6/29/2025#, 90, #6/1/2025#, #11/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#6/28/2025#, "포드 F-150 Lightning 2026년형 NCM622 배터리 40GWh 계약 체결", _
                "사내", "Marketing", "해결됨", "영업마케팅팀", _
                "포드_계약완료.pdf", #6/27/2025#, 100, #4/1/2025#, #6/28/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#6/24/2025#, "BYD, 유럽에서 주행거리 700km 전기버스 eBus B13.b 공개", _
                "사외", "경쟁사", "모니터링", "시장분석팀", _
                "BYD분석.pptx", #6/23/2025#, 85, #6/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#6/20/2025#, "실리콘 음극재 5% 적용 NCM9 배터리 에너지밀도 320Wh/kg 달성", _
                "사내", "R&D", "해결됨", "R&D센터", _
                "연구성과.docx", #6/19/2025#, 100, #2/1/2025#, #6/20/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#6/9/2025#, "삼성SDI 전고체 배터리 롤프레스 장비 PO 발주", _
                "사외", "Tech", "모니터링", "R&D센터", _
                "기술분석.pdf", #6/8/2025#, 85, #6/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    ' 2025년 5월 이슈들
    Set issue = CreateIssueFinalV8(#5/19/2025#, "Gotion 전고체 배터리 300Wh/kg 전기차 1000km 시험주행", _
                "사외", "Tech", "모니터링", "R&D센터", _
                "기술벤치마킹.pdf", #5/18/2025#, 95, #5/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#5/17/2025#, "멕시코 몬테레이 신규 공장 부지 250만㎡ 매입 완료", _
                "사내", "투자", "해결됨", "해외사업팀", _
                "멕시코투자.xlsx", #5/16/2025#, 100, #3/1/2025#, #5/17/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#5/15/2025#, "LG엔솔 폴란드 브로츠와프 ESS 전용 라인 연말 가동", _
                "사외", "ESS", "모니터링", "해외사업팀", _
                "경쟁사동향.pdf", #5/14/2025#, 100, #5/1/2025#, #11/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#5/14/2025#, "태국 정부, Sunwoda의 10억 달러 배터리 공장 투자 승인", _
                "사외", "투자", "모니터링", "시장분석팀", _
                "동남아시장.pdf", #5/13/2025#, 90, #5/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#5/10/2025#, "국내 최초 LFP 배터리 양산라인 월 5GWh 가동 개시", _
                "사내", "Production", "해결됨", "생산관리팀", _
                "LFP양산보고.pptx", #5/9/2025#, 100, #1/1/2025#, #5/10/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#5/9/2025#, "CATL 9MWh 초대형 ESS 'Tener Stack' 공개 - EES 유럽 2025", _
                "사외", "ESS", "모니터링", "ESS사업팀", _
                "제품분석.pdf", #5/8/2025#, 100, #5/1/2025#, #11/30/2025#, True)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#5/5/2025#, "충북 청주 배터리 리사이클링 센터 연 10만톤 처리시설 착공", _
                "사내", "투자", "진행중", "ESS사업팀", _
                "리사이클링계획.docx", #5/4/2025#, 40, #4/1/2025#, #12/31/2025#, False)
    allIssues.Add issue
    
    ' 추가 이슈들로 54개 채우기
    Set issue = CreateIssueFinalV8(#8/30/2025#, "Stellantis 2027년형 전기트럭 배터리 25GWh 공급 협상", _
                "사내", "Marketing", "진행중", "영업마케팅팀", _
                "Stellantis_제안서.pptx", #8/29/2025#, 55, #7/1/2025#, #10/31/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#8/5/2025#, "베트남 빈패스트 VF9 모델 NCM811 배터리 15GWh 계약", _
                "사내", "Marketing", "해결됨", "영업마케팅팀", _
                "빈패스트_계약.pdf", #8/4/2025#, 100, #6/1/2025#, #8/5/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#7/5/2025#, "인도네시아 니켈 광산 지분 30% 인수 검토", _
                "사내", "투자", "진행중", "전략기획팀", _
                "니켈광산_투자검토.xlsx", #7/4/2025#, 75, #5/1/2025#, #9/30/2025#, False)
    allIssues.Add issue
    
    Set issue = CreateIssueFinalV8(#6/15/2025#, "유럽 배터리 여권(Battery Passport) 대응 시스템 구축", _
                "사내", "정책", "진행중", "품질관리팀", _
                "배터리여권_대응.docx", #6/14/2025#, 60, #4/1/2025#, #11/30/2025#, False)
    allIssues.Add issue
End Sub

Private Function CreateIssueFinalV8(issueDate As Date, title As String, category1 As String, _
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
    
    Set CreateIssueFinalV8 = issue
End Function

Sub ApplyFiltersV8(ws As Worksheet)
    Dim filter1 As String, filter2 As String, filter3 As String, filter4 As String
    Dim searchTerm As String
    
    ' 필터 값 읽기
    filter1 = ws.Range("D8").Value  ' 분류1
    filter2 = ws.Range("E8").Value  ' 세부구분
    filter3 = ws.Range("F8").Value  ' 상태 (순서 변경됨)
    filter4 = ws.Range("G8").Value  ' 담당부서 (순서 변경됨)
    searchTerm = ws.Range("C5").Value
    
    ' 필터링된 컬렉션 생성
    Set filteredIssues = New Collection
    Dim issue As Object
    Dim includeIssue As Boolean
    
    ' allIssues가 비어있으면 로드
    If allIssues Is Nothing Then
        Call LoadAllIssuesFinalV8
    End If
    
    If allIssues.Count = 0 Then
        Call LoadAllIssuesFinalV8
    End If
    
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
    Call DisplayFilteredIssuesV8(ws)
End Sub

Sub DisplayFilteredIssuesV8(ws As Worksheet)
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
        Call AddFinalIssueRowV8(ws, row, displayCount, issue)
        row = row + 1
    Next issue
    
    ' 결과 메시지
    ws.Range("K5").Value = "총 " & displayCount & "개"
    ws.Range("K5").Font.Color = IIf(displayCount = allIssues.Count, RGB(0, 128, 0), RGB(0, 0, 255))
    
    ' 테두리 적용
    If row > 11 Then
        With ws.Range("A10:Q" & (row - 1))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        ' 자동 필터 적용
        ws.Range("A10:Q" & (row - 1)).AutoFilter
    End If
End Sub

Private Sub AddFinalIssueRowV8(ws As Worksheet, row As Long, no As Integer, issue As Object)
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
    
    ' 상태 (6열로 변경)
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
    
    ' 담당부서 (7열로 변경)
    ws.Cells(row, 7).Value = issue("dept")
    ws.Cells(row, 7).HorizontalAlignment = xlCenter
    
    ' 진행률
    ws.Cells(row, 8).Value = issue("progress") & "%"
    ws.Cells(row, 8).HorizontalAlignment = xlCenter
    
    ' 문서 참조 (16열)
    With ws.Cells(row, 16)
        .Value = issue("docRef")
        .Font.Color = RGB(0, 0, 255)
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 12
    End With
    
    ' 업데이트 날짜 (17열)
    ws.Cells(row, 17).Value = Format(issue("updateDate"), "yyyy-mm-dd")
    ws.Cells(row, 17).HorizontalAlignment = xlCenter
    
    ' 모든 텍스트 12pt
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 17)).Font.Size = 12
    
    ' 타임라인 그리기
    Call DrawFinalTimelineV8(ws, row, issue)
End Sub

Private Sub DrawFinalTimelineV8(ws As Worksheet, row As Long, issue As Object)
    Dim startCol As Integer, endCol As Integer, currentCol As Integer
    Dim barColor As Long
    
    ' 날짜를 컬럼으로 변환 (2025-05 = 9열, ... 2025-11 = 15열)
    startCol = Month(issue("startDate")) - 5 + 9
    If startCol < 9 Then startCol = 9
    If startCol > 15 Then Exit Sub
    
    currentCol = Month(issue("date")) - 5 + 9
    If currentCol < 9 Then currentCol = 9
    If currentCol > 15 Then currentCol = 15
    
    endCol = Month(issue("endDate")) - 5 + 9
    If endCol > 15 Then endCol = 15
    If endCol < 9 Then endCol = 9
    
    ' 상태별 색상
    Select Case issue("status")
        Case "해결됨"
            barColor = RGB(0, 176, 80)
        Case "모니터링"
            barColor = RGB(0, 176, 240)
        Case "진행중"
            barColor = RGB(255, 220, 0)
        Case "미해결"
            barColor = RGB(255, 80, 80)
    End Select
    
    ' 타임라인 바 그리기
    If startCol <= 15 Then
        Call DrawBar(ws, row, startCol, endCol, barColor)
        
        ' 시작점에 흰색 동그라미
        If startCol >= 9 And startCol <= 15 Then
            Call AddFinalMarkerV8(ws, row, startCol, "circle")
        End If
        
        ' 상태별 마커
        If issue("status") = "해결됨" Then
            ' 해결 시점에 체크 마크
            If endCol >= 9 And endCol <= 15 Then
                Call AddFinalMarkerV8(ws, row, endCol, "check")
            End If
        ElseIf currentCol = 12 Then  ' 8월 위치
            ' 현재 위치에 상태 마커
            Select Case issue("status")
                Case "진행중"
                    Call AddFinalMarkerV8(ws, row, 12, "arrow")
                Case "모니터링"
                    Call AddFinalMarkerV8(ws, row, 12, "arrow")
                Case "미해결"
                    Call AddFinalMarkerV8(ws, row, 12, "square")
            End Select
        End If
    End If
End Sub

Private Sub DrawBar(ws As Worksheet, row As Long, startCol As Integer, _
                   endCol As Integer, barColor As Long)
    Dim cell As Range
    
    For Each cell In ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol))
        cell.Interior.Color = barColor
    Next cell
End Sub

Private Sub AddFinalMarkerV8(ws As Worksheet, row As Long, col As Integer, markerType As String)
    Dim cell As Range
    Set cell = ws.Cells(row, col)
    
    Select Case markerType
        Case "circle"
            cell.Value = "●"
            cell.Font.Color = RGB(255, 255, 255)
        Case "arrow"
            cell.Value = "▶"
            cell.Font.Color = RGB(255, 255, 255)
        Case "check"
            cell.Value = Chr(252)
            cell.Font.Name = "Wingdings"
            cell.Font.Color = RGB(255, 255, 255)
        Case "square"
            cell.Value = "■"
            cell.Font.Color = RGB(255, 255, 255)
    End Select
    
    cell.HorizontalAlignment = xlCenter
    cell.Font.Size = 14
    cell.Font.Bold = True
End Sub

' 검색 실행
Sub SearchFinalV8()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    Call ApplyFiltersV8(ws)
End Sub

' 전체 보기
Sub ShowAllFinalV8()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    
    ' 모든 필터 초기화
    ws.Range("C5").Value = ""
    ws.Range("D8").Value = "전체"
    ws.Range("E8").Value = "전체"
    ws.Range("F8").Value = "전체"
    ws.Range("G8").Value = "전체"
    
    Call ApplyFiltersV8(ws)
End Sub

' 새로고침
Sub RefreshFinalV8()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    Call ApplyFiltersV8(ws)
End Sub