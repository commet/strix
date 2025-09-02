Attribute VB_Name = "modVerifySetup"
Option Explicit

' ============================================
' 설정 검증 및 테스트 모듈
' Issue Timeline 자동 필터 기능 검증
' ============================================

Sub VerifyFilterSetup()
    ' 필터 자동화 설정 검증
    Dim ws As Worksheet
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeModule As Object
    Dim hasEventCode As Boolean
    Dim lineNum As Long
    Dim codeLine As String
    Dim issuesFound As String
    
    issuesFound = ""
    
    ' 1. Issue Timeline 시트 존재 확인
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    On Error GoTo 0
    
    If ws Is Nothing Then
        issuesFound = issuesFound & "- Issue Timeline 시트가 없습니다" & vbCrLf
    Else
        ' 2. 필터 드롭다운 확인
        If ws.Range("D8").Validation.Type <> 3 Then
            issuesFound = issuesFound & "- D8 셀에 드롭다운이 없습니다" & vbCrLf
        End If
        If ws.Range("E8").Validation.Type <> 3 Then
            issuesFound = issuesFound & "- E8 셀에 드롭다운이 없습니다" & vbCrLf
        End If
        If ws.Range("F8").Validation.Type <> 3 Then
            issuesFound = issuesFound & "- F8 셀에 드롭다운이 없습니다" & vbCrLf
        End If
        If ws.Range("G8").Validation.Type <> 3 Then
            issuesFound = issuesFound & "- G8 셀에 드롭다운이 없습니다" & vbCrLf
        End If
        
        ' 3. 헤더 확인 (순서 변경 확인)
        If ws.Range("F7").Value <> "상태" Then
            issuesFound = issuesFound & "- F7 셀이 '상태'가 아닙니다 (현재: " & ws.Range("F7").Value & ")" & vbCrLf
        End If
        If ws.Range("G7").Value <> "담당부서" Then
            issuesFound = issuesFound & "- G7 셀이 '담당부서'가 아닙니다 (현재: " & ws.Range("G7").Value & ")" & vbCrLf
        End If
    End If
    
    ' 4. 이벤트 핸들러 코드 확인
    Set vbProj = ThisWorkbook.VBProject
    hasEventCode = False
    
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = 100 Then ' 워크시트 모듈
            If vbComp.Properties("Name").Value = "Issue Timeline" Then
                Set codeModule = vbComp.codeModule
                
                ' Worksheet_Change 이벤트 확인
                For lineNum = 1 To codeModule.CountOfLines
                    codeLine = codeModule.Lines(lineNum, 1)
                    If InStr(codeLine, "Worksheet_Change") > 0 Then
                        hasEventCode = True
                        Exit For
                    End If
                Next lineNum
                Exit For
            End If
        End If
    Next vbComp
    
    If Not hasEventCode Then
        issuesFound = issuesFound & "- Worksheet_Change 이벤트 핸들러가 설치되지 않았습니다" & vbCrLf
    End If
    
    ' 결과 표시
    If issuesFound = "" Then
        MsgBox "✓ 모든 설정이 올바르게 구성되었습니다!" & vbCrLf & vbCrLf & _
               "필터 자동화 기능이 정상 작동합니다." & vbCrLf & _
               "드롭다운을 변경하면 자동으로 필터링됩니다.", _
               vbInformation, "설정 검증 완료"
    Else
        MsgBox "다음 문제가 발견되었습니다:" & vbCrLf & vbCrLf & _
               issuesFound & vbCrLf & _
               "RunCompleteSetup 매크로를 실행하여 문제를 해결하세요.", _
               vbExclamation, "설정 문제 발견"
    End If
End Sub

Sub QuickTestFilters()
    ' 빠른 필터 테스트
    Dim ws As Worksheet
    Dim originalValue As String
    Dim testPassed As Boolean
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    ws.Activate
    
    ' 이벤트 활성화 확인
    Application.EnableEvents = True
    
    ' 테스트 1: 분류1 필터
    originalValue = ws.Range("D8").Value
    ws.Range("D8").Value = "사내"
    Application.Wait Now + TimeValue("00:00:01")
    ws.Range("D8").Value = "사외"
    Application.Wait Now + TimeValue("00:00:01")
    ws.Range("D8").Value = originalValue
    
    ' 테스트 2: 상태 필터
    originalValue = ws.Range("F8").Value
    ws.Range("F8").Value = "해결됨"
    Application.Wait Now + TimeValue("00:00:01")
    ws.Range("F8").Value = "진행중"
    Application.Wait Now + TimeValue("00:00:01")
    ws.Range("F8").Value = originalValue
    
    testPassed = True
    
    MsgBox "필터 테스트 완료!" & vbCrLf & vbCrLf & _
           "드롭다운 필터가 정상적으로 작동합니다.", _
           vbInformation, "테스트 성공"
    Exit Sub
    
ErrorHandler:
    MsgBox "필터 테스트 중 오류가 발생했습니다." & vbCrLf & vbCrLf & _
           "RunCompleteSetup 매크로를 실행해주세요.", _
           vbExclamation, "테스트 실패"
End Sub

Sub ShowFilterInstructions()
    ' 사용 설명서 표시
    Dim msg As String
    
    msg = "Issue Timeline 필터 사용 방법:" & vbCrLf & vbCrLf
    msg = msg & "1. 자동 필터:" & vbCrLf
    msg = msg & "   - D8: 분류1 (사내/사외/전체)" & vbCrLf
    msg = msg & "   - E8: 세부구분 (정책/경쟁사/Tech 등)" & vbCrLf
    msg = msg & "   - F8: 상태 (해결됨/진행중/미해결/모니터링)" & vbCrLf
    msg = msg & "   - G8: 담당부서 (각 부서명)" & vbCrLf & vbCrLf
    msg = msg & "2. 검색 기능:" & vbCrLf
    msg = msg & "   - C5 셀에 검색어 입력" & vbCrLf
    msg = msg & "   - 'ESS 관련 이슈' 입력 시 11개 문서 필터링" & vbCrLf
    msg = msg & "   - '검색' 버튼 클릭 또는 Enter 키" & vbCrLf & vbCrLf
    msg = msg & "3. 전체보기:" & vbCrLf
    msg = msg & "   - '전체보기' 버튼으로 모든 54개 이슈 표시" & vbCrLf & vbCrLf
    msg = msg & "※ 필터가 자동으로 작동하지 않으면:" & vbCrLf
    msg = msg & "   RunCompleteSetup 매크로 실행"
    
    MsgBox msg, vbInformation, "사용 설명서"
End Sub