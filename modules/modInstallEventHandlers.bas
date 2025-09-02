Attribute VB_Name = "modInstallEventHandlers"
Option Explicit

' ============================================
' 워크시트 이벤트 핸들러 자동 설치 모듈
' Issue Timeline 시트에 이벤트 핸들러를 자동으로 설치합니다
' ============================================

Sub InstallIssueTimelineEvents()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeModule As Object
    Dim ws As Worksheet
    Dim eventCode As String
    
    ' Issue Timeline 시트 찾기
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Issue Timeline 시트를 먼저 생성해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' VBA 프로젝트 접근
    Set vbProj = ThisWorkbook.VBProject
    
    ' Issue Timeline 시트의 코드 모듈 찾기
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = 100 Then ' 워크시트 모듈
            If vbComp.Properties("Name").Value = "Issue Timeline" Then
                Set codeModule = vbComp.codeModule
                Exit For
            End If
        End If
    Next vbComp
    
    If codeModule Is Nothing Then
        MsgBox "Issue Timeline 시트 모듈을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 이벤트 핸들러 코드 정의
    eventCode = "' ============================================" & vbCrLf & _
                "' Issue Timeline 시트 이벤트 핸들러" & vbCrLf & _
                "' 자동 필터 적용을 위한 코드" & vbCrLf & _
                "' ============================================" & vbCrLf & vbCrLf & _
                "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
                "    ' 필터 드롭다운 변경 감지 (D8:G8)" & vbCrLf & _
                "    If Not Intersect(Target, Range(""D8:G8"")) Is Nothing Then" & vbCrLf & _
                "        Application.EnableEvents = False" & vbCrLf & _
                "        " & vbCrLf & _
                "        ' 필터 적용" & vbCrLf & _
                "        On Error Resume Next" & vbCrLf & _
                "        Call ApplyFiltersV8(Me)" & vbCrLf & _
                "        On Error GoTo 0" & vbCrLf & _
                "        " & vbCrLf & _
                "        Application.EnableEvents = True" & vbCrLf & _
                "    End If" & vbCrLf & _
                "End Sub" & vbCrLf & vbCrLf & _
                "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & _
                "    ' 검색창 선택 시 처리" & vbCrLf & _
                "    If Target.Address = ""$C$5"" Then" & vbCrLf & _
                "        ' 검색창 선택 시 포커스 설정" & vbCrLf & _
                "        If Target.Value = """" Then" & vbCrLf & _
                "            Target.Select" & vbCrLf & _
                "        End If" & vbCrLf & _
                "    End If" & vbCrLf & _
                "End Sub"
    
    ' 기존 코드 삭제
    With codeModule
        If .CountOfLines > 0 Then
            .DeleteLines 1, .CountOfLines
        End If
        
        ' 새 코드 추가
        .InsertLines 1, eventCode
    End With
    
    MsgBox "Issue Timeline 시트 이벤트 핸들러가 성공적으로 설치되었습니다!" & vbCrLf & vbCrLf & _
           "이제 드롭다운 필터가 자동으로 작동합니다.", vbInformation
End Sub

Sub RunCompleteSetup()
    ' 전체 설정을 한번에 실행
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Issue Timeline 대시보드를 완전히 설정하시겠습니까?" & vbCrLf & vbCrLf & _
                      "다음 작업이 수행됩니다:" & vbCrLf & _
                      "1. Issue Timeline 대시보드 생성" & vbCrLf & _
                      "2. 이벤트 핸들러 자동 설치" & vbCrLf & _
                      "3. 필터 자동화 활성화", _
                      vbYesNo + vbQuestion, "완전 설정")
    
    If response = vbYes Then
        ' 대시보드 생성
        Call CreateFinalDashboardV8
        
        ' 이벤트 핸들러 설치
        Call InstallIssueTimelineEvents
        
        ' 초기 데이터 로드
        Call RefreshFinalV8
        
        MsgBox "Issue Timeline 대시보드가 완전히 설정되었습니다!" & vbCrLf & vbCrLf & _
               "드롭다운 필터가 자동으로 작동합니다.", vbInformation
    End If
End Sub

Sub TestFilterAutomation()
    ' 필터 자동화 테스트
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Issue Timeline 시트를 먼저 생성해주세요.", vbExclamation
        Exit Sub
    End If
    
    MsgBox "필터 자동화 테스트를 시작합니다." & vbCrLf & vbCrLf & _
           "각 드롭다운을 변경하면서 테스트하세요:" & vbCrLf & _
           "- D8: 분류1 (사내/사외)" & vbCrLf & _
           "- E8: 세부구분 (정책/경쟁사/Tech 등)" & vbCrLf & _
           "- F8: 상태 (해결됨/진행중 등)" & vbCrLf & _
           "- G8: 담당부서", vbInformation
    
    ' 시트 활성화
    ws.Activate
    
    ' 첫 번째 필터 선택
    ws.Range("D8").Select
End Sub