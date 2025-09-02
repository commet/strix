Attribute VB_Name = "modSimpleSetupV7"
Option Explicit

' ============================================
' 간단한 설정 - V7 필터 자동화
' ============================================

Sub SetupV7WithFilters()
    ' V7 대시보드 생성 및 필터 자동화 설정
    Dim ws As Worksheet
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeModule As Object
    
    ' 1. 대시보드 생성
    Call CreateFinalDashboardV7
    
    ' 2. Issue Timeline 시트 찾기
    Set ws = ThisWorkbook.Worksheets("Issue Timeline")
    
    ' 3. 시트 모듈에 이벤트 코드 추가
    Set vbProj = ThisWorkbook.VBProject
    
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = 100 Then ' 워크시트 모듈
            If vbComp.Properties("Name").Value = "Issue Timeline" Then
                Set codeModule = vbComp.codeModule
                
                ' 기존 코드 삭제
                If codeModule.CountOfLines > 0 Then
                    codeModule.DeleteLines 1, codeModule.CountOfLines
                End If
                
                ' 새 이벤트 코드 추가
                codeModule.InsertLines 1, "Private Sub Worksheet_Change(ByVal Target As Range)"
                codeModule.InsertLines 2, "    ' 드롭다운 필터 자동 적용"
                codeModule.InsertLines 3, "    If Not Intersect(Target, Range(""D8:G8"")) Is Nothing Then"
                codeModule.InsertLines 4, "        Application.EnableEvents = False"
                codeModule.InsertLines 5, "        On Error Resume Next"
                codeModule.InsertLines 6, "        Call ApplyFiltersV7(Me)"
                codeModule.InsertLines 7, "        On Error GoTo 0"
                codeModule.InsertLines 8, "        Application.EnableEvents = True"
                codeModule.InsertLines 9, "    End If"
                codeModule.InsertLines 10, "End Sub"
                
                Exit For
            End If
        End If
    Next vbComp
    
    ' 4. 초기 데이터 로드
    Call RefreshFinalV7
    
    MsgBox "V7 필터 자동화 설정 완료!" & vbCrLf & vbCrLf & _
           "드롭다운을 변경하면 자동으로 필터링됩니다.", vbInformation
End Sub