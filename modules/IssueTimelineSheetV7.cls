VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IssueTimelineSheetV7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ============================================
' Issue Timeline 시트 모듈용 코드
' 이 코드를 Issue Timeline 시트 모듈에 복사하세요
' ============================================

Private Sub Worksheet_Change(ByVal Target As Range)
    ' 드롭다운 필터 변경 시 자동 필터링
    If Not Intersect(Target, Range("D8:G8")) Is Nothing Then
        Application.EnableEvents = False
        
        On Error Resume Next
        Call ApplyFiltersV7(Me)
        On Error GoTo 0
        
        Application.EnableEvents = True
    End If
End Sub