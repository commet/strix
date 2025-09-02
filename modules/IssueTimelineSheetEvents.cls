VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IssueTimelineSheetEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ============================================
' Issue Timeline 시트 이벤트 핸들러
' 이 코드를 Issue Timeline 시트 모듈에 복사하세요
' ============================================

Private Sub Worksheet_Change(ByVal Target As Range)
    ' 필터 드롭다운 변경 감지
    If Not Intersect(Target, Range("D8:G8")) Is Nothing Then
        Application.EnableEvents = False
        
        ' 필터 적용
        On Error Resume Next
        Call ApplyFiltersV8(Me)
        On Error GoTo 0
        
        Application.EnableEvents = True
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' 검색창 선택 시 처리
    If Target.Address = "$C$5" Then
        ' 검색창이 비어있으면 포커스 설정
        If Target.Value = "" Then
            Target.Select
        End If
    End If
End Sub

' ============================================
' 설치 방법:
' 1. VBA 편집기에서 Issue Timeline 시트 모듈 더블클릭
' 2. 이 코드 전체를 복사하여 붙여넣기
' 3. 저장
' ============================================