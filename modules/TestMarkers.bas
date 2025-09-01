Attribute VB_Name = "TestMarkers"
' Test Module for Timeline Markers
Option Explicit

Sub TestTimelineMarkers()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    ' 테스트 영역
    Dim testRow As Integer
    testRow = 40
    
    ' 테스트 영역 지우기
    ws.Range("B" & testRow & ":M" & testRow + 5).Clear
    
    ' 제목
    ws.Range("B" & testRow).Value = "마커 테스트:"
    ws.Range("B" & testRow).Font.Bold = True
    
    ' 시작점 마커 테스트 - Wingdings 3
    ws.Range("D" & testRow + 1).Value = "시작:"
    ws.Range("E" & testRow + 1).Interior.Color = RGB(255, 165, 0)
    ws.Range("E" & testRow + 1).Value = Chr(16)  ' ▶
    ws.Range("E" & testRow + 1).Font.Name = "Wingdings 3"
    ws.Range("E" & testRow + 1).Font.Color = RGB(255, 255, 255)
    ws.Range("E" & testRow + 1).Font.Bold = True
    ws.Range("E" & testRow + 1).Font.Size = 11
    
    ' 대체 시작점 마커 - 일반 문자
    ws.Range("F" & testRow + 1).Interior.Color = RGB(255, 165, 0)
    ws.Range("F" & testRow + 1).Value = "▶"
    ws.Range("F" & testRow + 1).Font.Color = RGB(255, 255, 255)
    ws.Range("F" & testRow + 1).Font.Bold = True
    ws.Range("F" & testRow + 1).Font.Size = 11
    
    ' 진행중 마커 테스트
    ws.Range("D" & testRow + 2).Value = "진행:"
    ws.Range("E" & testRow + 2).Interior.Color = RGB(255, 165, 0)
    ws.Range("E" & testRow + 2).Value = Chr(110)  ' ■
    ws.Range("E" & testRow + 2).Font.Name = "Wingdings"
    ws.Range("E" & testRow + 2).Font.Color = RGB(255, 255, 255)
    ws.Range("E" & testRow + 2).Font.Bold = True
    
    ' 대체 진행중 마커
    ws.Range("F" & testRow + 2).Interior.Color = RGB(255, 165, 0)
    ws.Range("F" & testRow + 2).Value = "■"
    ws.Range("F" & testRow + 2).Font.Color = RGB(255, 255, 255)
    ws.Range("F" & testRow + 2).Font.Bold = True
    
    ' 완료 마커 테스트
    ws.Range("D" & testRow + 3).Value = "완료:"
    ws.Range("E" & testRow + 3).Interior.Color = RGB(0, 128, 0)
    ws.Range("E" & testRow + 3).Value = Chr(252)  ' ✓
    ws.Range("E" & testRow + 3).Font.Name = "Wingdings"
    ws.Range("E" & testRow + 3).Font.Color = RGB(255, 255, 255)
    ws.Range("E" & testRow + 3).Font.Bold = True
    
    ' 대체 완료 마커
    ws.Range("F" & testRow + 3).Interior.Color = RGB(0, 128, 0)
    ws.Range("F" & testRow + 3).Value = "✓"
    ws.Range("F" & testRow + 3).Font.Color = RGB(255, 255, 255)
    ws.Range("F" & testRow + 3).Font.Bold = True
    
    MsgBox "40번 행 주변에 마커 테스트가 표시되었습니다." & vbCrLf & _
           "E열: Wingdings 폰트 사용" & vbCrLf & _
           "F열: 일반 유니코드 문자", vbInformation
End Sub