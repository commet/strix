Attribute VB_Name = "TestColorSimple"
' Simple Color Test Module
Option Explicit

Sub TestSimpleColors()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Issue Timeline")
    
    MsgBox "간단한 색상 테스트를 시작합니다", vbInformation
    
    ' 테스트 영역 지우기
    ws.Range("B20:K25").Clear
    
    ' 제목
    ws.Range("B20").Value = "색상 테스트"
    ws.Range("B20").Font.Bold = True
    
    ' 빨간색 테스트
    ws.Range("C21").Value = "빨강"
    ws.Range("D21:F21").Interior.Color = RGB(255, 0, 0)
    ws.Range("E21").Value = "●"
    ws.Range("E21").Font.Color = RGB(255, 255, 255)
    
    ' 오렌지색 테스트
    ws.Range("C22").Value = "오렌지"
    ws.Range("D22:F22").Interior.Color = RGB(255, 165, 0)
    ws.Range("E22").Value = "●"
    ws.Range("E22").Font.Color = RGB(255, 255, 255)
    
    ' 초록색 테스트
    ws.Range("C23").Value = "초록"
    ws.Range("D23:F23").Interior.Color = RGB(0, 128, 0)
    ws.Range("E23").Value = "●"
    ws.Range("E23").Font.Color = RGB(255, 255, 255)
    
    ' 파란색 테스트
    ws.Range("C24").Value = "파랑"
    ws.Range("D24:F24").Interior.Color = RGB(0, 0, 255)
    ws.Range("E24").Value = "●"
    ws.Range("E24").Font.Color = RGB(255, 255, 255)
    
    MsgBox "색상 테스트 완료!" & vbCrLf & _
           "20-24행에 색상이 표시되었는지 확인하세요", vbInformation
End Sub