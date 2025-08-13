Attribute VB_Name = "UpdateSmartAlert"
' Smart Alert 시스템 업데이트 매크로
Option Explicit

Sub UpdateAndRefreshSmartAlert()
    ' 1. 기존 Smart Alerts 시트 삭제
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Smart Alerts").Delete
    ThisWorkbook.Sheets("Settings").Delete
    ThisWorkbook.Sheets("Email Log").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' 2. 새로 생성
    Call CreateSmartAlertSystem
    
    ' 3. 확인 메시지
    MsgBox "Smart Alert System이 업데이트되었습니다!" & vbLf & vbLf & _
           "새로운 기능:" & vbLf & _
           "1. 설정 - 실제 입력 가능" & vbLf & _
           "   - Critical 임계값 변경" & vbLf & _
           "   - 알림 주기 설정" & vbLf & _
           "   - 이메일 수신자 관리" & vbLf & vbLf & _
           "2. 이메일 전송 - 4가지 옵션" & vbLf & _
           "   - 기본 발송" & vbLf & _
           "   - 수신자 변경" & vbLf & _
           "   - 제목/본문 편집" & vbLf & _
           "   - 상세 설정" & vbLf & vbLf & _
           "테스트해보세요!", _
           vbInformation, "업데이트 완료"
End Sub

' 테스트용 - 설정 창 바로 열기
Sub TestSettings()
    Call ShowAlertSettings
End Sub

' 테스트용 - 이메일 창 바로 열기  
Sub TestEmail()
    Call SendAlertEmail
End Sub