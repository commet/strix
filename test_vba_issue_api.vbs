' VBA API 테스트 스크립트
Dim http, url, responseText

url = "http://localhost:5001/api/issues"
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

http.Open "GET", url, False
http.setRequestHeader "Accept", "application/json; charset=utf-8"
http.send

WScript.Echo "Status: " & http.Status
WScript.Echo "Response Length: " & Len(http.responseText)
WScript.Echo "First 500 chars: " & Left(http.responseText, 500)