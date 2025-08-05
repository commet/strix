# STRIX VBA 사용 가이드

## 1. 필수 준비사항

### 1.1 JSON 파서 설치
VBA에서 JSON을 처리하기 위해 `JsonConverter.bas`가 필요합니다.
1. https://github.com/VBA-tools/VBA-JSON 에서 다운로드
2. VBA 편집기에서 가져오기

### 1.2 참조 설정
VBA 편집기에서 도구 > 참조:
- Microsoft XML, v6.0
- Microsoft Scripting Runtime

### 1.3 Streamlit API 서버 실행
```bash
streamlit run streamlit_app.py
```

## 2. VBA 모듈 가져오기

1. Excel VBA 편집기 열기 (Alt + F11)
2. 파일 > 가져오기:
   - `modSTRIXAPI.bas`
   - `modSTRIXExcel.bas`
   - `frmSTRIX.frm`

## 3. 주요 기능

### 3.1 STRIX 대화창 열기
```vba
Sub OpenSTRIX()
    ShowSTRIXDialog
End Sub
```

### 3.2 셀 함수로 사용
```excel
=STRIX("전고체 배터리 개발 현황은?")
```

### 3.3 선택 영역 분석
```vba
Sub AnalyzeSelection()
    AskAboutSelection
End Sub
```

### 3.4 문서 업로드
```vba
Sub UploadDocs()
    BulkUploadDocuments
End Sub
```

## 4. 리본 메뉴 추가

### 4.1 CustomUI XML
```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="STRIXTab" label="STRIX">
        <group id="STRIXGroup" label="Intelligence">
          <button id="btnSTRIX" 
                  label="STRIX 열기" 
                  size="large" 
                  onAction="ShowSTRIXDialog" 
                  imageMso="DataAnalysis" />
          <button id="btnAnalyze" 
                  label="선택 분석" 
                  size="normal" 
                  onAction="AskAboutSelection" 
                  imageMso="ZoomToSelection" />
          <button id="btnUpload" 
                  label="문서 업로드" 
                  size="normal" 
                  onAction="BulkUploadDocuments" 
                  imageMso="FileUpload" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

## 5. 사용 예시

### 5.1 기본 질문
```vba
Dim answer As String
answer = AskSTRIX("최근 배터리 시장 동향은?")
MsgBox answer
```

### 5.2 문서 타입 지정
```vba
' 내부 문서만 검색
answer = AskSTRIX("우리 회사 전략은?", "internal")

' 외부 뉴스만 검색
answer = AskSTRIX("경쟁사 동향은?", "external")
```

### 5.3 검색 결과 시트에 표시
```vba
Call DisplayAnswer(answer, "배터리 시장 분석")
```

## 6. 주의사항

1. **Streamlit 서버**: VBA 사용 전 반드시 Streamlit 앱이 실행 중이어야 함
2. **네트워크**: localhost:8501 접근 가능해야 함
3. **보안**: API 키는 서버 측에서 관리됨

## 7. 문제 해결

### 오류: "Error: Connection refused"
- Streamlit 서버가 실행 중인지 확인
- 방화벽 설정 확인

### 오류: "JSON parsing error"
- JsonConverter.bas가 제대로 import 되었는지 확인
- 참조 설정 확인

### 한글 깨짐
- UTF-8 인코딩 확인
- Excel 파일 저장 시 인코딩 설정