# 🚨 STRIX Smart Alert System 사용 가이드

## 개요
Smart Alert System은 AI 기반으로 사내 이슈를 실시간 분석하고, 위험도를 예측하여 자동으로 알림을 제공하는 시스템입니다.

## 주요 기능

### 1. 실시간 이슈 모니터링
- **위험도 점수화**: 0-100점 스케일로 각 이슈의 위험도 자동 계산
- **Critical 알림**: 위험도 70% 이상 이슈 즉시 알림
- **예측 분석**: 향후 72시간 이슈 발전 가능성 예측

### 2. 자동 일일 브리핑
- 매일 오전 9시 자동 실행
- TOP 5 Critical Issues 요약
- 부서별 맞춤 액션 아이템 생성

### 3. 통합 대시보드
- **통계 요약**: Critical/High/Medium/Low 이슈 현황
- **액션 트래커**: 진행중인 대응 조치 추적
- **알림 히스토리**: 과거 알림 이력 관리

## 설치 및 실행

### Excel에서 실행
1. STRIX_System.xlsm 파일 열기
2. 개발 도구 > Visual Basic 열기
3. modSmartAlert 모듈 확인
4. 매크로 실행: `CreateSmartAlertDashboard`

### 자동 실행 설정
```vba
' ThisWorkbook 모듈에 추가
Private Sub Workbook_Open()
    If Hour(Now) = 9 And Minute(Now) < 5 Then
        Call DailyAutoRun
    End If
End Sub
```

## 사용법

### 즉시 분석 실행
1. Smart Alerts 시트 이동
2. "▶️ 즉시 분석" 버튼 클릭
3. AI 분석 결과 확인

### 알림 설정 변경
1. "⚙️ 설정" 버튼 클릭
2. Critical 임계값 조정 (기본: 70%)
3. 알림 대상 설정

### 이메일 알림
1. "📧 이메일 전송" 버튼 클릭
2. 수신자 목록 확인
3. 발송 확인

## 위험도 레벨

| 레벨 | 점수 | 색상 | 대응 |
|------|------|------|------|
| Critical | 90-100 | 🔴 빨강 | 즉시 대응 필수 |
| High | 70-89 | 🟠 주황 | 24시간 내 대응 |
| Medium | 50-69 | 🟡 노랑 | 주간 단위 모니터링 |
| Low | 0-49 | 🟢 초록 | 월간 검토 |

## 액션 아이템 관리

### 우선순위
- **Critical**: 즉시 실행 (당일)
- **High**: 긴급 (3일 내)
- **Medium**: 일반 (1주일 내)
- **Low**: 장기 과제

### 진행 상태
- 대기: 시작 전
- 진행중: 수행 중
- 완료: 종료
- 보류: 일시 중단

## API 연동

### 예측 API
```python
# api_server_with_issues.py 실행 필요
POST http://localhost:5000/api/issues/predict
```

### 데이터 구조
```json
{
  "issue": "이슈 내용",
  "risk_score": 85,
  "prediction": "72시간 내 Critical 전환 가능성",
  "recommended_action": "긴급 TF 구성"
}
```

## 트러블슈팅

### 일반적인 문제
1. **API 연결 오류**: api_server_with_issues.py 실행 확인
2. **자동 실행 안됨**: Windows 작업 스케줄러 설정 확인
3. **이메일 발송 실패**: Outlook 설정 및 권한 확인

### 로그 확인
- Excel: Smart Alerts 시트 하단 알림 히스토리
- API: terminal에서 서버 로그 확인

## 주의사항
- Critical 이슈는 반드시 당일 내 확인 및 대응
- 자동 알림 설정 시 이메일 수신 설정 필수
- 주기적인 임계값 조정으로 알림 최적화

## 문의
- 시스템 관련: IT팀
- 이슈 분석: 전략기획팀
- AI 예측: 데이터분석팀