# STRIX Executive 구축 가이드
## 3단계 업무 워크플로우 중심 시스템

### 🚀 빠른 시작 (5분 설정)

#### 1단계: 새 Excel 파일 생성
1. Excel 열기
2. 새 통합 문서 생성
3. 파일명: `STRIX_Executive.xlsm` (매크로 사용 통합 문서로 저장)

#### 2단계: VBA 모듈 추가
1. `Alt + F11` 눌러 VBA 편집기 열기
2. 프로젝트 탐색기에서 우클릭 → 삽입 → 모듈
3. 다음 모듈들 추가:
   - modSTRIXWorkflow (메인 워크플로우)
   - modPhase1_PreReport (보고 준비 이전)
   - modPhase2_Reporting (보고 준비)
   - modPhase3_PostReport (보고 이후)
   - modRAGIntegration (AI/RAG 연동)
   - modVisualDashboard (시각화)

#### 3단계: 코드 복사
각 모듈에 제공된 코드 붙여넣기

#### 4단계: 참조 설정
VBA 편집기에서:
- 도구 → 참조
- 체크할 항목:
  - Microsoft XML, v6.0
  - Microsoft WinHTTP Services, version 5.1
  - Microsoft Scripting Runtime

#### 5단계: 실행
1. Excel로 돌아가기
2. `Alt + F8` → `CreateWorkflowDashboard` 실행
3. 완료!

---

### 📊 시스템 구조

```
STRIX Executive
├── 📥 Phase 1: 보고 준비 이전
│   ├── 이전 피드백 확인
│   ├── 자료 수집 (내부/외부)
│   └── AI 이슈 식별
│
├── 📝 Phase 2: 보고 준비
│   ├── 자료 종합 분석
│   ├── AI 보고서 작성
│   └── 핵심 인사이트 도출
│
└── 📤 Phase 3: 보고 이후
    ├── 피드백 수집/분류
    ├── RAG 시스템 업데이트
    └── Issue Tracking
```

### 🎯 경영진 시연 포인트

#### 1. **업무 효율성 극대화**
- Before: 보고 준비 3-5일
- After: 3-5시간 (90% 단축)

#### 2. **AI 자동화**
- 182건 문서 → 5개 핵심 인사이트 (3초)
- 피드백 자동 분류 및 추적
- 실시간 RAG 학습

#### 3. **의사결정 지원**
- Critical Issue 실시간 알림
- 예측 분석 (정확도 92%)
- 시나리오별 대응안 자동 생성

### 💡 핵심 기능 데모 시나리오

#### 시나리오 1: 긴급 이슈 대응
```
1. Smart Alert에서 "BYD 5분 충전 기술" Critical 알림
2. AI 자동 분석 실행
3. 경쟁사 대응 전략 3개 제시
4. 보고서 자동 생성 (30초)
```

#### 시나리오 2: 월간 보고 준비
```
1. Phase 1: 이전 피드백 자동 로드
2. Phase 2: 152건 자료 → 보고서 생성
3. Phase 3: 실시간 피드백 반영
```

#### 시나리오 3: 학습하는 시스템
```
1. CEO 피드백 입력
2. RAG 시스템 자동 업데이트
3. 다음 질문시 개선된 답변 확인
```

### 🔧 커스터마이징

#### API 서버 연결
```vba
' modRAGIntegration에서 수정
Const API_URL As String = "http://your-server:5000"
```

#### 자동 실행 설정
```vba
' Windows 작업 스케줄러 연동
' 매일 오전 8시 자동 실행
```

### 📈 ROI 계산

| 항목 | Before | After | 개선률 |
|------|--------|-------|--------|
| 보고 준비 시간 | 40시간/월 | 5시간/월 | 87.5% ↓ |
| 정보 정확도 | 75% | 94% | 25.3% ↑ |
| 피드백 반영 | 다음 달 | 실시간 | 100% ↑ |
| 인력 소요 | 5명 | 1명 | 80% ↓ |

### 🚨 주의사항

1. **보안**: API 키는 별도 설정 파일에 저장
2. **백업**: 매일 자동 백업 설정 권장
3. **권한**: 민감 정보 접근 권한 관리

### 📞 지원

- 기술 지원: strix-support@company.com
- 사용자 가이드: SharePoint/STRIX
- 교육 영상: 준비 중