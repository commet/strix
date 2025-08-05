# STRIX 설치 가이드

## 1. 설치 준비

### 필요 환경
- Microsoft Excel 2016 이상 (Windows)
- VBA 매크로 실행 허용
- 약 50MB의 저장 공간

## 2. 설치 단계

### Step 1: 새 Excel 파일 생성
1. Excel을 실행하고 새 통합 문서 생성
2. 파일명을 `STRIX.xlsm`으로 저장 (매크로 사용 통합 문서)

### Step 2: VBA 모듈 Import
1. `Alt + F11`을 눌러 VBA 편집기 열기
2. 프로젝트 탐색기에서 `VBAProject (STRIX.xlsm)` 우클릭
3. `가져오기(Import File...)` 선택
4. 다음 순서로 모듈 파일 Import:
   - `modules\modConfig.bas`
   - `modules\modInit.bas`

### Step 3: 초기화 실행
1. VBA 편집기에서 `F5` 키를 누르거나 `실행 > 매크로 실행`
2. `InitializeSTRIX` 매크로 선택 후 실행
3. "STRIX 초기화 완료!" 메시지 확인

### Step 4: Mock 데이터 생성 (선택사항)
1. VBA 편집기에서 `CreateMockData` 매크로 실행
2. 자동으로 테스트용 데이터 폴더와 파일들이 생성됨

## 3. 초기 설정

### Config 시트 설정
1. `Config` 시트로 이동
2. 다음 항목 설정:
   - **내부문서 폴더**: 스캔할 내부 문서 폴더 경로
   - **외부뉴스 폴더**: 뉴스 메일을 저장할 폴더 경로
   - **스캔 주기**: 자동 스캔 간격 (기본 60분)

### 키워드 설정
1. `Config` 시트의 F열에서 카테고리별 키워드 수정
2. 쉼표로 구분하여 여러 키워드 입력 가능

## 4. 사용 시작

### 수동 실행
- **내부 문서 스캔**: `개발자 > 매크로 > ScanInternalFolder`
- **외부 뉴스 스캔**: `개발자 > 매크로 > ScanExternalFolder`

### 자동 실행 설정
- Excel 파일을 열 때마다 자동으로 설정된 주기에 따라 스캔 실행

## 5. 문제 해결

### 매크로 실행 오류
1. `파일 > 옵션 > 보안 센터 > 보안 센터 설정`
2. `매크로 설정 > 모든 매크로 사용` 선택

### 폴더 접근 오류
- 지정한 폴더 경로가 정확한지 확인
- 해당 폴더에 대한 읽기/쓰기 권한이 있는지 확인

### 초기화 실패
1. 기존 시트 이름이 충돌하는지 확인
2. VBA 편집기에서 `디버그 > 컴파일` 실행하여 오류 확인

## 6. 다음 단계

Phase 2 모듈들이 준비되면 추가 Import하여 기능을 확장할 수 있습니다:
- `modInternalIngest.bas`: 내부 문서 스캔
- `modExternalIngest.bas`: 외부 뉴스 스캔
- `modCorrelation.bas`: 연관도 계산

---

문의사항이 있으시면 STRIX 프로젝트 팀에 연락주세요.