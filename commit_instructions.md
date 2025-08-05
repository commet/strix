# STRIX 프로젝트 Commit/Push 가이드

## 1. Git 초기 설정 (처음 한 번만)
```bash
cd C:\Users\admin\documents\github\strix
git init
git remote add origin https://github.com/YOUR_USERNAME/strix.git
```

## 2. 현재 변경사항 확인
```bash
git status
```

## 3. 모든 변경사항 추가
```bash
git add .
```

## 4. Commit 생성
```bash
git commit -m "STRIX Excel VBA Integration 완성

- Excel Dashboard 자동 생성 기능
- 한글 UTF-8 인코딩 완벽 지원
- API 서버 통신 (Flask)
- VBA 모듈 3개 완성:
  - Module1: 기본 테스트
  - Module2: API 통신 및 한글 처리
  - Module3: Dashboard UI 생성
- 빠른 질문 템플릿 추가
- RAG 검색 시스템 통합"
```

## 5. Push to GitHub
```bash
git push -u origin main
```

## 주요 추가/수정된 파일들:
- `api_server_korean.py` - 한글 지원 API 서버
- `CreateDashboard.bas` - Dashboard 자동 생성
- `QuickSetup.bas` - 빠른 설정 도구
- `modSTRIXexcel.bas` - Excel 통합 모듈
- `STRIX_Dashboard_설정가이드.md` - 설정 문서

## 보안 주의사항:
- GitHub 토큰은 절대 코드나 커밋에 포함시키지 마세요
- `.env` 파일에 민감한 정보 저장
- `.gitignore`에 `.env` 추가 필수