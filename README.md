# PPC 조사연구 도우미

캄보디아·베트남 주요 뉴스 사이트에서 조사연구 자료를 수집하고, `ppc_report_template.docx` 양식에 맞춰 Word 보고서를 생성하는 로컬 Windows 앱입니다.

## 실행

개발/테스트용 실행:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_app.ps1
```

앱은 기본적으로 아래 주소에서 열립니다.

```text
http://127.0.0.1:8766
```

## 기본 설정

- 시작일: 실행일 기준 2주 전
- 종료일: 실행일
- 사이트별 수집 건수: 10건
- 보고서 템플릿: `%LOCALAPPDATA%\PPCResearchAutomation\resources\ppc_report_template.docx`
- OpenAI API Key: 앱에서 별도 저장

## 사용자 데이터

앱 데이터는 기존 조사연구/금융권 앱과 분리해 아래 위치에 저장합니다.

```text
%LOCALAPPDATA%\PPCResearchAutomation
```

주요 하위 경로:

- `config\openai_key.bin`: OpenAI API Key
- `resources\ppc_report_template.docx`: 보고서 템플릿
- `resources\report_prompt.md`: 보고서 bullet 생성 프롬프트
- `runs\YYMMDD_HHMM`: 실행 기록, metadata, 생성 보고서

## 수집 대상

고정 매핑 사이트:

- 캄보디아 인사이트
- 캄푸치아 신문
- 인사이드비나

자동 분류 보조 사이트:

- Khmer Times
- Phnom Penh Post
- 베트남 코리아 타임즈
- 시티타임즈

섹션은 다음 4개로 정리됩니다.

- 캄보디아 금융/경제
- 캄보디아 정치/사회
- 베트남 금융/경제
- 베트남 정치/사회

## 보고서 생성

- 선택된 항목 전체를 OpenAI에 전달해 기사별 bullet summary를 생성합니다.
- LLM 응답에서 누락된 항목이 있으면 해당 항목만 재요청합니다.
- 영어 제목은 `한글 제목 (English title) (date)` 형식으로 변환합니다.
- 상세 표의 bullet summary는 각 줄 앞에 `-`를 붙입니다.
- URL은 Word 문서 내 하이퍼링크로 삽입됩니다.

## 패키징

NSIS 설치 파일 생성:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_installer.ps1
```

필요 도구:

- Python runtime
- PyInstaller
- NSIS: `C:\Program Files (x86)\NSIS\makensis.exe`

생성 산출물:

```text
dist\PPCResearchHelper.exe
dist\PPCResearchHelperSetup.exe
```

설치 프로그램 정보:

- 이름: `PPC 조사연구 도우미`
- 게시자: `AI`
- 설치 경로: `%LOCALAPPDATA%\Programs\PPCResearchHelper`
- 설치 시 `ppc_report_template.docx`를 `%LOCALAPPDATA%\PPCResearchAutomation\resources`에 반영합니다.

## 삭제 정책

Uninstall 시 다음 문구를 표시합니다.

```text
사용자 데이터도 함께 삭제할까요?
실행 기록, API Key 등의 사용자 데이터를 모두 삭제합니다.
```

사용자가 `예`를 선택하면 아래 폴더를 삭제합니다.

```text
%LOCALAPPDATA%\PPCResearchAutomation
```
