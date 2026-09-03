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
- 사이트별 수집 건수: 10건 (10/20/30/50/100 또는 `모두`)
  - `모두`는 지정한 기간에 해당하는 기사를 목록 끝까지 가져옵니다.
    사이트당 목록 40페이지가 안전 상한이며, 여기서 끊긴 사이트는 완료 메시지에 표시됩니다.
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
- 버전: `updater.py` 의 `APP_VERSION` 이 단일 원본입니다. `build_installer.ps1` 이 이 값을 읽어
  NSIS 로 넘기므로, 버전을 올릴 때는 `updater.py` 만 고치면 됩니다.
- 게시자: `전북은행 AI혁신부`
- 설치 경로: `%LOCALAPPDATA%\Programs\PPCResearchHelper`
- 설치 시 `ppc_report_template.docx`를 `%LOCALAPPDATA%\PPCResearchAutomation\resources`에 반영합니다.

## 원격 업데이트

앱이 켜질 때 GitHub 릴리스의 `latest.json` 을 읽어 새 버전이 있으면 상단에 배너를 띄웁니다.
「지금 업데이트」를 누르면 팩(zip)을 내려받아 **SHA-256 을 검증**하고, 앱이 스스로 종료 →
`PPCResearchHelper.exe` 교체 → 재시작합니다. 사용자가 설치 파일을 따로 받을 필요가 없습니다.

- 배포처: <https://github.com/lunanalyze/JBPPCResearchHelper> (**공개** 저장소여야 합니다 —
  앱에 토큰을 심지 않으려면 인증 없이 읽혀야 합니다)
- 교체는 `updater/apply.ps1` 이 **별도 프로세스**로 합니다. 실행 중인 exe 는 Windows 가 잠그고
  있어 자기 자신을 못 바꾸기 때문입니다.
- 실패하면 `PPCResearchHelper.exe.bak` 으로 **자동 복구**하고 이전 버전을 다시 띄웁니다.
- 수집·보고서 생성이 진행 중이면 업데이트를 **거부**합니다(HTTP 409).
- 개발 실행(`python app.py`)에서는 동작하지 않습니다 — 소스 트리를 덮어쓰는 사고를 막기 위해
  설치본(`Uninstall.exe` 가 옆에 있는 exe)일 때만 적용됩니다.

환경변수로 끄거나 배포처를 바꿀 수 있습니다.

| 변수 | 뜻 |
|---|---|
| `PPCRH_UPDATE_ENABLED=0` | 자동 업데이트를 끕니다 |
| `PPCRH_UPDATE_FEED=<URL>` | `latest.json` 주소를 바꿉니다(사내 파일서버 등) |
| `PPCRH_NO_BROWSER=1` | 기동 시 브라우저를 열지 않습니다(업데이터가 재시작할 때 씁니다) |

### 릴리스 절차

```powershell
# 1) updater.py 의 APP_VERSION 을 올린다
# 2) exe + Setup.exe
powershell -ExecutionPolicy Bypass -File .\build_installer.ps1
# 3) 업데이트 팩 + latest.json
powershell -ExecutionPolicy Bypass -File .\build_release.ps1 -Notes "무엇이 바뀌었는지 한 줄"
```

그 다음 GitHub 에 **`v<버전>` 태그로 릴리스**를 만들고 아래 셋을 **에셋으로** 올립니다.

- `latest.json`
- `PPCResearchHelper-Update-<버전>.zip`
- `PPCResearchHelperSetup.exe` (신규 설치용)

`latest.json` 이 **최신 릴리스의 에셋**으로 붙어 있어야 `/releases/latest/download/latest.json`
이 그것을 가리킵니다.

> ⚠ 이미 설치된 1.0.1 에는 업데이트 기능이 없습니다. 1.0.2 는 `PPCResearchHelperSetup.exe` 로
> **한 번은 직접 설치**해야 하고, 그 다음부터 앱 안에서 업데이트됩니다.

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
