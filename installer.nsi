Unicode true
Name "PPC 조사연구 도우미"
OutFile "dist\PPCResearchHelperSetup.exe"
InstallDir "$LOCALAPPDATA\Programs\PPCResearchHelper"
RequestExecutionLevel user
Icon "setup_icon.ico"

; 버전의 단일 원본은 updater.py 의 APP_VERSION 이다. build_installer.ps1 이 거기서 읽어
; /DAPP_VERSION=... 으로 넘긴다. 직접 makensis 를 돌릴 때를 위한 폴백만 아래에 둔다.
!ifndef APP_VERSION
  !define APP_VERSION "1.0.2"
!endif

!define APP_EXE "PPCResearchHelper.exe"
!define APP_NAME "PPC 조사연구 도우미"
!define COMPANY_NAME "PPCResearchHelper"
!define PUBLISHER_NAME "전북은행 AI혁신부"
!define APP_DATA_DIR "$LOCALAPPDATA\PPCResearchAutomation"

VIProductVersion "${APP_VERSION}.0"
VIAddVersionKey /LANG=1042 "ProductName" "${APP_NAME}"
VIAddVersionKey /LANG=1042 "CompanyName" "${PUBLISHER_NAME}"
VIAddVersionKey /LANG=1042 "FileDescription" "${APP_NAME} 설치 프로그램"
VIAddVersionKey /LANG=1042 "FileVersion" "${APP_VERSION}"
VIAddVersionKey /LANG=1042 "ProductVersion" "${APP_VERSION}"
VIAddVersionKey /LANG=1042 "LegalCopyright" "${PUBLISHER_NAME}"

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

Function .onInit
  nsExec::ExecToLog 'taskkill /F /T /IM "${APP_EXE}"'
FunctionEnd

Function un.onInit
  nsExec::ExecToLog 'taskkill /F /T /IM "${APP_EXE}"'
FunctionEnd

Section "Install"
  nsExec::ExecToLog 'taskkill /F /T /IM "${APP_EXE}"'
  Sleep 1000

  SetOutPath "$INSTDIR"
  SetOverwrite on
  File "dist\${APP_EXE}"
  File "PPC.ico"

  CreateDirectory "${APP_DATA_DIR}\resources"
  SetOutPath "${APP_DATA_DIR}\resources"
  SetOverwrite on
  File "ppc_report_template.docx"

  WriteUninstaller "$INSTDIR\Uninstall.exe"
  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\PPC.ico" 0
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\삭제.lnk" "$INSTDIR\Uninstall.exe"
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\PPC.ico" 0

  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "DisplayName" "${APP_NAME}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "DisplayVersion" "${APP_VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "Publisher" "${PUBLISHER_NAME}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "DisplayIcon" "$INSTDIR\PPC.ico"
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "NoRepair" 1
SectionEnd

Section "Uninstall"
  nsExec::ExecToLog 'taskkill /F /T /IM "${APP_EXE}"'
  Sleep 1000

  MessageBox MB_YESNO|MB_ICONQUESTION "사용자 데이터도 함께 삭제할까요?$\r$\n실행 기록, API Key 등의 사용자 데이터를 모두 삭제합니다." IDNO skipUserDataDelete
  RMDir /r "${APP_DATA_DIR}"
skipUserDataDelete:

  Delete "$DESKTOP\${APP_NAME}.lnk"
  Delete "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk"
  Delete "$SMPROGRAMS\${APP_NAME}\삭제.lnk"
  RMDir "$SMPROGRAMS\${APP_NAME}"

  Delete "$INSTDIR\${APP_EXE}"
  Delete "$INSTDIR\PPC.ico"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"

  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}"
SectionEnd
