; VVC Installer - Simple NSIS
!include "MUI2.nsh"
!include "x64.nsh"

; Basic settings
!ifndef OUTDIR
  !define OUTDIR "."
!endif

Name "VVC"
OutFile "${OUTDIR}\VVC-Installer.exe"
InstallDir "$APPDATA\VVC"
RequestExecutionLevel user

; MUI Settings
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_LANGUAGE "English"

Var LaunchAfter
Var InstallArgs
Var PsLog
Var ExitCode

; LaunchAfter defaults to 1, updated by .onSelChange
Function .onInit
  SetAutoClose false
  StrCpy $LaunchAfter "1"
FunctionEnd

; Installer sections
Section "!Install VVC" SecMain
  SectionIn RO  ; required, can't uncheck
  SetOutPath "$INSTDIR"
  
  ${If} ${RunningX64}
    SetRegView 64
  ${EndIf}
  
  File "payload.zip"
  File "installer\install.ps1"
  File "installer\uninstall.ps1"
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  StrCpy $PsLog "$TEMP\VVC-installer-console.log"
  StrCpy $InstallArgs "-InstallDir $\"$INSTDIR$\" -LogFilePath $\"$PsLog$\""

  ${If} $LaunchAfter == "1"
    StrCpy $InstallArgs "$InstallArgs -LaunchAfterInstall"
  ${EndIf}

  DetailPrint "Installer console log: $PsLog"
  ExecWait '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\install.ps1" $InstallArgs' $ExitCode
  DetailPrint "PowerShell exit code: $ExitCode"
  ${If} $ExitCode != 0
    MessageBox MB_ICONSTOP|MB_OK "Installation failed (exit code $ExitCode). Log: $PsLog"
    Abort
  ${EndIf}
SectionEnd

Section "Launch VVC after install" SecLaunch
  ; selected by default via SectionIn
  SectionIn 1
SectionEnd

; Track checkbox state
Function .onSelChange
  SectionGetFlags ${SecLaunch} $0
  IntOp $0 $0 & ${SF_SELECTED}
  ${If} $0 <> 0
    StrCpy $LaunchAfter "1"
  ${Else}
    StrCpy $LaunchAfter "0"
  ${EndIf}
FunctionEnd

; Section descriptions shown on hover
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecMain} "Installs VVC to your AppData folder."
  !insertmacro MUI_DESCRIPTION_TEXT ${SecLaunch} "Starts VVC immediately after installation."
!insertmacro MUI_FUNCTION_DESCRIPTION_END

Section "Uninstall"
  ExecWait 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\uninstall.ps1" -InstallDir "$INSTDIR"'
  RMDir /r "$INSTDIR"
SectionEnd
