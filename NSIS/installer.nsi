; NSIS 3.x
SetCompressor /SOLID lzma
ShowInstDetails show
!define APPNAME "SendGuard"
!define PROGID  "SendGuard"

RequestExecutionLevel user
InstallDir "$LocalAppData\${APPNAME}"
OutFile "SendGuard-UserSetup.exe"

; Set a custom title for the installer
Caption "SendGuard Installer"

; Set a custom title for the uninstaller
UninstallCaption "SendGuard Uninstaller"

; Add a license page for terms and conditions
LicenseText "Please read and accept the terms and conditions to proceed." "I Agree"
LicenseData "license.txt"

; Add a welcome page and confirmation page
PageEx license
  LicenseForceSelection checkbox
PageExEnd

; Currently, only per-user installation is supported
; Future: add choice between per-user and per-machine installation
Section "Install (Per-User)"
  ; Check if Outlook is running
  FindWindow $0 "rctrl_renwnd32" ; Class name for Outlook main window
  StrCmp $0 0 +4
  IfSilent +2
  MessageBox MB_ICONEXCLAMATION|MB_OK "Please close Microsoft Outlook before proceeding with the installation."
  Quit

  SetOutPath "$LocalAppData\${APPNAME}"
  File /r "..\bin\Release\*"     ; includes DLLs, .vsto, manifests

  ; Per-user policy (create if missing)
  SetOutPath "$AppData\SendGuard"

  ; Ensure target directory exists
  IfFileExists "$AppData\SendGuard\*" +3
    CreateDirectory "$AppData\SendGuard"
    DetailPrint "Created $AppData\SendGuard directory."

  ;
  ; silent install process
  ;
  IfSilent +1 +4
    IfFileExists "$EXEDIR\policy.user.json" +1 +2
      Goto WriteCustomPolicy
    Goto WriteDefaultPolicy

  ;
  ; non-silent interactive install process
  ;
  ; if policy.user.json exists in installer directory, offer to use it overwriting any existing policy.json
  ; otherwise offer to write default sample policy.json overwriting any existing policy.json
  ;
  IfFileExists "$EXEDIR\policy.user.json" 0 +5
    IfFileExists "$AppData\SendGuard\policy.json" 0 +3
      MessageBox MB_YESNO|MB_ICONQUESTION "$AppData\SendGuard\policy.json already exists. Do you want to overwrite this with custom policy.user.json? You will lose your existing policies if you choose yes." IDYES +2
      Goto PolicyDone
      Goto WriteCustomPolicy
  IfFileExists "$AppData\SendGuard\policy.json" 0 +3
    MessageBox MB_YESNO|MB_ICONQUESTION "$AppData\SendGuard\policy.json already exists. Do you want to overwrite this with sample policy from installer? You will lose your existing policies if you choose yes." IDYES +2
    Goto PolicyDone
    Goto WriteDefaultPolicy

WriteCustomPolicy:
  IfFileExists "$EXEDIR\policy.user.json" 0 +3
    CopyFiles /SILENT "$EXEDIR\policy.user.json" "$AppData\SendGuard\policy.json"
    DetailPrint "Custom policy.user.json to $AppData\SendGuard\policy.json"
  Goto PolicyDone

WriteDefaultPolicy:
  File "policy.json"
  DetailPrint "Sample policy.json written to $AppData\SendGuard\policy.json"
  Goto PolicyDone

PolicyDone:
  ; COM/VSTO registration under HKCU
  StrCpy $0 "file:///$LocalAppData/${APPNAME}/SendGuard.vsto|vstolocal"
  WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\${PROGID}" "FriendlyName" "${APPNAME}"
  WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\${PROGID}" "Description" "Blocks non-GPG attachments to protected domains"
  WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\${PROGID}" "Manifest" "$0"
  WriteRegDWORD HKCU "Software\Microsoft\Office\Outlook\Addins\${PROGID}" "LoadBehavior" 3

  ; Write uninstaller
  WriteUninstaller "$LocalAppData\${APPNAME}\uninstall.exe"

  ; Add uninstall entry to Add/Remove Programs
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$LocalAppData\${APPNAME}\uninstall.exe"
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoRepair" 1

  DetailPrint "${APPNAME} installed successfully."
  DetailPrint "You can uninstall it via Control Panel or Settings."
  DetailPrint "Please restart Outlook to enable the SendGuard add-in."
  DetailPrint "You can customise rules and policies file at:"
  DetailPrint "  $AppData\SendGuard\policy.json"
SectionEnd

Section "Uninstall"
  ; Remove registry entries
  DeleteRegKey HKCU "Software\Microsoft\Office\Outlook\Addins\${PROGID}"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"

  ; Remove application files
  RMDir /r "$LocalAppData\${APPNAME}"

  ; Keep $AppData\GpgSendGuard\policy.json by default
SectionEnd

; Add the installation progress page
Page instfiles
