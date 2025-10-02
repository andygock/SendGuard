; NSIS 3.x
SetCompressor /SOLID lzma
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

Page custom ConfirmInstallPage

Function ConfirmInstallPage
  MessageBox MB_YESNO|MB_ICONQUESTION "Are you sure you want to install ${APPNAME}?" IDYES +2
  Quit
FunctionEnd

Section "Install (Per-User)"
  ; Check if Outlook is running
  FindWindow $0 "rctrl_renwnd32" ; Class name for Outlook main window
  StrCmp $0 0 +3
  MessageBox MB_ICONEXCLAMATION|MB_OK "Please close Microsoft Outlook before proceeding with the installation."
  Quit

  SetOutPath "$LocalAppData\${APPNAME}"
  File /r "..\bin\Release\*"     ; includes DLLs, .vsto, manifests

  ; Per-user policy (create if missing)
  SetOutPath "$AppData\SendGuard"
  ; Ensure target directory exists
  IfFileExists "$AppData\SendGuard\*" +2 0
    CreateDirectory "$AppData\SendGuard"
  DetailPrint "Ensured $AppData\SendGuard directory exists."

  ; Only copy or write policy.json if it does not exist
  IfFileExists "$AppData\SendGuard\policy.json" 0 +5
    DetailPrint "policy.json already exists, not overwritten."
    Goto PolicyDone
  IfFileExists "$EXEDIR\policy.user.json" 0 +3
    CopyFiles /SILENT "$EXEDIR\policy.user.json" "$AppData\SendGuard\policy.json"
    Goto PolicyDone
  File "policy.json"
  DetailPrint "Default policy.json written successfully."
PolicyDone:

  ; COM/VSTO registration under HKCU
  StrCpy $0 "file:///$LocalAppData/${APPNAME}/SendGuard.vsto|vstolocal"
  WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\${PROGID}" "FriendlyName" "${APPNAME}"
  WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\${PROGID}" "Description" "Blocks non-GPG attachments to protected domains"
  WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\${PROGID}" "Manifest" "$0"
  WriteRegDWORD HKCU "Software\Microsoft\Office\Outlook\Addins\${PROGID}" "LoadBehavior" 3

  ; Optional: import publisher cert into CurrentUser store (no admin needed)
  ; File "publisher.cer"
  ; ExecWait 'certutil -user -addstore "TrustedPublisher" "$LocalAppData\${APPNAME}\publisher.cer"'

  ; Notify user to open Outlook after installation
  MessageBox MB_ICONINFORMATION|MB_OK "Installation complete. Please open Microsoft Outlook to continue the setup and use the SendGuard add-in."

  ; Write uninstaller
  WriteUninstaller "$LocalAppData\${APPNAME}\uninstall.exe"

  ; Add uninstall entry to Add/Remove Programs
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$LocalAppData\${APPNAME}\uninstall.exe"
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoRepair" 1
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
