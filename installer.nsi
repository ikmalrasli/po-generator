
; PO Generator Installer Script
; Written for NSIS 3.0+

!define APPNAME "PO Generator"
!define VERSION "1.0.0"
!define PUBLISHER "ikmalrasli"
!define DESCRIPTION "Purchase Order Generator with AI-powered PDF processing"
!define URL "https://github.com/ikmalrasli/po-generator"

; Basic configuration
Name "${APPNAME}"
OutFile "${APPNAME}-Setup-${VERSION}.exe"
InstallDir "$PROGRAMFILES64\${APPNAME}"
InstallDirRegKey HKLM "Software\${APPNAME}" "InstallPath"
RequestExecutionLevel admin

; Interface settings
!include "MUI2.nsh"
!define MUI_ABORTWARNING
!define MUI_ICON "${APPNAME}.ico"
!define MUI_UNICON "${APPNAME}.ico"

; Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

; Languages
!insertmacro MUI_LANGUAGE "English"

; Installer sections
Section "MainSection" SEC01
    SetOutPath "$INSTDIR"
    
    ; Copy all files from the dist directory
    File /r "C:\cc\PO-generator\dist\PO-Generator\*.*"
    
    ; Create Start Menu shortcut
    CreateDirectory "$SMPROGRAMS\${APPNAME}"
    CreateShortCut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\PO-Generator.exe" "" "$INSTDIR\PO-Generator.exe" 0
    CreateShortCut "$SMPROGRAMS\${APPNAME}\Uninstall.lnk" "$INSTDIR\uninstall.exe"
    
    ; Create desktop shortcut (optional)
    CreateShortCut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\PO-Generator.exe" "" "$INSTDIR\PO-Generator.exe" 0
    
    ; Registry entries for Add/Remove Programs
    WriteRegStr HKLM "Software\${APPNAME}" "InstallPath" "$INSTDIR"
    WriteRegStr HKLM "Software\${APPNAME}" "Version" "${VERSION}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$INSTDIR\uninstall.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayVersion" "${VERSION}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "Publisher" "${PUBLISHER}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "URLInfoAbout" "${URL}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayIcon" "$INSTDIR\PO-Generator.exe"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoRepair" 1
    
    ; Create uninstaller
    WriteUninstaller "$INSTDIR\uninstall.exe"
SectionEnd

; Uninstaller section
Section "Uninstall"
    Delete "$INSTDIR\uninstall.exe"
    RMDir /r "$INSTDIR"
    
    ; Remove shortcuts
    Delete "$DESKTOP\${APPNAME}.lnk"
    Delete "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk"
    Delete "$SMPROGRAMS\${APPNAME}\Uninstall.lnk"
    RMDir "$SMPROGRAMS\${APPNAME}"
    
    ; Remove registry entries
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
    DeleteRegKey HKLM "Software\${APPNAME}"
SectionEnd
