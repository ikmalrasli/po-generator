import os
import subprocess
import sys
import shutil
from pathlib import Path

def create_nsis_script():
    """Create NSIS installer script"""
    
    current_dir = Path.cwd()
    dist_dir = current_dir / "dist" / "PO-Generator"
    
    nsis_script = f"""
; PO Generator Installer Script
; Written for NSIS 3.0+

!define APPNAME "PO Generator"
!define VERSION "1.0.0"
!define PUBLISHER "ikmalrasli"
!define DESCRIPTION "Purchase Order Generator with AI-powered PDF processing"
!define URL "https://github.com/ikmalrasli/po-generator"

; Basic configuration
Name "${{APPNAME}}"
OutFile "${{APPNAME}}-Setup-${{VERSION}}.exe"
InstallDir "$PROGRAMFILES64\\${{APPNAME}}"
InstallDirRegKey HKLM "Software\\${{APPNAME}}" "InstallPath"
RequestExecutionLevel admin

; Interface settings
!include "MUI2.nsh"
!define MUI_ABORTWARNING
!define MUI_ICON "${{APPNAME}}.ico"
!define MUI_UNICON "${{APPNAME}}.ico"

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
    File /r "{dist_dir}\\*.*"
    
    ; Create Start Menu shortcut
    CreateDirectory "$SMPROGRAMS\\${{APPNAME}}"
    CreateShortCut "$SMPROGRAMS\\${{APPNAME}}\\${{APPNAME}}.lnk" "$INSTDIR\\PO-Generator.exe" "" "$INSTDIR\\PO-Generator.exe" 0
    CreateShortCut "$SMPROGRAMS\\${{APPNAME}}\\Uninstall.lnk" "$INSTDIR\\uninstall.exe"
    
    ; Create desktop shortcut (optional)
    CreateShortCut "$DESKTOP\\${{APPNAME}}.lnk" "$INSTDIR\\PO-Generator.exe" "" "$INSTDIR\\PO-Generator.exe" 0
    
    ; Registry entries for Add/Remove Programs
    WriteRegStr HKLM "Software\\${{APPNAME}}" "InstallPath" "$INSTDIR"
    WriteRegStr HKLM "Software\\${{APPNAME}}" "Version" "${{VERSION}}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "DisplayName" "${{APPNAME}}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "UninstallString" "$INSTDIR\\uninstall.exe"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "DisplayVersion" "${{VERSION}}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "Publisher" "${{PUBLISHER}}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "URLInfoAbout" "${{URL}}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "DisplayIcon" "$INSTDIR\\PO-Generator.exe"
    WriteRegDWORD HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "NoModify" 1
    WriteRegDWORD HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "NoRepair" 1
    
    ; Create uninstaller
    WriteUninstaller "$INSTDIR\\uninstall.exe"
SectionEnd

; Uninstaller section
Section "Uninstall"
    Delete "$INSTDIR\\uninstall.exe"
    RMDir /r "$INSTDIR"
    
    ; Remove shortcuts
    Delete "$DESKTOP\\${{APPNAME}}.lnk"
    Delete "$SMPROGRAMS\\${{APPNAME}}\\${{APPNAME}}.lnk"
    Delete "$SMPROGRAMS\\${{APPNAME}}\\Uninstall.lnk"
    RMDir "$SMPROGRAMS\\${{APPNAME}}"
    
    ; Remove registry entries
    DeleteRegKey HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}"
    DeleteRegKey HKLM "Software\\${{APPNAME}}"
SectionEnd
"""
    
    # Write NSIS script to file
    with open("installer.nsi", "w", encoding="utf-8") as f:
        f.write(nsis_script)
    
    return "installer.nsi"

def create_license_file():
    """Create a basic LICENSE file if it doesn't exist"""
    license_content = """MIT License

Copyright (c) 2025 PO Generator

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE."""
    
    if not Path("LICENSE.txt").exists():
        with open("LICENSE.txt", "w", encoding="utf-8") as f:
            f.write(license_content)
        print("✅ Created LICENSE.txt")

def build_installer():
    """Build the NSIS installer"""
    
    print("🔨 Building NSIS installer...")
    
    # Check if dist directory exists
    dist_dir = Path("dist/PO-Generator")
    if not dist_dir.exists():
        print("❌ Error: dist/PO-Generator directory not found!")
        print("Please run 'python build.py' first to create the executable.")
        return False
    
    # Create license file
    create_license_file()
    
    # Create NSIS script
    nsis_script = create_nsis_script()
    print(f"✅ Created {nsis_script}")
    
    # Try to compile with NSIS
    try:
        # Try common NSIS installation paths
        nsis_paths = [
            r"C:\Program Files (x86)\NSIS\makensis.exe",
            r"C:\Program Files\NSIS\makensis.exe",
            r"C:\NSIS\makensis.exe"
        ]
        
        makensis_path = None
        for path in nsis_paths:
            if Path(path).exists():
                makensis_path = path
                break
        
        if makensis_path:
            print(f"📦 Compiling installer with NSIS...")
            result = subprocess.run([makensis_path, nsis_script], 
                                  capture_output=True, text=True)
            
            if result.returncode == 0:
                print("✅ Installer created successfully!")
                print(f"📁 Installer location: {Path.cwd() / 'PO-Generator-Setup-1.0.0.exe'}")
                return True
            else:
                print(f"❌ NSIS compilation failed: {result.stderr}")
                return False
        else:
            print("⚠️  NSIS not found in common locations.")
            print("Please install NSIS from https://nsis.sourceforge.io/Download")
            print(f"Then manually compile: {nsis_script}")
            return False
            
    except Exception as e:
        print(f"❌ Error building installer: {e}")
        return False

if __name__ == "__main__":
    success = build_installer()
    sys.exit(0 if success else 1)
