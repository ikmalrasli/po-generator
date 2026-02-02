# Build Instructions - PO Generator

This guide will help you create a distributable Windows executable with professional installer for your PO Generator application.

## Prerequisites

### Required Software
1. **Python 3.8+** (already installed if you're running the app)
2. **PyInstaller** - For creating the executable
3. **NSIS** (Nullsoft Scriptable Install System) - For creating the installer

### Installation Commands
```bash
# Install PyInstaller
pip install pyinstaller

# Download and install NSIS from: https://nsis.sourceforge.io/Download
# Choose the "NSIS 3.08" or later version
```

## Step-by-Step Build Process

### Step 1: Prepare Your Project
Ensure your project is ready:
- ✅ All dependencies in `requirements.txt`
- ✅ Excel template in `data/templates/po_template.xlsx`
- ✅ Application runs correctly with `python main.py`

### Step 2: Build the Executable
```bash
# Build the optimized executable
python build.py
```

**What this does:**
- Creates a directory-based executable (faster startup than single file)
- Optimizes for performance with pre-imports
- Includes all necessary data files
- Excludes unused modules to reduce size

**Expected output:**
```
Building PO Generator executable...
Working directory: C:\cc\PO-generator
✅ Build completed successfully!
Executable location: C:\cc\PO-generator\dist\PO-Generator\PO-Generator.exe
```

### Step 3: Test the Executable
1. Navigate to `dist/PO-Generator/`
2. Run `PO-Generator.exe`
3. Verify all features work correctly:
   - GUI loads properly
   - PDF processing works
   - Excel generation works
   - Settings are saved to AppData

### Step 4: Create the Installer
```bash
# Build the professional installer
python build_installer.py
```

**What this does:**
- Creates NSIS installer script
- Generates `PO-Generator-Setup-1.0.0.exe`
- Includes uninstaller
- Creates Start Menu and Desktop shortcuts

### Step 5: Test the Installer
1. Run `PO-Generator-Setup-1.0.0.exe`
2. Verify installation:
   - Installs to `C:\Program Files\PO-Generator\`
   - Creates Start Menu shortcuts
   - Creates Desktop shortcut
   - Appears in Add/Remove Programs
3. Test uninstallation works correctly

## File Structure After Build

### Executable Directory
```
dist/PO-Generator/
├── PO-Generator.exe          # Main executable
├── _internal/                # Python dependencies
│   ├── tkinter.dll
│   ├── openpyxl/
│   ├── google/
│   └── ... (dependency files)
├── data/                     # Excel templates
│   └── templates/
│       └── po_template.xlsx
└── config/                   # Configuration files
    └── settings.py
```

### User Data Storage
User data is stored in Windows AppData (not in program directory):
```
%APPDATA%/PO-Generator/
├── temp/                     # Temporary files
├── jsons/                    # Extracted PDF data
└── config/
    └── user_settings.json    # User preferences
```

## Distribution

### What to Distribute
- **Single file**: `PO-Generator-Setup-1.0.0.exe` (≈50-100MB)
- This is all users need to install your application

### Installation Experience for Users
1. Download `PO-Generator-Setup-1.0.0.exe`
2. Run installer (Windows may show security warning - this is normal)
3. Follow installation wizard
4. Launch from Start Menu or Desktop

## Performance Optimizations

### Startup Speed
- **One-directory mode**: ~1-2 seconds startup
- **Runtime hook**: Pre-imports common modules
- **Excluded modules**: Removes heavy unused dependencies
- **No compression**: Faster startup at cost of larger file size

### Memory Usage
- ~50-100MB RAM usage (typical for Tkinter apps)
- Dependencies loaded on-demand where possible

## Troubleshooting

### Common Issues

#### Build Fails
```bash
# Check PyInstaller installation
pip install --upgrade pyinstaller

# Try building without optimizations
pyinstaller --onedir --windowed main.py
```

#### Executable Won't Run
1. Check antivirus isn't blocking it
2. Run as administrator
3. Check Windows Event Viewer for errors
4. Try building with `--console` to see error messages

#### Installer Issues
1. Ensure NSIS is properly installed
2. Check `installer.nsi` script for errors
3. Manually compile with NSIS if automated build fails

#### Missing Files
1. Verify `data/templates/po_template.xlsx` exists
2. Check all required files are included in build script
3. Test with `--debug all` flag for detailed logging

## Advanced Options

### Custom Icon
Add an icon file to your project:
1. Create `icon.ico` (256x256 pixels)
2. Uncomment `--icon=icon.ico` in `build.py`
3. Rebuild executable

### Version Information
Update version numbers in:
- `setup.py` (line 5)
- `build_installer.py` (line 7)
- `installer.nsi` (generated automatically)

### Silent Installation
For enterprise deployment:
```bash
PO-Generator-Setup-1.0.0.exe /S
```

## Support

For build issues:
1. Check this document first
2. Review PyInstaller documentation: https://pyinstaller.readthedocs.io/
3. Review NSIS documentation: https://nsis.sourceforge.io/Main_Page
4. Create an issue on your project repository

## Security Notes

- The executable may trigger antivirus warnings (false positives)
- Code signing certificate eliminates warnings (costs money)
- Always scan downloads before distribution
- Consider code signing for enterprise distribution
