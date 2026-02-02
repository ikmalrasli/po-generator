import PyInstaller.__main__
import sys
import os

def build_executable():
    """Build the PO Generator executable with optimized settings"""
    
    # Get the directory of this script
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    print("Building PO Generator executable...")
    print(f"Working directory: {current_dir}")
    
    # PyInstaller arguments for optimal performance
    args = [
        'main.py',
        '--onedir',                    # Directory mode for faster startup
        '--windowed',                   # Hide console window
        '--name=PO-Generator',
        '--clean',                      # Clean temporary files
        '--noconfirm',                  # Auto-confirm overwrite
        
        # Performance optimizations
        '--noupx',                      # Don't compress (faster startup)
        '--strip',                      # Remove debug symbols
        '--exclude-module=matplotlib', # Exclude unused heavy modules
        '--exclude-module=pandas',
        '--exclude-module=numpy',
        '--exclude-module=scipy',
        
        # Include necessary data files
        f'--add-data={os.path.join(current_dir, "templates")};templates',
        f'--add-data={os.path.join(current_dir, "config")};config',
        f'--add-data={os.path.join(current_dir, "PO Generator.ico")};.',
        
        # Hidden imports (prevents import scanning delays)
        '--hidden-import=tkinter',
        '--hidden-import=tkinter.ttk',
        '--hidden-import=tkinter.messagebox',
        '--hidden-import=tkinter.filedialog',
        '--hidden-import=tkcalendar',
        '--hidden-import=openpyxl',
        '--hidden-import=openpyxl.workbook',
        '--hidden-import=openpyxl.worksheet',
        '--hidden-import=google.genai',
        '--hidden-import=google.generativeai',
        '--hidden-import=num2words',
        '--hidden-import=json',
        '--hidden-import=pathlib',
        '--hidden-import=PIL',
        '--hidden-import=PIL.Image',
        
        # Runtime hook for faster startup
        '--runtime-hook=runtime_hook.py',
        
        
        '--icon=PO Generator.ico',
    ]
    
    try:
        PyInstaller.__main__.run(args)
        print("\n✅ Build completed successfully!")
        print(f"Executable location: {os.path.join(current_dir, 'dist', 'PO-Generator', 'PO-Generator.exe')}")
        print("\nNext steps:")
        print("1. Test the executable by running PO-Generator.exe")
        print("2. If it works, run 'python build_installer.py' to create the installer")
        
    except Exception as e:
        print(f"\n❌ Build failed: {e}")
        print("\nTroubleshooting:")
        print("1. Make sure PyInstaller is installed: pip install pyinstaller")
        print("2. Check that all required files exist")
        print("3. Try running without --windowed to see error messages")
        return False
    
    return True

if __name__ == "__main__":
    success = build_executable()
    sys.exit(0 if success else 1)
