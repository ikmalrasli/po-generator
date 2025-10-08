import tkinter as tk
from gui.app import POGUI

def main():
    # Create necessary directories
    from config.settings import TEMP_DIR, JSONS_DIR
    TEMP_DIR.mkdir(exist_ok=True)
    JSONS_DIR.mkdir(exist_ok=True)
    
    root = tk.Tk()
    app = POGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()