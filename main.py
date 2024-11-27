import tkinter as tk
from ui.app_ui import HelperApp

def main():
    root = tk.Tk()
    app = HelperApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()