from tkinterdnd2 import TkinterDnD
from ui.main_window import PDFUtilityApp

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFUtilityApp(root)
    root.mainloop()