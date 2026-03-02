import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinterdnd2 import DND_FILES
import os
from services.pdf_service import PDFService
from utils.validators import parse_page_ranges

class PDFUtilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Utility")
        self.root.geometry("760x460")
        self.root.resizable(False, False)

        self.pdf_files = []

        self.setup_styles()
        self.create_widgets()
        self.update_buttons()

    # ---------- UI ----------
    def setup_styles(self):
        style = ttk.Style(self.root)
        style.configure("TButton", padding=6, font=("Segoe UI", 10))

    def create_widgets(self):
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, width=80, height=15, font=("Segoe UI", 10))
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind("<<Drop>>", self.on_drop)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)

        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Add PDF", command=self.add_pdf).grid(row=0, column=0, padx=4)
        ttk.Button(button_frame, text="Remove PDF", command=self.remove_pdf).grid(row=0, column=1, padx=4)
        ttk.Button(button_frame, text="Merge PDFs", command=self.merge_pdfs).grid(row=0, column=2, padx=4)
        ttk.Button(button_frame, text="Rotate PDFs", command=self.rotate_pdfs).grid(row=0, column=3, padx=4)
        ttk.Button(button_frame, text="Delete Pages", command=self.delete_pages).grid(row=0, column=4, padx=4)
        ttk.Button(button_frame, text="Split PDF", command=self.split_pdf).grid(row=0, column=5, padx=4)
        ttk.Button(button_frame, text="Extract Pages", command=self.extract_pages_range).grid(row=0, column=6, padx=4)

        buttons = button_frame.winfo_children()
        self.remove_button = buttons[1]
        self.merge_button = buttons[2]
        self.rotate_button = buttons[3]
        self.delete_button = buttons[4]
        self.split_button = buttons[5]
        self.extract_button = buttons[6]

    # ---------- HELPERS ----------
    def update_buttons(self):
        state = tk.NORMAL if self.pdf_files else tk.DISABLED
        for btn in (self.remove_button, self.merge_button, self.rotate_button,
                    self.delete_button, self.split_button, self.extract_button):
            btn.config(state=state)

    def safe_read_pdf(self, path):
        return PDFService.safe_read_pdf(path)

    # ---------- DRAG & DROP ----------
    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for file in files:
            if file.lower().endswith(".pdf") and file not in self.pdf_files:
                self.pdf_files.append(file)
                self.listbox.insert(tk.END, os.path.basename(file))
        self.update_buttons()

    # ---------- ACTIONS ----------
    def add_pdf(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.listbox.insert(tk.END, os.path.basename(file))
        self.update_buttons()

    def remove_pdf(self):
        selected = list(self.listbox.curselection())
        if not selected or not messagebox.askyesno("Confirm", "Remove selected PDFs?"):
            return
        for index in reversed(selected):
            self.listbox.delete(index)
            self.pdf_files.pop(index)
        self.update_buttons()

    def merge_pdfs(self):
        PDFService.merge_pdfs(self.pdf_files)

    def rotate_pdfs(self):
        PDFService.rotate_pdfs(self.pdf_files, self.listbox)

    def delete_pages(self):
        PDFService.delete_pages(self.pdf_files, self.listbox)

    def split_pdf(self):
        PDFService.split_pdf(self.pdf_files, self.listbox)

    def extract_pages_range(self):
        PDFService.extract_pages_range(self.pdf_files, self.listbox)