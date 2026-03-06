import tkinter as tk
from tkinter import filedialog, ttk
from tkinterdnd2 import DND_FILES
import os
from services.pdf_service import PDFService

class PDFUtilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDFMagitiana")
        self.root.geometry("820x480")
        self.root.resizable(False, False)

        self.pdf_files = []

        self.setup_styles()
        self.create_widgets()
        self.update_buttons()

    def setup_styles(self):
        style = ttk.Style(self.root)
        style.configure("TButton", padding=6, font=("Segoe UI", 10))

    def create_widgets(self):
        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, width=90, height=15)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind("<<Drop>>", self.on_drop)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)

        buttons = [
            ("Add PDF", self.add_pdf),
            ("Remove PDF", self.remove_pdf),
            ("Merge PDFs", self.merge_pdfs),
            ("Rotate PDFs", self.rotate_pdfs),
            ("Delete Pages", self.delete_pages),
            ("Split PDF", self.split_pdf),
            ("Extract Pages", self.extract_pages_range),
            ("Reorder Pages", self.reorder_pages),
        ]

        self.action_buttons = []
        for i, (text, cmd) in enumerate(buttons):
            btn = ttk.Button(button_frame, text=text, command=cmd)
            btn.grid(row=i//4, column=i%4, padx=3, pady=3)
            self.action_buttons.append(btn)

    def update_buttons(self):
        any_pdf = bool(self.pdf_files)
        self.action_buttons[0].config(state=tk.NORMAL)  # Add PDF
        self.action_buttons[1].config(state=any_pdf)   # Remove PDF
        self.action_buttons[2].config(state=tk.NORMAL if len(self.pdf_files) > 1 else tk.DISABLED)  # Merge
        for btn in self.action_buttons[3:]:
            btn.config(state=any_pdf)  # Rotate, Delete, Split, Extract, Reorder

    def on_drop(self, event):
        files = [f.strip("{}") for f in self.root.tk.splitlist(event.data)]
        for f in files:
            if f.lower().endswith(".pdf") and f not in self.pdf_files:
                self.pdf_files.append(f)
                self.listbox.insert(tk.END, os.path.basename(f))
        self.update_buttons()

    def add_pdf(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for f in files:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
                self.listbox.insert(tk.END, os.path.basename(f))
        self.update_buttons()

    def remove_pdf(self):
        selected = list(self.listbox.curselection())
        if not selected:
            return
        for i in reversed(selected):
            self.listbox.delete(i)
            self.pdf_files.pop(i)
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

    def reorder_pages(self):
        PDFService.open_reorder_window(self.root, self.pdf_files, self.listbox)