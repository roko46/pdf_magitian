import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from pypdf import PdfReader, PdfWriter
from tkinterdnd2 import DND_FILES, TkinterDnD
import os


class PDFUtilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Utility")
        self.root.geometry("650x450")
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

        self.listbox = tk.Listbox(
            frame,
            selectmode=tk.MULTIPLE,
            width=70,
            height=15,
            font=("Segoe UI", 10)
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # DRAG & DROP ENABLE
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind("<<Drop>>", self.on_drop)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Add PDF", command=self.add_pdf).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Remove PDF", command=self.remove_pdf).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Merge PDFs", command=self.merge_pdfs).grid(row=0, column=2, padx=5)
        ttk.Button(button_frame, text="Rotate PDFs", command=self.rotate_pdfs).grid(row=0, column=3, padx=5)
        ttk.Button(button_frame, text="Delete Pages", command=self.delete_pages).grid(row=0, column=4, padx=5)

        self.merge_button = button_frame.winfo_children()[2]
        self.rotate_button = button_frame.winfo_children()[3]
        self.delete_button = button_frame.winfo_children()[4]
        self.remove_button = button_frame.winfo_children()[1]

    # ---------- HELPERS ----------
    def update_buttons(self):
        state = tk.NORMAL if self.pdf_files else tk.DISABLED
        for btn in (self.merge_button, self.rotate_button, self.delete_button, self.remove_button):
            btn.config(state=state)

    def safe_read_pdf(self, path):
        try:
            return PdfReader(path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open PDF:\n{path}\n\n{e}")
            return None

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
        if not selected:
            return
        for index in reversed(selected):
            self.listbox.delete(index)
            self.pdf_files.pop(index)
        self.update_buttons()

    def merge_pdfs(self):
        writer = PdfWriter()
        for file in self.pdf_files:
            reader = self.safe_read_pdf(file)
            if not reader:
                continue
            for page in reader.pages:
                writer.add_page(page)

        output = filedialog.asksaveasfilename(defaultextension=".pdf")
        if output:
            with open(output, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", "PDFs merged successfully!")

    def rotate_pdfs(self):
        selected = list(self.listbox.curselection())
        if not selected:
            return

        degrees = simpledialog.askinteger("Rotate", "Enter 90, 180 or 270:")
        if degrees not in (90, 180, 270):
            return

        for index in selected:
            reader = self.safe_read_pdf(self.pdf_files[index])
            if not reader:
                continue
            writer = PdfWriter()
            for page in reader.pages:
                page.rotate_clockwise(degrees)
                writer.add_page(page)

            output = filedialog.asksaveasfilename(defaultextension=".pdf")
            if output:
                with open(output, "wb") as f:
                    writer.write(f)

    def delete_pages(self):
        selected = list(self.listbox.curselection())
        if not selected:
            return

        pages_input = simpledialog.askstring("Delete Pages", "e.g. 1,3,5-7")
        if not pages_input:
            return

        for index in selected:
            reader = self.safe_read_pdf(self.pdf_files[index])
            if not reader:
                continue

            pages = self.parse_pages(pages_input, len(reader.pages))
            writer = PdfWriter()
            for i, page in enumerate(reader.pages):
                if i not in pages:
                    writer.add_page(page)

            output = filedialog.asksaveasfilename(defaultextension=".pdf")
            if output:
                with open(output, "wb") as f:
                    writer.write(f)

    def parse_pages(self, text, max_pages):
        pages = set()
        try:
            for part in text.split(","):
                part = part.strip()
                if "-" in part:
                    start, end = map(int, part.split("-"))
                    pages.update(range(start - 1, min(end, max_pages)))
                else:
                    pages.add(int(part) - 1)
        except ValueError:
            messagebox.showerror("Error", "Invalid page format")
        return pages


# ---------- RUN ----------
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFUtilityApp(root)
    root.mainloop()