import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from pypdf import PdfReader, PdfWriter
from tkinterdnd2 import DND_FILES, TkinterDnD
import os


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

        self.listbox = tk.Listbox(
            frame,
            selectmode=tk.MULTIPLE,
            width=80,
            height=15,
            font=("Segoe UI", 10)
        )
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
        for btn in (
            self.remove_button,
            self.merge_button,
            self.rotate_button,
            self.delete_button,
            self.split_button,
            self.extract_button,
        ):
            btn.config(state=state)

    def safe_read_pdf(self, path):
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                reader.decrypt("")
            return reader
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
        if not messagebox.askyesno("Confirm", "Remove selected PDFs from list?"):
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
                return
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

        folder = filedialog.askdirectory(title="Select output folder")
        if not folder:
            return

        for index in selected:
            reader = self.safe_read_pdf(self.pdf_files[index])
            if not reader:
                continue

            writer = PdfWriter()
            for page in reader.pages:
                page.rotate_clockwise(degrees)
                writer.add_page(page)

            name = os.path.basename(self.pdf_files[index])
            out = os.path.join(folder, f"rotated_{name}")
            with open(out, "wb") as f:
                writer.write(f)

        messagebox.showinfo("Success", "Rotation completed!")

    def delete_pages(self):
        selected = list(self.listbox.curselection())
        if not selected:
            return

        pages_input = simpledialog.askstring("Delete Pages", "e.g. 1,3,5-7")
        if not pages_input:
            return

        folder = filedialog.askdirectory(title="Select output folder")
        if not folder:
            return

        for index in selected:
            reader = self.safe_read_pdf(self.pdf_files[index])
            if not reader:
                continue

            delete_pages = self.parse_pages(pages_input, len(reader.pages))
            writer = PdfWriter()

            for i, page in enumerate(reader.pages):
                if i not in delete_pages:
                    writer.add_page(page)

            name = os.path.basename(self.pdf_files[index])
            out = os.path.join(folder, f"pages_deleted_{name}")
            with open(out, "wb") as f:
                writer.write(f)

        messagebox.showinfo("Success", "Pages deleted successfully!")

    # ---------- SPLIT ----------
    def split_pdf(self):
        selected = list(self.listbox.curselection())
        if len(selected) != 1:
            messagebox.showwarning("Warning", "Select exactly ONE PDF.")
            return

        reader = self.safe_read_pdf(self.pdf_files[selected[0]])
        if not reader:
            return

        total = len(reader.pages)
        split_page = simpledialog.askinteger(
            "Split PDF",
            f"Split AFTER page (1–{total - 1}):"
        )

        if not split_page or split_page < 1 or split_page >= total:
            return

        folder = filedialog.askdirectory()
        if not folder:
            return

        base = os.path.splitext(os.path.basename(self.pdf_files[selected[0]]))[0]

        w1, w2 = PdfWriter(), PdfWriter()
        for i, page in enumerate(reader.pages):
            (w1 if i < split_page else w2).add_page(page)

        with open(os.path.join(folder, f"{base}_part1.pdf"), "wb") as f:
            w1.write(f)
        with open(os.path.join(folder, f"{base}_part2.pdf"), "wb") as f:
            w2.write(f)

        messagebox.showinfo("Success", "PDF split successfully!")

    # ---------- EXTRACT PAGES (FROM–TO) ----------
    def extract_pages_range(self):
        selected = list(self.listbox.curselection())
        if len(selected) != 1:
            messagebox.showwarning("Warning", "Select exactly ONE PDF.")
            return

        reader = self.safe_read_pdf(self.pdf_files[selected[0]])
        if not reader:
            return

        total = len(reader.pages)

        start = simpledialog.askinteger("Extract Pages", f"From page (1–{total}):")
        if start is None:
            return

        end = simpledialog.askinteger("Extract Pages", f"To page ({start}–{total}):")
        if end is None:
            return

        if start < 1 or end > total or start > end:
            messagebox.showerror("Error", "Invalid page range.")
            return

        output = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile=f"extracted_{start}-{end}.pdf"
        )
        if not output:
            return

        writer = PdfWriter()
        for i in range(start - 1, end):
            writer.add_page(reader.pages[i])

        with open(output, "wb") as f:
            writer.write(f)

        messagebox.showinfo("Success", "Pages extracted successfully!")

    # ---------- UTIL ----------
    def parse_pages(self, text, max_pages):
        pages = set()
        try:
            for part in text.split(","):
                part = part.strip()
                if "-" in part:
                    s, e = map(int, part.split("-"))
                    pages.update(range(s - 1, min(e, max_pages)))
                else:
                    pages.add(int(part) - 1)
        except ValueError:
            messagebox.showerror("Error", "Invalid page format")
        return {p for p in pages if 0 <= p < max_pages}


# ---------- RUN ----------
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFUtilityApp(root)
    root.mainloop()