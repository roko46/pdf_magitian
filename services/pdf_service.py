from pypdf import PdfReader, PdfWriter
from tkinter import filedialog, messagebox, simpledialog, ttk
import tkinter as tk
import os
from utils.validators import parse_page_ranges

class PDFService:

    @staticmethod
    def safe_read_pdf(path):
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                reader.decrypt("")
            return reader
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open PDF:\n{path}\n\n{e}")
            return None

    # ------------------ MERGE ------------------
    @staticmethod
    def merge_pdfs(pdf_files):
        if not pdf_files:
            messagebox.showwarning("Warning", "No PDFs selected!")
            return

        writer = PdfWriter()
        for file in pdf_files:
            reader = PDFService.safe_read_pdf(file)
            if not reader:
                return
            for page in reader.pages:
                writer.add_page(page)

        output = filedialog.asksaveasfilename(defaultextension=".pdf")
        if output:
            with open(output, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", "PDFs merged successfully!")

    # ------------------ ROTATE ------------------
    @staticmethod
    def rotate_pdfs(pdf_files, listbox):
        selected = list(listbox.curselection())
        if not selected:
            return

        degrees = simpledialog.askinteger("Rotate", "Enter 90, 180 or 270:")
        if degrees not in (90, 180, 270):
            return

        folder = filedialog.askdirectory(title="Select output folder")
        if not folder:
            return

        for index in selected:
            reader = PDFService.safe_read_pdf(pdf_files[index])
            if not reader:
                continue
            writer = PdfWriter()
            for page in reader.pages:
                page.rotate_clockwise(degrees)
                writer.add_page(page)
            name = os.path.basename(pdf_files[index])
            out = os.path.join(folder, f"rotated_{name}")
            with open(out, "wb") as f:
                writer.write(f)
        messagebox.showinfo("Success", "Rotation completed!")

    # ------------------ DELETE PAGES ------------------
    @staticmethod
    def delete_pages(pdf_files, listbox):
        selected = list(listbox.curselection())
        if not selected:
            return

        pages_input = simpledialog.askstring("Delete Pages", "e.g. 1,3,5-7")
        if not pages_input:
            return

        folder = filedialog.askdirectory(title="Select output folder")
        if not folder:
            return

        for index in selected:
            reader = PDFService.safe_read_pdf(pdf_files[index])
            if not reader:
                continue
            delete_pages = parse_page_ranges(pages_input, len(reader.pages))
            writer = PdfWriter()
            for i, page in enumerate(reader.pages):
                if i not in delete_pages:
                    writer.add_page(page)
            name = os.path.basename(pdf_files[index])
            out = os.path.join(folder, f"pages_deleted_{name}")
            with open(out, "wb") as f:
                writer.write(f)
        messagebox.showinfo("Success", "Pages deleted successfully!")

    # ------------------ SPLIT PDF ------------------
    @staticmethod
    def split_pdf(pdf_files, listbox):
        selected = list(listbox.curselection())
        if len(selected) != 1:
            messagebox.showwarning("Warning", "Select exactly ONE PDF.")
            return

        reader = PDFService.safe_read_pdf(pdf_files[selected[0]])
        if not reader:
            return

        total = len(reader.pages)
        split_page = simpledialog.askinteger("Split PDF", f"Split AFTER page (1–{total - 1}):")
        if not split_page or split_page < 1 or split_page >= total:
            return

        folder = filedialog.askdirectory()
        if not folder:
            return

        base = os.path.splitext(os.path.basename(pdf_files[selected[0]]))[0]
        w1, w2 = PdfWriter(), PdfWriter()
        for i, page in enumerate(reader.pages):
            (w1 if i < split_page else w2).add_page(page)
        with open(os.path.join(folder, f"{base}_part1.pdf"), "wb") as f:
            w1.write(f)
        with open(os.path.join(folder, f"{base}_part2.pdf"), "wb") as f:
            w2.write(f)
        messagebox.showinfo("Success", "PDF split successfully!")

    # ------------------ EXTRACT PAGES ------------------
    @staticmethod
    def extract_pages_range(pdf_files, listbox):
        selected = list(listbox.curselection())
        if len(selected) != 1:
            messagebox.showwarning("Warning", "Select exactly ONE PDF.")
            return

        reader = PDFService.safe_read_pdf(pdf_files[selected[0]])
        if not reader:
            return
        total = len(reader.pages)

        start = simpledialog.askinteger("Extract Pages", f"From page (1–{total}):")
        if start is None:
            return
        end = simpledialog.askinteger("Extract Pages", f"To page ({start}–{total}):")
        if end is None or start > end:
            return

        output = filedialog.asksaveasfilename(defaultextension=".pdf",
                                              initialfile=f"extracted_{start}-{end}.pdf")
        if not output:
            return

        writer = PdfWriter()
        for i in range(start - 1, end):
            writer.add_page(reader.pages[i])
        with open(output, "wb") as f:
            writer.write(f)
        messagebox.showinfo("Success", "Pages extracted successfully!")

    # ------------------ REORDER PAGES (MOVE UP/DOWN) ------------------
    @staticmethod
    def open_reorder_window(root, pdf_files, listbox):
        selected = list(listbox.curselection())
        if len(selected) != 1:
            messagebox.showwarning("Warning", "Select exactly ONE PDF.")
            return

        reader = PDFService.safe_read_pdf(pdf_files[selected[0]])
        if not reader:
            return

        pages = list(range(len(reader.pages)))

        win = tk.Toplevel(root)
        win.title("Reorder Pages")
        win.geometry("300x400")
        win.resizable(False, False)

        lb = tk.Listbox(win, selectmode=tk.EXTENDED)
        lb.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        for i in pages:
            lb.insert(tk.END, f"Page {i + 1}")

        def move_up():
            sel = lb.curselection()
            for i in sel:
                if i == 0:
                    continue
                pages[i - 1], pages[i] = pages[i], pages[i - 1]
                txt = lb.get(i)
                lb.delete(i)
                lb.insert(i - 1, txt)
                lb.selection_set(i - 1)

        def move_down():
            sel = lb.curselection()[::-1]
            for i in sel:
                if i == lb.size() - 1:
                    continue
                pages[i + 1], pages[i] = pages[i], pages[i + 1]
                txt = lb.get(i)
                lb.delete(i)
                lb.insert(i + 1, txt)
                lb.selection_set(i + 1)

        def save():
            output = filedialog.asksaveasfilename(defaultextension=".pdf")
            if not output:
                return
            writer = PdfWriter()
            for i in pages:
                writer.add_page(reader.pages[i])
            with open(output, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", "Pages reordered successfully!")
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=5)

        ttk.Button(btn_frame, text="⬆ Move Up", command=move_up).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="⬇ Move Down", command=move_down).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="Save", command=save).grid(row=1, column=0, columnspan=2, pady=10)