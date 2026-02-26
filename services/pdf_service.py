from pypdf import PdfReader, PdfWriter
from tkinter import filedialog, messagebox, simpledialog
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