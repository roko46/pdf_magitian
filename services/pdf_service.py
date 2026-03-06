from pypdf import PdfReader, PdfWriter
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from utils.validators import parse_page_ranges

class PDFService:

    @staticmethod
    def safe_read_pdf(path):
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except:
                    messagebox.showerror("Error", f"Encrypted PDF not supported:\n{path}")
                    return None
            return reader
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open PDF:\n{path}\n\n{e}")
            return None

    # ------------------ MERGE ------------------
    @staticmethod
    def merge_pdfs(pdf_files):
        if len(pdf_files) < 2:
            messagebox.showwarning("Warning", "Select at least 2 PDFs to merge.")
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
                page.rotate(degrees)  # modern pypdf API
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
        end = min(end, total)
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

    # ------------------ REORDER PAGES WITH DRAG & DROP ------------------
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

        reorder_listbox = tk.Listbox(win, selectmode=tk.EXTENDED)
        reorder_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        for i in pages:
            reorder_listbox.insert(tk.END, f"Page {i + 1}")

        # -------- DRAG & DROP REORDER --------
        def on_start_drag(event):
            reorder_listbox.drag_index = reorder_listbox.nearest(event.y)

        def on_drag_motion(event):
            i = reorder_listbox.nearest(event.y)
            if i == reorder_listbox.drag_index:
                return
            reorder_listbox.insert(i, reorder_listbox.get(reorder_listbox.drag_index))
            reorder_listbox.delete(reorder_listbox.drag_index + (1 if i < reorder_listbox.drag_index else 0))
            reorder_listbox.drag_index = i

        reorder_listbox.bind("<Button-1>", on_start_drag)
        reorder_listbox.bind("<B1-Motion>", on_drag_motion)

        # -------- SAVE --------
        def save():
            output = filedialog.asksaveasfilename(defaultextension=".pdf")
            if not output:
                return
            writer = PdfWriter()
            for i in range(reorder_listbox.size()):
                page_index = int(reorder_listbox.get(i).split()[1]) - 1
                writer.add_page(reader.pages[page_index])
            with open(output, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", "Pages reordered successfully!")
            win.destroy()

        tk.Button(win, text="Save", command=save).pack(pady=10)