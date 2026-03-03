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
            messagebox.showerror("Error", str(e))
            return None

    # ------------------ REORDER PAGES ------------------

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

    # ------------------ OSTALI POSTOJEĆI METODI ------------------

    @staticmethod
    def merge_pdfs(pdf_files):
        writer = PdfWriter()
        for f in pdf_files:
            r = PDFService.safe_read_pdf(f)
            if not r:
                return
            for p in r.pages:
                writer.add_page(p)
        out = filedialog.asksaveasfilename(defaultextension=".pdf")
        if out:
            with open(out, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", "PDFs merged!")