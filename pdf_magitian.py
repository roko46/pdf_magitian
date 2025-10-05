import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from pypdf import PdfReader, PdfWriter

# --- Functions ---
pdf_files = []

def add_pdf():
    global pdf_files
    filenames = filedialog.askopenfilenames(title="Select PDFs", filetypes=[("PDF files", "*.pdf")])
    for filename in filenames:
        if filename not in pdf_files:
            pdf_files.append(filename)
            listbox.insert(tk.END, filename)

def remove_pdf():
    global pdf_files
    selected_indices = list(listbox.curselection())
    if not selected_indices:
        messagebox.showwarning("Warning", "No PDF selected to remove!")
        return
    for index in reversed(selected_indices):
        pdf_files.pop(index)
        listbox.delete(index)

def parse_pages_to_delete(input_text):
    pages_to_delete = set()
    parts = input_text.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            start, end = part.split('-')
            pages_to_delete.update(range(int(start)-1, int(end)))
        elif part.isdigit():
            pages_to_delete.add(int(part)-1)
    return pages_to_delete

def merge_pdfs():
    if not pdf_files:
        messagebox.showwarning("Warning", "No PDFs selected!")
        return
    pdf_writer = PdfWriter()
    for file in pdf_files:
        pdf = PdfReader(file)
        for page in pdf.pages:
            pdf_writer.add_page(page)

    output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if output_file:
        with open(output_file, 'wb') as out:
            pdf_writer.write(out)
        messagebox.showinfo("Success", "PDFs merged successfully!")

def rotate_pdf():
    if not pdf_files:
        messagebox.showwarning("Warning", "No PDFs selected!")
        return

    selected_indices = list(listbox.curselection())
    if not selected_indices:
        messagebox.showwarning("Warning", "Select PDFs to rotate!")
        return

    degrees = simpledialog.askinteger("Rotate PDF", "Enter rotation degrees (90, 180, 270):", minvalue=0, maxvalue=360)
    if degrees not in [0, 90, 180, 270]:
        messagebox.showerror("Error", "Invalid rotation degree!")
        return

    for index in selected_indices:
        file = pdf_files[index]
        pdf_reader = PdfReader(file)
        pdf_writer = PdfWriter()
        for page in pdf_reader.pages:
            page.rotate(degrees)
            pdf_writer.add_page(page)

        output_file = filedialog.asksaveasfilename(title=f"Save rotated PDF: {file.split('/')[-1]}", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_file:
            with open(output_file, 'wb') as out:
                pdf_writer.write(out)
            messagebox.showinfo("Success", f"Rotated PDF saved as: {output_file}")

def delete_pages():
    if not pdf_files:
        messagebox.showwarning("Warning", "No PDFs selected!")
        return

    selected_indices = list(listbox.curselection())
    if not selected_indices:
        messagebox.showwarning("Warning", "Select PDFs to delete pages from!")
        return

    pages_input = simpledialog.askstring("Delete Pages", "Enter pages to delete (e.g., 1,3,5-7):")
    if not pages_input:
        return
    pages_to_delete = parse_pages_to_delete(pages_input)

    for index in selected_indices:
        file = pdf_files[index]
        pdf_reader = PdfReader(file)
        pdf_writer = PdfWriter()
        for i, page in enumerate(pdf_reader.pages):
            if i not in pages_to_delete:
                pdf_writer.add_page(page)

        output_file = filedialog.asksaveasfilename(title=f"Save updated PDF: {file.split('/')[-1]}", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_file:
            with open(output_file, 'wb') as out:
                pdf_writer.write(out)
            messagebox.showinfo("Success", f"Updated PDF saved as: {output_file}")

# --- Main Window ---
root = tk.Tk()
root.title("PDF Utility")
root.geometry("650x450")
root.resizable(False, False)

# --- Styles ---
style = ttk.Style(root)
style.configure("TButton", padding=6, font=("Segoe UI", 10))
style.configure("TLabel", font=("Segoe UI", 10))
style.configure("TEntry", font=("Segoe UI", 10))

# --- Frame for Listbox ---
frame = ttk.Frame(root)
frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, width=70, height=15, font=("Segoe UI", 10))
listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=listbox.yview)
scrollbar.pack(side=tk.LEFT, fill=tk.Y)
listbox.config(yscrollcommand=scrollbar.set)

# --- Buttons Frame ---
button_frame = ttk.Frame(root)
button_frame.pack(pady=10)

add_button = ttk.Button(button_frame, text="Add PDF", command=add_pdf)
add_button.grid(row=0, column=0, padx=5, pady=5)

remove_button = ttk.Button(button_frame, text="Remove PDF", command=remove_pdf)
remove_button.grid(row=0, column=1, padx=5, pady=5)

merge_button = ttk.Button(button_frame, text="Merge PDFs", command=merge_pdfs)
merge_button.grid(row=0, column=2, padx=5, pady=5)

rotate_button = ttk.Button(button_frame, text="Rotate PDFs", command=rotate_pdf)
rotate_button.grid(row=0, column=3, padx=5, pady=5)

delete_button = ttk.Button(button_frame, text="Delete Pages", command=delete_pages)
delete_button.grid(row=0, column=4, padx=5, pady=5)

# --- Run App ---
root.mainloop()
