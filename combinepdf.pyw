import subprocess
import sys

# Function to check and install the module if needed
def install_module(module_name):
    try:
        __import__(module_name)
    except ImportError:
        print(f"{module_name} not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", module_name])

if not getattr(sys, 'frozen', False):   # Only install if not running as a bundled executable
    try:
        modules = ['tkinterdnd2', 'PyPDF2']

        for module in modules:
            install_module(module)
    except Exception as e:
        print(f"Error installing modules: {e}")
        sys.exit(1)

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PyPDF2 import PdfMerger
import os as os
from urllib.parse import urlparse, unquote
from urllib.request import url2pathname


class PDFMergerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Merger")
        self.geometry("500x400")

        # Listbox for files
        self.file_listbox = tk.Listbox(self, selectmode=tk.SINGLE, width=60, height=15)
        self.file_listbox.pack(pady=10)

        # Drag & drop support
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.drop)

        # Control buttons
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=5)

        tk.Button(btn_frame, text="Move Up", command=self.move_up).grid(row=0, column=0, padx=5)
        tk.Button(btn_frame, text="Move Down", command=self.move_down).grid(row=0, column=1, padx=5)
        tk.Button(btn_frame, text="Remove", command=self.remove_file).grid(row=0, column=2, padx=5)
        tk.Button(btn_frame, text="Merge PDFs", command=self.merge_pdfs).grid(row=0, column=3, padx=5)

    # ----------------------------
    # Drag & drop handling
    # ----------------------------
    def drop(self, event):
        for path in self._parse_drop_data(event.data):
            if path.lower().endswith(".pdf"):
                self.file_listbox.insert(tk.END, path)
            else:
                messagebox.showwarning("Invalid file", f"Not a PDF: {path}")

    def _parse_drop_data(self, data):
        """
        Robustly parse TkDnD drop data:
        - Uses Tk's splitlist to handle spaces/braces properly.
        - Converts file:// URLs to OS paths.
        """
        items = self.tk.splitlist(data)
        paths = []
        for item in items:
            paths.append(self._to_os_path(item))
        return paths

    @staticmethod
    def _to_os_path(item: str) -> str:
        # Convert file:// URL → local file path; otherwise normalize as-is.
        if item.startswith("file://"):
            p = urlparse(item)
            local_path = url2pathname(unquote(p.path))
            if p.netloc:  # UNC share like file://server/share/...
                return os.path.normpath(rf"\\{p.netloc}{local_path}")
            return os.path.normpath(local_path)
        return os.path.normpath(item)

    # ----------------------------
    # Listbox controls
    # ----------------------------
    def move_up(self):
        try:
            idx = self.file_listbox.curselection()[0]
            if idx > 0:
                text = self.file_listbox.get(idx)
                self.file_listbox.delete(idx)
                self.file_listbox.insert(idx - 1, text)
                self.file_listbox.selection_set(idx - 1)
        except IndexError:
            pass

    def move_down(self):
        try:
            idx = self.file_listbox.curselection()[0]
            if idx < self.file_listbox.size() - 1:
                text = self.file_listbox.get(idx)
                self.file_listbox.delete(idx)
                self.file_listbox.insert(idx + 1, text)
                self.file_listbox.selection_set(idx + 1)
        except IndexError:
            pass

    def remove_file(self):
        try:
            idx = self.file_listbox.curselection()[0]
            self.file_listbox.delete(idx)
        except IndexError:
            pass

    # ----------------------------
    # Merge PDFs
    # ----------------------------
    def merge_pdfs(self):
        if self.file_listbox.size() == 0:
            messagebox.showwarning("No files", "Please add PDF files first.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                 filetypes=[("PDF files", "*.pdf")],
                                                 title="Save merged PDF as")
        if not save_path:
            return

        merger = PdfMerger()
        for i in range(self.file_listbox.size()):
            merger.append(self.file_listbox.get(i))

        try:
            merger.write(save_path)
            merger.close()
            messagebox.showinfo("Success", f"Merged PDF saved at:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    import sys
    
    if getattr(sys, 'frozen', False):  # Check if running as a bundled executable
        try:
            import pyi_splash
            pyi_splash.close()
        except ImportError:
            # pyi_splash might not be available if not running through PyInstaller
            pass
    
    app = PDFMergerApp()
    app.mainloop()

