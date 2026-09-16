import subprocess
import sys

# Function to check and install the module if needed
def install_module(import_name, pip_name=None):
    pip_name = pip_name or import_name
    try:
        __import__(import_name)
    except ImportError:
        print(f"{pip_name} not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])

if not getattr(sys, 'frozen', False):  # Only install if not running as a bundled executable
    try:
        modules = [
            ('tkinterdnd2', 'tkinterdnd2'),
            ('PyPDF2', 'PyPDF2'),
            ('fitz', 'PyMuPDF'),  # import name is "fitz", pip package is "PyMuPDF"
        ]
        for import_name, pip_name in modules:
            install_module(import_name, pip_name)
    except Exception as e:
        print(f"Error installing modules: {e}")
        sys.exit(1)

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PyPDF2 import PdfMerger
import fitz  # PyMuPDF, used for the "Print as picture" (rasterize) option
import os as os
from urllib.parse import urlparse, unquote
from urllib.request import url2pathname


class PDFMergerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Merger")
        self.geometry("500x430")

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

        # Options
        options_frame = tk.Frame(self)
        options_frame.pack(pady=5)

        self.print_as_picture_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            options_frame,
            text="Print as picture (flatten pages to images before merging)",
            variable=self.print_as_picture_var,
            command=self._update_quality_state
        ).grid(row=0, column=0, columnspan=2, padx=5, sticky="w")

        self.quality_label = tk.Label(options_frame, text="Image quality:")
        self.quality_label.grid(row=1, column=0, padx=(5, 2), pady=(4, 0), sticky="e")
        self.picture_quality_var = tk.StringVar(value="Medium (150 DPI)")
        self.quality_menu = tk.OptionMenu(
            options_frame,
            self.picture_quality_var,
            "Low (100 DPI, smallest file)",
            "Medium (150 DPI)",
            "High (200 DPI)",
            "Very high (300 DPI, largest file)",
        )
        self.quality_menu.grid(row=1, column=1, pady=(4, 0), sticky="w")

        self._update_quality_state()

    def _update_quality_state(self):
        """Enable the Image quality dropdown only when Print as picture is ticked."""
        state = tk.NORMAL if self.print_as_picture_var.get() else tk.DISABLED
        self.quality_label.config(state=state)
        self.quality_menu.config(state=state)

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

        try:
            if self.print_as_picture_var.get():
                self._merge_as_pictures(save_path)
            else:
                self._merge_normal(save_path)
            messagebox.showinfo("Success", f"Merged PDF saved at:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _merge_normal(self, save_path):
        """Standard merge: keeps each PDF's pages as-is (vector text, layers, etc.)."""
        merger = PdfMerger()
        for i in range(self.file_listbox.size()):
            merger.append(self.file_listbox.get(i))
        merger.write(save_path)
        merger.close()

    # DPI / JPEG quality behind each "Image quality" menu choice. JPEG at a
    # moderate quality is used instead of PNG because it is dramatically
    # smaller for scanned/rendered pages, with little visible difference.
    _QUALITY_PRESETS = {
        "Low (100 DPI, smallest file)": (100, 60),
        "Medium (150 DPI)": (150, 75),
        "High (200 DPI)": (200, 85),
        "Very high (300 DPI, largest file)": (300, 90),
    }

    def _merge_as_pictures(self, save_path):
        """
        "Print as picture" merge: renders every page of every input PDF to an
        image first (like printing to an image and back to PDF), then places
        that image on a same-size output page. This flattens text, forms,
        annotations, and layers into a single picture per page - useful when
        a source PDF is malformed, has fonts that render oddly, or you simply
        want a flattened, non-editable copy.

        Pages are encoded as JPEG (not PNG) and rendered at the DPI chosen in
        the "Image quality" dropdown, since that is what actually controls
        the resulting file size.
        """
        dpi, jpg_quality = self._QUALITY_PRESETS[self.picture_quality_var.get()]
        zoom = dpi / 72  # PDF points are 72 per inch
        mat = fitz.Matrix(zoom, zoom)

        out_doc = fitz.open()
        try:
            for i in range(self.file_listbox.size()):
                path = self.file_listbox.get(i)
                src = fitz.open(path)
                try:
                    for page in src:
                        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB, alpha=False)
                        img_bytes = pix.tobytes("jpg", jpg_quality=jpg_quality)
                        new_page = out_doc.new_page(width=page.rect.width, height=page.rect.height)
                        new_page.insert_image(new_page.rect, stream=img_bytes)
                finally:
                    src.close()
            # deflate/garbage-collect the output so page objects and any
            # duplicate resources are compressed as tightly as possible
            out_doc.save(save_path, garbage=4, deflate=True)
        finally:
            out_doc.close()


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
