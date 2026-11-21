import importlib
import os
import site
import subprocess
import sys


def _python_exe():
    """
    Return an interpreter suitable for running pip. A .pyw normally runs under
    pythonw.exe, which has no console / valid std handles and can make the pip
    subprocess fail silently, so prefer the sibling python.exe on Windows.
    """
    exe = sys.executable
    if os.name == "nt" and exe.lower().endswith("pythonw.exe"):
        candidate = exe[:-len("pythonw.exe")] + "python.exe"
        if os.path.exists(candidate):
            return candidate
    return exe


def _run_pip(args):
    kwargs = dict(stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                  stderr=subprocess.STDOUT, text=True)
    if os.name == "nt":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW  # no console flash
    return subprocess.run([_python_exe(), "-m", "pip", "install"] + args, **kwargs)


def _refresh_import_paths():
    """Make freshly installed packages importable without restarting."""
    importlib.invalidate_caches()
    try:
        user_site = site.getusersitepackages()
        if os.path.isdir(user_site) and user_site not in sys.path:
            sys.path.append(user_site)
    except Exception:
        pass


# Function to check and install the module if needed
def install_module(import_name, pip_name=None):
    pip_name = pip_name or import_name
    try:
        importlib.import_module(import_name)
        return
    except ImportError:
        pass

    result = _run_pip([pip_name])
    if result.returncode != 0:
        # Fall back to a per-user install (e.g. no write access to site-packages)
        result = _run_pip(["--user", pip_name])
    if result.returncode != 0:
        raise RuntimeError(f"pip could not install {pip_name}:\n{result.stdout[-1500:]}")

    _refresh_import_paths()
    importlib.import_module(import_name)  # raises ImportError if still not found


def _fatal(message):
    """Show an error dialog (pythonw has no console, so print() is invisible)."""
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        root = _tk.Tk()
        root.withdraw()
        _mb.showerror("PDF Merger - setup failed", message)
        root.destroy()
    except Exception:
        print(message)
    sys.exit(1)


if not getattr(sys, 'frozen', False):  # Only install if not running as a bundled executable
    modules = [
        ('tkinterdnd2', 'tkinterdnd2'),
        ('PyPDF2', 'PyPDF2'),
        ('fitz', 'PyMuPDF'),  # import name is "fitz", pip package is "PyMuPDF"
    ]
    for import_name, pip_name in modules:
        try:
            install_module(import_name, pip_name)
        except Exception as e:
            _fatal(f"Could not install required package '{pip_name}'.\n\n{e}\n\n"
                   f"Try running:  python -m pip install {pip_name}")


import base64
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PyPDF2 import PdfMerger
try:
    import pymupdf as fitz  # PyMuPDF >= 1.24.3 (new name; "fitz" import is deprecated)
except ImportError:
    import fitz  # older PyMuPDF
from urllib.parse import urlparse, unquote
from urllib.request import url2pathname

# Page thumbnail geometry (pixels)
THUMB_W, THUMB_H = 130, 175
CELL_W, CELL_H = THUMB_W + 12, THUMB_H + 32
CELL_PAD = 4

COLOR_BG = "#e6e6e6"
COLOR_CELL = "#f7f7f7"
COLOR_CELL_SELECTED = "#cfe3ff"
COLOR_BORDER = "#b8b8b8"
COLOR_BORDER_SELECTED = "#2a6fdb"
COLOR_DELETED = "#d62828"


class PdfEntry:
    """One PDF in the merge list, plus the set of its pages (0-based) marked as deleted."""
    __slots__ = ("path", "pages", "deleted")

    def __init__(self, path, pages):
        self.path = path
        self.pages = pages
        self.deleted = set()

    @property
    def name(self):
        return os.path.basename(self.path)

    @property
    def kept(self):
        return self.pages - len(self.deleted)

    def label(self):
        if self.deleted:
            return f"{self.name}   [{self.kept}/{self.pages} pages]"
        return f"{self.name}   [{self.pages} pages]"

    def kept_ranges(self):
        """Contiguous (start, stop) runs of pages that are NOT deleted (stop is exclusive)."""
        runs, start = [], None
        for i in range(self.pages):
            if i in self.deleted:
                if start is not None:
                    runs.append((start, i))
                    start = None
            elif start is None:
                start = i
        if start is not None:
            runs.append((start, self.pages))
        return runs


class PDFMergerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Merger")
        self.geometry("980x640")
        self.minsize(780, 480)

        self.entries = []            # list[PdfEntry], same order as the file listbox
        self.cells = []              # thumbnail cells of the currently shown PDF
        self.selected_pages = set()  # selected page indexes of the currently shown PDF
        self._shown_entry = None
        self._anchor = None          # last clicked page (for shift-click ranges)
        self._gen = 0                # bumped to cancel an in-progress thumbnail render
        self._cols = 0

        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self._build_file_panel()
        self._build_pages_panel()
        self._build_bottom_bar()

        # one global wheel binding; only acts when the pointer is over the page grid
        self.bind_all("<MouseWheel>", self._on_mousewheel)
        self.bind_all("<Button-4>", self._on_mousewheel)
        self.bind_all("<Button-5>", self._on_mousewheel)

        self._update_quality_state()
        self._update_page_buttons()

    # ----------------------------
    # UI construction
    # ----------------------------
    def _build_file_panel(self):
        left = tk.Frame(self)
        left.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)
        left.rowconfigure(1, weight=1)

        tk.Label(left, text="PDF files (drop here)", anchor="w").grid(row=0, column=0, columnspan=2, sticky="ew")

        self.file_listbox = tk.Listbox(left, selectmode=tk.SINGLE, width=38, exportselection=False)
        self.file_listbox.grid(row=1, column=0, sticky="nsew")
        sb = tk.Scrollbar(left, command=self.file_listbox.yview)
        sb.grid(row=1, column=1, sticky="ns")
        self.file_listbox.config(yscrollcommand=sb.set)
        self.file_listbox.bind("<<ListboxSelect>>", lambda e: self._show_pages())

        # Drag & drop support
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.drop)

        btn_frame = tk.Frame(left)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=(8, 0))
        tk.Button(btn_frame, text="Move Up", command=self.move_up).grid(row=0, column=0, padx=3)
        tk.Button(btn_frame, text="Move Down", command=self.move_down).grid(row=0, column=1, padx=3)
        tk.Button(btn_frame, text="Remove", command=self.remove_file).grid(row=0, column=2, padx=3)

    def _build_pages_panel(self):
        right = tk.LabelFrame(self, text="Pages preview")
        right.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        toolbar = tk.Frame(right)
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
        self.delete_btn = tk.Button(toolbar, text="Delete page(s)", command=self.delete_selected_pages)
        self.delete_btn.pack(side=tk.LEFT, padx=(0, 4))
        self.restore_btn = tk.Button(toolbar, text="Restore page(s)", command=self.restore_selected_pages)
        self.restore_btn.pack(side=tk.LEFT, padx=4)
        self.preview_btn = tk.Button(toolbar, text="Preview", command=self.preview_selected_page)
        self.preview_btn.pack(side=tk.LEFT, padx=4)
        self.pages_info = tk.Label(right, text="", anchor="w")
        self.pages_info.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 4))

        self.pages_canvas = tk.Canvas(right, bg=COLOR_BG, highlightthickness=0)
        self.pages_canvas.grid(row=2, column=0, sticky="nsew")
        vsb = tk.Scrollbar(right, orient=tk.VERTICAL, command=self.pages_canvas.yview)
        vsb.grid(row=2, column=1, sticky="ns")
        self.pages_canvas.config(yscrollcommand=vsb.set)

        self.pages_inner = tk.Frame(self.pages_canvas, bg=COLOR_BG)
        self.pages_canvas.create_window((0, 0), window=self.pages_inner, anchor="nw")
        self.pages_inner.bind(
            "<Configure>", lambda e: self.pages_canvas.config(scrollregion=self.pages_canvas.bbox("all")))
        self.pages_canvas.bind("<Configure>", self._on_canvas_resize)

        # keyboard: Delete removes the selected pages, Ctrl+A selects all
        self.pages_canvas.bind("<Delete>", self.delete_selected_pages)
        self.pages_canvas.bind("<Control-a>", self._select_all_pages)
        self.pages_canvas.drop_target_register(DND_FILES)
        self.pages_canvas.dnd_bind('<<Drop>>', self.drop)

    def _build_bottom_bar(self):
        bar = tk.Frame(self)
        bar.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 10))
        bar.columnconfigure(0, weight=1)

        options_frame = tk.Frame(bar)
        options_frame.grid(row=0, column=0, sticky="w")

        self.print_as_picture_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            options_frame,
            text="Print as picture (flatten pages to images before merging)",
            variable=self.print_as_picture_var,
            command=self._update_quality_state
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        self.quality_label = tk.Label(options_frame, text="Image quality:")
        self.quality_label.grid(row=1, column=0, padx=(20, 2), pady=(4, 0), sticky="e")
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

        tk.Button(bar, text="Merge PDFs", width=14, command=self.merge_pdfs).grid(row=0, column=1, sticky="e")

    def _update_quality_state(self):
        """Enable the Image quality dropdown only when Print as picture is ticked."""
        state = tk.NORMAL if self.print_as_picture_var.get() else tk.DISABLED
        self.quality_label.config(state=state)
        self.quality_menu.config(state=state)

    def _update_page_buttons(self):
        has_entry = self._shown_entry is not None
        sel = self.selected_pages
        entry = self._shown_entry
        can_delete = bool(sel) and entry is not None and bool(sel - entry.deleted)
        can_restore = bool(sel) and entry is not None and bool(sel & entry.deleted)
        self.delete_btn.config(state=tk.NORMAL if can_delete else tk.DISABLED)
        self.restore_btn.config(state=tk.NORMAL if can_restore else tk.DISABLED)
        self.preview_btn.config(state=tk.NORMAL if has_entry and entry.pages else tk.DISABLED)

    # ----------------------------
    # Drag & drop handling
    # ----------------------------
    def drop(self, event):
        added = False
        for path in self._parse_drop_data(event.data):
            if path.lower().endswith(".pdf"):
                if self._add_file(path):
                    added = True
            else:
                messagebox.showwarning("Invalid file", f"Not a PDF: {path}")
        if added:
            last = len(self.entries) - 1
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(last)
            self.file_listbox.see(last)
            self._show_pages()

    def _add_file(self, path):
        """Open the PDF just long enough to read its page count; returns True if added."""
        try:
            doc = fitz.open(path)
            try:
                if doc.needs_pass:
                    messagebox.showwarning("Password protected", f"Password-protected PDFs are not supported:\n{path}")
                    return False
                pages = doc.page_count
            finally:
                doc.close()
        except Exception as e:
            messagebox.showerror("Cannot open PDF", f"{path}\n\n{e}")
            return False
        if pages == 0:
            messagebox.showwarning("Empty PDF", f"This PDF has no pages:\n{path}")
            return False
        entry = PdfEntry(path, pages)
        self.entries.append(entry)
        self.file_listbox.insert(tk.END, entry.label())
        return True

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
    def _current_index(self):
        sel = self.file_listbox.curselection()
        return sel[0] if sel else None

    def _current_entry(self):
        idx = self._current_index()
        return self.entries[idx] if idx is not None else None

    def _reload_listbox(self, select=None):
        self.file_listbox.delete(0, tk.END)
        for e in self.entries:
            self.file_listbox.insert(tk.END, e.label())
        if select is not None and 0 <= select < len(self.entries):
            self.file_listbox.selection_set(select)
            self.file_listbox.see(select)

    def _refresh_item(self, entry):
        """Update one listbox row (e.g. the kept-page count) without losing the selection."""
        for idx, e in enumerate(self.entries):
            if e is entry:
                selected = idx in self.file_listbox.curselection()
                self.file_listbox.delete(idx)
                self.file_listbox.insert(idx, entry.label())
                if selected:
                    self.file_listbox.selection_set(idx)
                return

    def move_up(self):
        idx = self._current_index()
        if idx is not None and idx > 0:
            self.entries[idx - 1], self.entries[idx] = self.entries[idx], self.entries[idx - 1]
            self._reload_listbox(select=idx - 1)

    def move_down(self):
        idx = self._current_index()
        if idx is not None and idx < len(self.entries) - 1:
            self.entries[idx + 1], self.entries[idx] = self.entries[idx], self.entries[idx + 1]
            self._reload_listbox(select=idx + 1)

    def remove_file(self):
        idx = self._current_index()
        if idx is None:
            return
        del self.entries[idx]
        self._reload_listbox(select=min(idx, len(self.entries) - 1))
        self._show_pages()

    # ----------------------------
    # Page thumbnails
    # ----------------------------
    def _clear_cells(self):
        for cell in self.cells:
            cell["canvas"].destroy()
        self.cells = []
        self._cols = 0

    def _show_pages(self):
        """Rebuild the thumbnail grid for the PDF selected in the file list."""
        self._gen += 1  # cancels any render still running for the previous selection
        self._clear_cells()
        self.selected_pages = set()
        self._anchor = None
        entry = self._current_entry()
        self._shown_entry = entry
        self.pages_canvas.yview_moveto(0)

        if entry is None:
            self._update_info()
            self._update_page_buttons()
            return

        for i in range(entry.pages):
            c = tk.Canvas(self.pages_inner, width=CELL_W, height=CELL_H, highlightthickness=0, bg=COLOR_BG)
            c.bind("<Button-1>", lambda e, i=i: self._on_page_click(i, e))
            c.bind("<Double-Button-1>", lambda e, i=i: self.preview_page(entry, i))
            self.cells.append({"canvas": c, "img": None, "state": "loading"})
            self._draw_cell(i)
        self._layout_cells()
        self._update_info()
        self._update_page_buttons()

        try:
            doc = fitz.open(entry.path)
        except Exception as e:
            messagebox.showerror("Cannot open PDF", f"{entry.path}\n\n{e}")
            return
        self.after(1, self._render_next, self._gen, doc, entry, 0)

    def _render_next(self, gen, doc, entry, i):
        """Render thumbnails one page per event-loop tick so the UI stays responsive."""
        if gen != self._gen or i >= entry.pages:
            doc.close()
            return
        cell = self.cells[i]
        try:
            page = doc[i]
            r = page.rect
            scale = min(THUMB_W / r.width, THUMB_H / r.height)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB, alpha=False)
            cell["img"] = tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")))
            cell["state"] = "ok"
        except Exception:
            cell["state"] = "failed"
        self._draw_cell(i)
        self.after(1, self._render_next, gen, doc, entry, i + 1)

    def _draw_cell(self, i):
        entry = self._shown_entry
        cell = self.cells[i]
        c = cell["canvas"]
        c.delete("all")
        deleted = i in entry.deleted
        selected = i in self.selected_pages

        c.create_rectangle(1, 1, CELL_W - 1, CELL_H - 1,
                           fill=COLOR_CELL_SELECTED if selected else COLOR_CELL,
                           outline=COLOR_BORDER_SELECTED if selected else COLOR_BORDER,
                           width=3 if selected else 1)

        cx, cy = CELL_W // 2, (CELL_H - 22) // 2 + 2
        if cell["img"] is not None:
            img = cell["img"]
            c.create_image(cx, cy, image=img)
            w, h = img.width(), img.height()
            c.create_rectangle(cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2, outline="#888888")
        else:
            w, h = THUMB_W, THUMB_H
            c.create_text(cx, cy, text="…" if cell["state"] == "loading" else "(no preview)", fill="#777777")

        if deleted:
            x0, y0, x1, y1 = cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2
            c.create_line(x0, y0, x1, y1, fill=COLOR_DELETED, width=3)
            c.create_line(x0, y1, x1, y0, fill=COLOR_DELETED, width=3)

        c.create_text(CELL_W // 2, CELL_H - 12,
                      text=f"Page {i + 1}" + (" (deleted)" if deleted else ""),
                      fill=COLOR_DELETED if deleted else "#222222")

    def _layout_cells(self):
        width = self.pages_canvas.winfo_width()
        cols = max(1, width // (CELL_W + 2 * CELL_PAD))
        self._cols = cols
        for i, cell in enumerate(self.cells):
            cell["canvas"].grid(row=i // cols, column=i % cols, padx=CELL_PAD, pady=CELL_PAD)

    def _on_canvas_resize(self, event):
        cols = max(1, event.width // (CELL_W + 2 * CELL_PAD))
        if self.cells and cols != self._cols:
            self._layout_cells()

    def _is_over_pages(self, widget):
        while widget is not None:
            if widget is self.pages_canvas:
                return True
            widget = getattr(widget, "master", None)
        return False

    def _on_mousewheel(self, event):
        try:
            widget = self.winfo_containing(event.x_root, event.y_root)
        except Exception:
            return
        if not self._is_over_pages(widget):
            return
        if getattr(event, "num", None) == 4:
            step = -1
        elif getattr(event, "num", None) == 5:
            step = 1
        else:
            step = -1 if event.delta > 0 else 1
        self.pages_canvas.yview_scroll(step * 2, "units")

    def _update_info(self):
        entry = self._shown_entry
        if entry is None:
            self.pages_info.config(text="Select a PDF to see its pages")
            return
        text = f"{entry.name}: {entry.kept} of {entry.pages} pages will be merged"
        if self.selected_pages:
            text += f"  |  {len(self.selected_pages)} selected"
        self.pages_info.config(text=text)

    # ----------------------------
    # Page selection / deletion
    # ----------------------------
    def _set_selection(self, new_selection):
        changed = self.selected_pages ^ new_selection
        self.selected_pages = new_selection
        for i in changed:
            self._draw_cell(i)
        self._update_info()
        self._update_page_buttons()

    def _on_page_click(self, i, event):
        self.pages_canvas.focus_set()
        if event.state & 0x1 and self._anchor is not None:      # Shift: range
            lo, hi = sorted((self._anchor, i))
            self._set_selection(set(range(lo, hi + 1)))
        elif event.state & 0x4:                                 # Ctrl: toggle
            new = set(self.selected_pages)
            new.symmetric_difference_update({i})
            self._anchor = i
            self._set_selection(new)
        else:
            self._anchor = i
            self._set_selection({i})

    def _select_all_pages(self, event=None):
        if self._shown_entry is not None:
            self._set_selection(set(range(self._shown_entry.pages)))
        return "break"

    def delete_selected_pages(self, event=None):
        entry = self._shown_entry
        if entry is not None and self.selected_pages:
            changed = self.selected_pages - entry.deleted
            entry.deleted |= self.selected_pages
            self._after_pages_changed(entry, changed)
        return "break"

    def restore_selected_pages(self, event=None):
        entry = self._shown_entry
        if entry is not None and self.selected_pages:
            changed = self.selected_pages & entry.deleted
            entry.deleted -= self.selected_pages
            self._after_pages_changed(entry, changed)
        return "break"

    def _after_pages_changed(self, entry, changed_pages):
        """Refresh the list row and (if shown) the thumbnails after pages were deleted/restored."""
        self._refresh_item(entry)
        if entry is self._shown_entry:
            for i in changed_pages:
                if i < len(self.cells):
                    self._draw_cell(i)
            self._update_info()
            self._update_page_buttons()

    # ----------------------------
    # Full-size preview window
    # ----------------------------
    def preview_selected_page(self):
        entry = self._shown_entry
        if entry is None or not entry.pages:
            return
        page = min(self.selected_pages) if self.selected_pages else 0
        self.preview_page(entry, page)

    def preview_page(self, entry, page_no):
        top = tk.Toplevel(self)
        top.transient(self)
        state = {"page": page_no, "img": None}

        img_label = tk.Label(top, bd=1, relief=tk.SOLID)
        img_label.pack(padx=10, pady=(10, 4))
        status = tk.Label(top)
        status.pack()
        bar = tk.Frame(top)
        bar.pack(pady=(4, 10))

        def still_listed():
            return any(e is entry for e in self.entries)

        def render():
            if not still_listed():
                top.destroy()
                return
            p = state["page"]
            try:
                doc = fitz.open(entry.path)
                try:
                    page = doc[p]
                    r = page.rect
                    max_w = int(self.winfo_screenwidth() * 0.6)
                    max_h = int(self.winfo_screenheight() * 0.70)
                    scale = min(max_w / r.width, max_h / r.height, 3.0)
                    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB, alpha=False)
                    data = base64.b64encode(pix.tobytes("png"))
                finally:
                    doc.close()
                state["img"] = tk.PhotoImage(data=data)
                img_label.config(image=state["img"], text="")
            except Exception as e:
                state["img"] = None
                img_label.config(image="", text=f"Cannot render this page:\n{e}", width=50, height=10)
            deleted = p in entry.deleted
            top.title(f"{entry.name} - page {p + 1}/{entry.pages}")
            status.config(
                text=f"Page {p + 1} of {entry.pages}" + ("   -   DELETED (will be skipped)" if deleted else ""),
                fg=COLOR_DELETED if deleted else "black")
            toggle_btn.config(text="Restore page" if deleted else "Delete page")
            prev_btn.config(state=tk.NORMAL if p > 0 else tk.DISABLED)
            next_btn.config(state=tk.NORMAL if p < entry.pages - 1 else tk.DISABLED)

        def go(delta):
            if not still_listed():
                top.destroy()
                return "break"
            new = state["page"] + delta
            if 0 <= new < entry.pages:
                state["page"] = new
                render()
            return "break"

        def toggle(event=None):
            if not still_listed():
                top.destroy()
                return "break"
            p = state["page"]
            if p in entry.deleted:
                entry.deleted.discard(p)
            else:
                entry.deleted.add(p)
            self._after_pages_changed(entry, {p})
            render()
            return "break"

        prev_btn = tk.Button(bar, text="< Previous", command=lambda: go(-1))
        prev_btn.pack(side=tk.LEFT, padx=4)
        toggle_btn = tk.Button(bar, text="Delete page", command=toggle, width=14)
        toggle_btn.pack(side=tk.LEFT, padx=4)
        next_btn = tk.Button(bar, text="Next >", command=lambda: go(1))
        next_btn.pack(side=tk.LEFT, padx=4)
        tk.Button(bar, text="Close", command=top.destroy).pack(side=tk.LEFT, padx=(16, 4))

        top.bind("<Left>", lambda e: go(-1))
        top.bind("<Right>", lambda e: go(1))
        top.bind("<Delete>", toggle)
        top.bind("<Escape>", lambda e: top.destroy())
        top.focus_set()
        render()

    # ----------------------------
    # Merge PDFs
    # ----------------------------
    def merge_pdfs(self):
        if not self.entries:
            messagebox.showwarning("No files", "Please add PDF files first.")
            return
        if sum(e.kept for e in self.entries) == 0:
            messagebox.showwarning("No pages", "All pages are deleted - there is nothing to merge.")
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
        """Standard merge: keeps each PDF's pages as-is (vector text, layers, etc.), skipping deleted pages."""
        merger = PdfMerger()
        try:
            for e in self.entries:
                if not e.deleted:
                    merger.append(e.path)  # whole file, keeps its outlines
                else:
                    for start, stop in e.kept_ranges():
                        merger.append(e.path, pages=(start, stop))
            merger.write(save_path)
        finally:
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
        "Print as picture" merge: renders every kept page of every input PDF to
        an image first (like printing to an image and back to PDF), then places
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
            for e in self.entries:
                src = fitz.open(e.path)
                try:
                    for i in range(src.page_count):
                        if i in e.deleted:
                            continue
                        page = src[i]
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
    if getattr(sys, 'frozen', False):  # Check if running as a bundled executable
        try:
            import pyi_splash
            pyi_splash.close()
        except ImportError:
            # pyi_splash might not be available if not running through PyInstaller
            pass

    app = PDFMergerApp()
    app.mainloop()
