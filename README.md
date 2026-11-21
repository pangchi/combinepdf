# combinepdf

Python GUI that lets users drag and drop PDF files, arrange them, and merge them into a single PDF.

## Features

- Drag and drop PDFs into the list (or arrange manually)
- Move Up / Move Down / Remove to reorder before merging
- **Page preview**: select a PDF in the list to see thumbnails of all its pages on the right (rendered progressively, so big PDFs don't freeze the window). Double-click a thumbnail, or press **Preview**, for a full-size view with Previous / Next (also the Left / Right arrow keys).
- **Delete individual pages**: click a thumbnail to select it (Ctrl-click to add or remove pages, Shift-click for a range, Ctrl+A for all), then press **Delete page(s)** or the Delete key. Deleted pages get a red cross and are skipped when merging; **Restore page(s)** brings them back. The file list shows how many pages will be kept, e.g. `report.pdf [4/6 pages]`. The full-size preview window also has a Delete / Restore page button (Delete key). Your original files are never modified.
- Merge PDFs into a single output file
- **Print as picture** option: when checked, every page of every input PDF is rendered to a JPEG image first (like printing to an image and back to PDF) before being placed into the merged output. This flattens text, form fields, annotations, and layers into a single picture per page — useful when a source PDF is malformed, renders fonts oddly, or you just want a flattened, non-editable copy. Leave it unchecked for a normal, lossless merge that keeps each PDF's pages exactly as they are.
  - An **Image quality** dropdown controls the DPI and JPEG quality used when flattening (Low/Medium/High/Very high). It is greyed out unless "Print as picture" is ticked. Lower settings give a much smaller output file at the cost of sharpness — Medium (150 DPI) is a good default; use Low for small files or Very high only when fine detail matters.

## Requirements

Installed automatically on first run (when not bundled as an executable):

- `tkinterdnd2` — drag and drop support
- `PyPDF2` — normal (non-flattened) merging
- `PyMuPDF` (imported as `pymupdf`, or `fitz` in older versions) — page thumbnails/preview and the "Print as picture" option

If automatic installation fails, an error dialog shows the reason and the manual command, e.g. `python -m pip install tkinterdnd2`.

## Usage

1. Run `combinepdf.pyw`.
2. Drag PDF files into the list, or reorder/remove as needed.
3. Select a PDF to preview its pages; select and delete any pages you don't want in the merged file.
4. Check "Print as picture" if you want the merged pages flattened to images, and choose an image quality.
5. Click "Merge PDFs" and choose where to save the result.

## Changelog

- Fixed filename with space issue.
- Added splash screen for PyInstaller build.
- Added "Print as picture" option to flatten pages to images before merging.
- Flattened pages are now JPEG-compressed, with an Image quality dropdown (greyed out unless Print as picture is ticked).
- Fixed first-run auto-install failing with "tkinterdnd2 not found": pip now runs with `python.exe` instead of `pythonw.exe`, falls back to `--user`, refreshes import paths after installing, and reports any failure in a dialog instead of exiting silently.
- Added page preview (thumbnails + full-size viewer) and individual page deletion/restore; deleted pages are skipped in both normal and "Print as picture" merges. Password-protected or unreadable PDFs are now reported when added instead of failing at merge time.
