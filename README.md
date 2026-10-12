# combinepdf

Python GUI that lets users drag and drop PDF files, arrange them, and merge them into a single PDF.

## Features

- Drag and drop PDFs into the list (or arrange manually)
- Move Up / Move Down / Remove to reorder before merging
- Merge PDFs into a single output file
- **Print as picture** option: when checked, every page of every input PDF is rendered to a JPEG image first (like printing to an image and back to PDF) before being placed into the merged output. This flattens text, form fields, annotations, and layers into a single picture per page — useful when a source PDF is malformed, renders fonts oddly, or you just want a flattened, non-editable copy. Leave it unchecked for a normal, lossless merge that keeps each PDF's pages exactly as they are.
  - An **Image quality** dropdown controls the DPI and JPEG quality used when flattening (Low/Medium/High/Very high). It is greyed out unless "Print as picture" is ticked. Lower settings give a much smaller output file at the cost of sharpness — Medium (150 DPI) is a good default; use Low for small files or Very high only when fine detail matters.

## Requirements

Installed automatically on first run (when not bundled as an executable):

- `tkinterdnd2` — drag and drop support
- `PyPDF2` — normal (non-flattened) merging
- `PyMuPDF` (imported as `fitz`) — page rendering for the "Print as picture" option

If automatic installation fails, an error dialog shows the reason and the manual command, e.g. `python -m pip install tkinterdnd2`.

## Usage

1. Run `combinepdf.pyw`.
2. Drag PDF files into the list, or reorder/remove as needed.
3. Check "Print as picture" if you want the merged pages flattened to images, and choose an image quality.
4. Click "Merge PDFs" and choose where to save the result.

## Changelog

- Fixed filename with space issue.
- Added splash screen for PyInstaller build.
- Added "Print as picture" option to flatten pages to images before merging.
- Flattened pages are now JPEG-compressed, with an Image quality dropdown (greyed out unless Print as picture is ticked).
- Fixed first-run auto-install failing with "tkinterdnd2 not found": pip now runs with `python.exe` instead of `pythonw.exe`, falls back to `--user`, refreshes import paths after installing, and reports any failure in a dialog instead of exiting silently.
