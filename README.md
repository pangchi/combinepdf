# combinepdf

Python GUI that lets users drag and drop PDF files, arrange them, and merge them into a single PDF.

ChatGPT prompt: "python gui to allow user to drag in pdfs and print as one pdf file according to sequence in ui"

## Features

- Drag and drop PDFs into the list (or arrange manually)
- Move Up / Move Down / Remove to reorder before merging
- Merge PDFs into a single output file
- **Print as picture** option: when checked, every page of every input PDF is rendered to an image first (like printing to an image and back to PDF) before being placed into the merged output. This flattens text, form fields, annotations, and layers into a single picture per page — useful when a source PDF is malformed, renders fonts oddly, or you just want a flattened, non-editable copy. Leave it unchecked for a normal, lossless merge that keeps each PDF's pages exactly as they are.

## Requirements

Installed automatically on first run (when not bundled as an executable):

- `tkinterdnd2` — drag and drop support
- `PyPDF2` — normal (non-flattened) merging
- `PyMuPDF` (imported as `fitz`) — page rendering for the "Print as picture" option

## Usage

1. Run `combinepdf.pyw`.
2. Drag PDF files into the list, or reorder/remove as needed.
3. Check "Print as picture" if you want the merged pages flattened to images.
4. Click "Merge PDFs" and choose where to save the result.

## Changelog

- Fixed filename with space issue.
- Added splash screen for PyInstaller build.
- Added "Print as picture" option to flatten pages to images before merging.
