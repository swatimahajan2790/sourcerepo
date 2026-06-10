# -*- coding: utf-8 -*-
"""CPython helper – invoked via subprocess by DocLink.py (IronPython).

Usage
-----
Convert a PDF page to PNG:
    python _pdf_helper.py convert <lib_dir> <pdf_path> <output_png> <page_number> <dpi>

Get the number of pages in a PDF:
    python _pdf_helper.py pagecount <lib_dir> <pdf_path>

Exit codes:  0 = success, 1 = error (message on stderr).
On success the *convert* command prints the output PNG path to stdout;
the *pagecount* command prints the integer page count.
"""

import sys
import os


def _setup_lib(lib_dir):
    if lib_dir and os.path.isdir(lib_dir):
        sys.path.insert(0, lib_dir)


def cmd_convert(lib_dir, pdf_path, output_png, page_number, dpi):
    _setup_lib(lib_dir)
    import fitz  # noqa: E402  (PyMuPDF)

    doc = fitz.open(pdf_path)
    page_idx = page_number - 1
    if page_idx < 0 or page_idx >= len(doc):
        doc.close()
        print("ERROR: page {} out of range (PDF has {} pages)".format(
            page_number, len(doc)), file=sys.stderr)
        return 1

    page = doc[page_idx]
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=True)
    pix.save(output_png)
    doc.close()

    print(output_png)
    return 0


def cmd_pagecount(lib_dir, pdf_path):
    _setup_lib(lib_dir)
    import fitz  # noqa: E402

    doc = fitz.open(pdf_path)
    count = len(doc)
    doc.close()
    print(count)
    return 0


def main():
    if len(sys.argv) < 2:
        print("Usage: _pdf_helper.py <convert|pagecount> ...", file=sys.stderr)
        return 1

    cmd = sys.argv[1]

    if cmd == "convert":
        if len(sys.argv) != 7:
            print("convert requires: lib_dir pdf_path output_png page_number dpi",
                  file=sys.stderr)
            return 1
        lib_dir = sys.argv[2]
        pdf_path = sys.argv[3]
        output_png = sys.argv[4]
        page_number = int(sys.argv[5])
        dpi = int(sys.argv[6])
        return cmd_convert(lib_dir, pdf_path, output_png, page_number, dpi)

    elif cmd == "pagecount":
        if len(sys.argv) != 4:
            print("pagecount requires: lib_dir pdf_path", file=sys.stderr)
            return 1
        lib_dir = sys.argv[2]
        pdf_path = sys.argv[3]
        return cmd_pagecount(lib_dir, pdf_path)

    else:
        print("Unknown command: " + cmd, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main() or 0)
