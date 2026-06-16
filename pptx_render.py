"""
pptx_render.py — render a .pptx to a PDF (and optionally per-slide PNGs) so a
human / LLM reviewer can actually LOOK at the slides (the JUDGMENT layer).

Why PDF: LibreOffice renders pptx -> pdf faithfully and the Claude Read tool can
open PDF pages as images directly. PNG export is best-effort (needs PyMuPDF).

Usage:
    py pptx_render.py deck.pptx [--outdir DIR] [--png] [--dpi 150]

Prints the produced file paths, one per line, prefixed with their type.
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile

SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "soffice",
    "libreoffice",
]


def find_soffice():
    for c in SOFFICE_CANDIDATES:
        if os.path.isfile(c):
            return c
        found = shutil.which(c)
        if found:
            return found
    return None


def to_pdf(pptx_path, outdir, soffice):
    os.makedirs(outdir, exist_ok=True)
    # isolated profile so it works even if a LibreOffice GUI is open
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    profile_uri = "file:///" + profile.replace("\\", "/")
    cmd = [
        soffice, "--headless", "--norestore",
        "-env:UserInstallation=" + profile_uri,
        "--convert-to", "pdf", "--outdir", outdir, pptx_path,
    ]
    subprocess.run(cmd, check=True, capture_output=True, timeout=180)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    if not os.path.isfile(pdf):
        raise RuntimeError("LibreOffice did not produce a PDF at " + pdf)
    return pdf


def pdf_to_pngs(pdf_path, outdir, dpi):
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return []
    doc = fitz.open(pdf_path)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pngs = []
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=mat)
        out = os.path.join(outdir, f"slide-{i:02d}.png")
        pix.save(out)
        pngs.append(out)
    doc.close()
    return pngs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--outdir", default=None)
    ap.add_argument("--png", action="store_true")
    ap.add_argument("--dpi", type=int, default=150)
    args = ap.parse_args()

    pptx_path = os.path.abspath(args.pptx)
    if not os.path.isfile(pptx_path):
        print("ERROR: no such file:", pptx_path, file=sys.stderr)
        sys.exit(2)
    outdir = os.path.abspath(args.outdir) if args.outdir else os.path.join(
        os.path.dirname(pptx_path), "_render")

    soffice = find_soffice()
    if not soffice:
        print("ERROR: LibreOffice (soffice) not found. Install it to enable rendering.", file=sys.stderr)
        sys.exit(3)

    pdf = to_pdf(pptx_path, outdir, soffice)
    print("PDF\t" + pdf)

    if args.png:
        pngs = pdf_to_pngs(pdf, outdir, args.dpi)
        if pngs:
            for p in pngs:
                print("PNG\t" + p)
        else:
            print("NOTE\tPNG export skipped (PyMuPDF not installed). Read the PDF pages directly.",
                  file=sys.stderr)


if __name__ == "__main__":
    main()
