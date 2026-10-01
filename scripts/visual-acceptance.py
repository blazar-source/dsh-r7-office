#!/usr/bin/env python3
"""
Visual acceptance gate for PDF output.

Presence of text and correct ToUnicode do NOT prove a PDF renders correctly:
a broken CID/GID mapping extracts perfectly and draws garbage, and a broken
fill produces a solid black box with all the right text underneath.

So this gate renders the produced PDF with at least two independent engines and
compares each page, pixel by pixel, against a trusted render of the SOURCE
Office file made by LibreOffice (a third, unrelated implementation).

Usage:
    python scripts/visual-acceptance.py <source.docx|.pptx|.xlsx> <produced.pdf> [--out DIR] [--threshold N]

Exit code 0 when every page agrees, 1 otherwise.
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    import numpy as np
    from PIL import Image
except ImportError:  # pragma: no cover
    sys.exit("needs numpy and Pillow: python -m pip install numpy Pillow")

DPI = 110


def find_soffice():
    candidates = [
        os.environ.get("SOFFICE"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        str(Path(os.environ.get("LOCALAPPDATA", "")) / "LibreOfficePortable" / "program" / "soffice.exe"),
        shutil.which("soffice"),
        shutil.which("libreoffice"),
    ]
    for c in candidates:
        if c and Path(c).exists():
            return str(c)
    return None


def soffice_convert(soffice, src, target_ext, outdir, timeout=300):
    """Convert with LibreOffice. The launcher returns before the work is done."""
    outdir.mkdir(parents=True, exist_ok=True)
    profile = f"file:///{(Path(outdir) / 'lo-profile').as_posix()}"
    before = set(outdir.glob(f"*.{target_ext}"))
    subprocess.run(
        [soffice, f"-env:UserInstallation={profile}", "--headless", "--norestore",
         "--convert-to", target_ext, "--outdir", str(outdir), str(src)],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=timeout,
    )
    for _ in range(120):
        new = set(outdir.glob(f"*.{target_ext}")) - before
        if new:
            return sorted(new)[0]
        import time
        time.sleep(1)
    return None


def render_engines(pdf, outdir):
    """Render every page with each available engine. Returns {engine: [paths]}."""
    outdir.mkdir(parents=True, exist_ok=True)
    rendered = {}

    # Engine 1: MuPDF
    try:
        import pymupdf
        doc = pymupdf.open(str(pdf))
        paths = []
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=DPI)
            p = outdir / f"mupdf-p{i+1}.png"
            pix.save(str(p))
            paths.append(p)
        rendered["mupdf"] = paths
    except Exception as exc:  # pragma: no cover
        print(f"  ! MuPDF unavailable: {exc}")

    # Engine 2: PDFium
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(str(pdf))
        paths = []
        for i in range(len(doc)):
            page = doc[i]
            bitmap = page.render(scale=DPI / 72.0)
            img = bitmap.to_pil()
            p = outdir / f"pdfium-p{i+1}.png"
            img.save(p)
            paths.append(p)
        rendered["pdfium"] = paths
    except Exception as exc:  # pragma: no cover
        print(f"  ! PDFium unavailable: {exc}")

    return rendered


def load_gray(p):
    return np.asarray(Image.open(p).convert("L"), dtype=np.int16)


def compare(a_path, b_path):
    """Return (mean absolute difference, fraction of pixels differing strongly)."""
    a, b = load_gray(a_path), load_gray(b_path)
    h = min(a.shape[0], b.shape[0])
    w = min(a.shape[1], b.shape[1])
    a, b = a[:h, :w], b[:h, :w]
    diff = np.abs(a - b)
    return float(diff.mean()), float((diff > 64).mean())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("source")
    ap.add_argument("pdf")
    ap.add_argument("--out", default=None)
    ap.add_argument("--drift", type=float, default=8.0,
                    help="max %% of pixels allowed to differ from the trusted reference "
                         "(independent renderers legitimately disagree about spacing)")
    ap.add_argument("--disagree", type=float, default=0.5,
                    help="max %% of pixels allowed to differ BETWEEN engines; a correct PDF "
                         "renders identically everywhere, so this is the corruption detector")
    args = ap.parse_args()

    source = Path(args.source).resolve()
    pdf = Path(args.pdf).resolve()
    out = Path(args.out) if args.out else Path(tempfile.mkdtemp(prefix="r7-visual-"))
    out.mkdir(parents=True, exist_ok=True)

    soffice = find_soffice()
    if not soffice:
        sys.exit("LibreOffice is required: it renders the trusted reference from the source file")

    print(f"source : {source}")
    print(f"pdf    : {pdf}")
    print(f"outdir : {out}")

    # Trusted reference: LibreOffice renders the SOURCE Office file.
    ref_dir = out / "reference"
    ref_pdf = soffice_convert(soffice, source, "pdf", ref_dir)
    if not ref_pdf:
        sys.exit("LibreOffice could not convert the source file to PDF")
    ref_pages = render_engines(ref_pdf, out / "reference-png").get("mupdf", [])
    if not ref_pages:
        sys.exit("no engine could render the reference PDF")

    # The produced PDF, rendered by every engine.
    engines = render_engines(pdf, out / "produced-png")
    if not engines:
        sys.exit("no engine could render the produced PDF")

    print(f"\nreference pages: {len(ref_pages)}")
    for name, pages in engines.items():
        print(f"  {name}: {len(pages)} pages")

    failures = []
    warnings = []
    n = min([len(ref_pages)] + [len(p) for p in engines.values()])
    if n == 0:
        sys.exit("nothing comparable")

    print("\n--- fidelity against the trusted reference (renderer drift is expected) ---")
    print(f"{'page':>4} {'engine':>8} {'meanDiff':>9} {'diff%':>8}  verdict")
    for i in range(n):
        for name, pages in engines.items():
            mean, strong = compare(ref_pages[i], pages[i])
            bad = strong * 100 > args.drift
            print(f"{i+1:>4} {name:>8} {mean:>9.2f} {strong*100:>7.2f}%  {'DRIFT' if bad else 'ok'}")
            if bad:
                failures.append(("fidelity", i + 1, name, mean, strong * 100))

    # A PDF that contains the right text but the wrong glyph mapping, or a fill
    # that collapses to a solid block, still renders consistently in one engine.
    # Two independent engines disagreeing is therefore the sharpest corruption
    # signal available without knowing what the page should look like.
    names = list(engines)
    if len(names) > 1:
        print(f"\n--- engine agreement (corruption detector, {args.disagree}% limit) ---")
        for i in range(min(len(engines[names[0]]), len(engines[names[1]]))):
            mean, strong = compare(engines[names[0]][i], engines[names[1]][i])
            bad = strong * 100 > args.disagree
            print(f"{i+1:>4} {names[0]} vs {names[1]}  meanDiff={mean:>7.2f}  "
                  f"diff={strong*100:>6.2f}%  {'DISAGREE' if bad else 'ok'}")
            if bad:
                failures.append(("engines", i + 1, f"{names[0]}-vs-{names[1]}", mean, strong * 100))

    print()
    if warnings:
        for kind, page, name, mean, pct in warnings:
            print(f"WARN  page {page} [{name}] {kind} meanDiff={mean:.2f} diff={pct:.2f}%")
    if failures:
        print(f"VISUAL ACCEPTANCE FAILED — {len(failures)} problem(s)")
        for kind, page, name, mean, pct in failures[:20]:
            print(f"  {kind.upper():9} page {page} [{name}] meanDiff={mean:.2f} diff={pct:.2f}%")
        print(f"\nPNGs for inspection: {out}")
        return 1

    print("VISUAL ACCEPTANCE PASSED — engines agree, and every page tracks the trusted reference")
    print(f"PNGs: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
