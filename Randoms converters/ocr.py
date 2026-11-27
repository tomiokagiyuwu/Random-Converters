from pathlib import Path
import argparse
import shutil
import sys
import os
import cv2
import numpy as np
import pymupdf
import pytesseract
from PIL import Image

exts = (".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff")
cache = Path(__file__).parent / ".ocr_cache"
mdl_name = "microsoft/trocr-base-handwritten"

proc = None
mdl = None
dev = None


def set_tess():
    cmd = shutil.which("tesseract")
    if cmd:
        pytesseract.pytesseract.tesseract_cmd = cmd
        return

    paths = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Tesseract-OCR" / "tesseract.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR" / "tesseract.exe",
    ]
    for p in paths:
        if p.is_file():
            pytesseract.pytesseract.tesseract_cmd = str(p)
            return


def load_hw():
    global proc, mdl, dev
    if mdl is not None:
        return

    os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"
    import torch
    from transformers import TrOCRProcessor, VisionEncoderDecoderModel
    from transformers.utils import logging

    logging.set_verbosity_error()
    cache.mkdir(parents=True, exist_ok=True)
    loc = cache / mdl_name.replace("/", "--")
    dev = "cuda" if torch.cuda.is_available() else "cpu"

    if (loc / "config.json").is_file():
        proc = TrOCRProcessor.from_pretrained(str(loc))
        mdl = VisionEncoderDecoderModel.from_pretrained(str(loc))
    else:
        print("Downloading model...")
        proc = TrOCRProcessor.from_pretrained(mdl_name, cache_dir=str(cache))
        mdl = VisionEncoderDecoderModel.from_pretrained(mdl_name, cache_dir=str(cache))
        proc.save_pretrained(str(loc))
        mdl.save_pretrained(str(loc))

    mdl.to(dev)
    mdl.eval()


def prep_img(img):
    arr = np.array(img.convert("RGB"))
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    h, w = gray.shape

    if max(h, w) < 1000:
        gray = cv2.resize(gray, (w * 2, h * 2), interpolation=cv2.INTER_CUBIC)

    blur = cv2.fastNlMeansDenoising(gray, h=10)
    _, th = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return Image.fromarray(th)


def get_lines(img):
    arr = np.array(img.convert("RGB"))
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    h, w = gray.shape

    _, inv = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    msk = cv2.morphologyEx(inv, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8))

    if cv2.countNonZero(msk) == 0:
        return []

    sums = np.sum(msk > 0, axis=1)
    min_px = max(2, int(w * 0.003))
    spans = []
    ins = False
    y0 = 0

    for y, cnt in enumerate(sums):
        if cnt >= min_px and not ins:
            ins = True
            y0 = y
        elif cnt < min_px and ins:
            ins = False
            if y - y0 >= 6:
                spans.append((y0, y))
    if ins and h - y0 >= 6:
        spans.append((y0, h))

    rows = []
    for y1, y2 in spans:
        if rows and y1 - rows[-1][1] <= 4:
            rows[-1] = (rows[-1][0], y2)
        else:
            rows.append((y1, y2))

    crops = []
    for y1, y2 in rows:
        cols = np.where(np.sum(msk[y1:y2, :] > 0, axis=0) > 0)[0]
        if len(cols) == 0 or (cols[-1] - cols[0]) < 8:
            continue
        x1, x2 = int(cols[0]), int(cols[-1]) + 1
        crops.append(img.crop((max(0, x1 - 10), max(0, y1 - 6), min(w, x2 + 10), min(h, y2 + 6))))

    return crops if crops else [img]


def read_hw(img):
    load_hw()
    import torch

    lines = get_lines(img)
    if not lines:
        return ""

    out = []
    with torch.no_grad():
        for i in range(0, len(lines), 8):
            batch = [x.convert("RGB") for x in lines[i : i + 8]]
            px = proc(images=batch, return_tensors="pt").pixel_values.to(dev)
            ids = mdl.generate(px, max_new_tokens=128)
            for t in proc.batch_decode(ids, skip_special_tokens=True):
                if t.strip():
                    out.append(t.strip())

    return "\n".join(out)


def read_prt(img):
    cln = prep_img(img)
    data = pytesseract.image_to_data(cln, output_type=pytesseract.Output.DICT)
    scrs = [float(c) for t, c in zip(data["text"], data["conf"]) if t.strip() and float(c) >= 0]
    txt = pytesseract.image_to_string(cln).strip() or pytesseract.image_to_string(img).strip()
    avg = sum(scrs) / len(scrs) if scrs else 0
    return txt, avg


def read_pg(img, mode):
    if mode == "printed":
        return read_prt(img)[0]
    if mode == "handwritten":
        return read_hw(img)

    txt, avg = read_prt(img)
    if txt and avg >= 82:
        return txt
    return read_hw(img) or txt


def ocr_file(f, mode):
    if f.suffix.lower() != ".pdf":
        with Image.open(f) as img:
            return read_pg(img.convert("RGB"), mode)

    doc = pymupdf.open(str(f))
    mat = pymupdf.Matrix(300 / 72, 300 / 72)
    pgs = []

    try:
        for i in range(doc.page_count):
            print(f"Page {i + 1}/{doc.page_count}: {f.name}")
            pix = doc.load_page(i).get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            txt = read_pg(img, mode)
            pgs.append(f"--- Page {i + 1} ---\n{txt}" if doc.page_count > 1 else txt)
    finally:
        doc.close()

    return "\n\n".join(pgs)


def imgs_to_ocr_pdf(imgs, out):
    set_tess()
    doc = pymupdf.open()
    try:
        for f in imgs:
            print("OCR & adding:", Path(f).name)
            with Image.open(f) as img:
                pdf_bytes = pytesseract.image_to_pdf_or_hocr(img.convert("RGB"), extension="pdf")
            with pymupdf.open("pdf", pdf_bytes) as pg_doc:
                doc.insert_pdf(pg_doc)
        Path(out).parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out))
    finally:
        doc.close()


def make_pdf_searchable(pdf_path):
    set_tess()
    src = pymupdf.open(str(pdf_path))
    out = pymupdf.open()
    mat = pymupdf.Matrix(300 / 72, 300 / 72)
    changed = False

    try:
        for i in range(src.page_count):
            pg = src.load_page(i)
            if len(pg.get_text().strip()) >= 15 and not pg.get_images():
                out.insert_pdf(src, from_page=i, to_page=i)
                continue

            pix = pg.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, extension="pdf")
            with pymupdf.open("pdf", pdf_bytes) as ocr_pg:
                if not pg.get_text().strip():
                    out.insert_pdf(ocr_pg)
                else:
                    out.insert_pdf(src, from_page=i, to_page=i)
                    new_pg = out[-1]
                    for w in ocr_pg[0].get_text("words"):
                        r = pymupdf.Rect(w[:4]) * (72 / 300)
                        new_pg.insert_text(r.bl, w[4], fontsize=max(4, r.height * 0.8), render_mode=3)
            changed = True

        if changed:
            tmp = Path(pdf_path).with_suffix(".tmp.pdf")
            out.save(str(tmp))
    finally:
        src.close()
        out.close()

    if changed:
        tmp.replace(pdf_path)


def run_ocr(src, out_dir, mode="auto"):
    src = Path(src).expanduser().resolve()
    out_dir = Path(out_dir).expanduser().resolve()
    set_tess()

    if src.is_file() and src.suffix.lower() in exts:
        files = [src]
    elif src.is_dir():
        files = sorted(
            (f for f in src.iterdir() if f.is_file() and f.suffix.lower() in exts),
            key=lambda f: f.name.lower(),
        )
    else:
        print("Invalid input path.")
        return False

    if not files:
        print("No valid files found.")
        return False

    out_dir.mkdir(parents=True, exist_ok=True)
    ok = 0

    for f in files:
        try:
            print("Processing:", f.name)
            txt = ocr_file(f, mode)
            out = out_dir / f"{f.stem}.txt"
            out.write_text(txt + "\n", encoding="utf-8")
            print("Done:", out.name)
            ok += 1
        except Exception as e:
            print("Failed:", f.name, e)

    print("OCR completed!")
    return ok > 0


def main():
    if len(sys.argv) > 1:
        p = argparse.ArgumentParser()
        p.add_argument("input")
        p.add_argument("--mode", choices=["printed", "handwritten", "auto"], default="auto")
        p.add_argument("--output", "-o", default="output")
        args = p.parse_args()
        if not run_ocr(args.input, args.output, args.mode):
            sys.exit(1)
        return

    print("\n=== OCR (PDF / IMAGES TO TEXT) ===")
    print("1. Single file")
    print("2. Entire folder")
    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("File path: ").strip().strip('"'))
    elif choice == "2":
        src = Path(input("Source folder: ").strip().strip('"'))
    else:
        print("Invalid choice.")
        return

    out = Path(input("Output folder: ").strip().strip('"'))
    print("\n1. Auto\n2. Printed\n3. Handwritten")
    m = {"1": "auto", "2": "printed", "3": "handwritten"}.get(input("Mode (1/2/3): ").strip(), "auto")
    run_ocr(src, out, m)


if __name__ == "__main__":
    main()
