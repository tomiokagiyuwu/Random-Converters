from pathlib import Path
import html
import importlib.util
import io
import re
import subprocess
import sys

pkgs = {
    "win32com": "pywin32",
    "pypdf": "pypdf",
    "pymupdf": "pymupdf",
    "PIL": "pillow",
    "pytesseract": "pytesseract",
    "cv2": "opencv-python",
    "numpy": "numpy",
    "torch": "torch",
    "transformers": "transformers",
}


def ensure_pkgs(*mods):
    missing = [pkgs[m] for m in mods if importlib.util.find_spec(m) is None]
    if missing:
        print("Installing missing libraries:", ", ".join(missing))
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])


def get_path(prompt):
    return Path(input(prompt).strip().strip('"')).expanduser().resolve()


def choose_mode():
    print("\n1. Single file")
    print("2. Entire folder")
    return input("Choose (1/2): ").strip()


def get_output_path(src, ext=".pdf", suffix=""):
    out = get_path(f"Output {ext.upper().lstrip('.')} path: ")
    if out.suffix.lower() != ext:
        out = out / f"{src.stem}{suffix}{ext}" if out.is_dir() else out.with_suffix(ext)
    out.parent.mkdir(parents=True, exist_ok=True)
    return out


def collect_jobs(mode, prompt, exts, label, out_ext=".pdf", suffix=""):
    if mode == "1":
        src = get_path(prompt)
        if not src.is_file() or src.suffix.lower() not in exts:
            print(f"Invalid {label} file.")
            return None
        return [(src, get_output_path(src, out_ext, suffix))]

    if mode == "2":
        src_dir = get_path("Source folder: ")
        out_dir = get_path("Output folder: ")

        if not src_dir.is_dir():
            print("Folder not found.")
            return None

        out_dir.mkdir(parents=True, exist_ok=True)
        files = sorted(
            (
                f for f in src_dir.iterdir()
                if f.is_file() and f.suffix.lower() in exts and not f.name.startswith("~$")
            ),
            key=lambda f: f.name.lower(),
        )
        return [(f, out_dir / f"{f.stem}{suffix}{out_ext}") for f in files]

    print("Invalid choice.")
    return None


def parse_pages(spec, total):
    if spec.strip().lower() == "all":
        return list(range(total))
    pages = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            s, e = int(a), int(b)
            step = 1 if e >= s else -1
            for p in range(s, e + step, step):
                if 1 <= p <= total:
                    pages.append(p - 1)
        else:
            p = int(part)
            if 1 <= p <= total:
                pages.append(p - 1)
    return pages


def word_to_pdf():
    print("\n=== WORD TO PDF ===")
    jobs = collect_jobs(choose_mode(), "Word file path: ", (".docx", ".doc"), "Word")
    if not jobs:
        if jobs == []:
            print("No Word files found.")
        return

    ensure_pkgs("win32com")
    import win32com.client as win32

    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0

    try:
        for src, out in jobs:
            doc = None
            try:
                print("Converting:", src.name)
                doc = app.Documents.Open(str(src), ReadOnly=True)
                doc.ExportAsFixedFormat(str(out), 17)
                doc.Close(False)
                doc = None
                print("Done:", out.name)
            except Exception as e:
                print("Failed:", src.name, e)
            finally:
                if doc:
                    doc.Close(False)
    finally:
        app.Quit()

    print("Word conversion completed!")


def ppt_to_pdf():
    print("\n=== POWERPOINT TO PDF ===")
    jobs = collect_jobs(choose_mode(), "PPTX file path: ", (".pptx", ".ppt"), "PowerPoint")
    if not jobs:
        if jobs == []:
            print("No PowerPoint files found.")
        return

    ensure_pkgs("win32com")
    import win32com.client as win32

    app = win32.DispatchEx("PowerPoint.Application")

    try:
        for src, out in jobs:
            deck = None
            try:
                print("Converting:", src.name)
                deck = app.Presentations.Open(
                    str(src), ReadOnly=True, Untitled=False, WithWindow=False
                )
                deck.SaveAs(str(out), 32)
                deck.Close()
                deck = None
                print("Done:", out.name)
            except Exception as e:
                print("Failed:", src.name, e)
            finally:
                if deck:
                    deck.Close()
    finally:
        app.Quit()

    print("PowerPoint conversion completed!")


def excel_to_pdf():
    print("\n=== EXCEL TO PDF ===")
    jobs = collect_jobs(choose_mode(), "Excel file path: ", (".xlsx", ".xls", ".csv"), "Excel")
    if not jobs:
        if jobs == []:
            print("No Excel files found.")
        return

    ensure_pkgs("win32com")
    import win32com.client as win32

    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False

    try:
        for src, out in jobs:
            wb = None
            try:
                print("Converting:", src.name)
                wb = app.Workbooks.Open(str(src), ReadOnly=True)
                for ws in wb.Worksheets:
                    try:
                        ws.PageSetup.Zoom = False
                        ws.PageSetup.FitToPagesWide = 1
                        ws.PageSetup.FitToPagesTall = False
                    except Exception:
                        pass
                wb.ExportAsFixedFormat(0, str(out))
                wb.Close(False)
                wb = None
                print("Done:", out.name)
            except Exception as e:
                print("Failed:", src.name, e)
            finally:
                if wb:
                    wb.Close(False)
    finally:
        app.Quit()

    print("Excel conversion completed!")


def images_to_pdf():
    print("\n=== IMAGES TO PDF ===")
    mode = choose_mode()
    exts = (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff")

    if mode == "1":
        src = get_path("Image file path: ")
        if not src.is_file() or src.suffix.lower() not in exts:
            print("Invalid image file.")
            return
        images = [src]
        out = get_output_path(src)

    elif mode == "2":
        src_dir = get_path("Source folder: ")
        out_dir = get_path("Output folder: ")

        if not src_dir.is_dir():
            print("Folder not found.")
            return

        out_dir.mkdir(parents=True, exist_ok=True)
        images = sorted(
            (f for f in src_dir.iterdir() if f.is_file() and f.suffix.lower() in exts),
            key=lambda f: f.name.lower(),
        )
        out = out_dir / "images.pdf"
    else:
        print("Invalid choice.")
        return

    if not images:
        print("No supported images found.")
        return

    ensure_pkgs("PIL")
    from PIL import Image

    try:
        pil_imgs = []
        for f in images:
            with Image.open(f) as img:
                pil_imgs.append(img.convert("RGB"))
        pil_imgs[0].save(str(out), "PDF", resolution=100.0, save_all=True, append_images=pil_imgs[1:])
        print(f"Successfully converted {len(images)} image(s) to PDF!")
        print("Output:", out)
    except Exception as e:
        print("Image conversion failed:", e)


def md_to_html(text):
    out = []
    in_code = False

    for line in text.splitlines():
        s = line.strip()
        if s.startswith("```"):
            out.append("<pre>" if not in_code else "</pre>")
            in_code = not in_code
            continue

        if in_code:
            out.append(html.escape(line))
            continue

        if not s:
            out.append("<br>")
            continue

        line = html.escape(line)
        line = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", line)
        line = re.sub(r"\*(.+?)\*", r"<i>\1</i>", line)
        line = re.sub(r"`(.+?)`", r"<code>\1</code>", line)

        if s.startswith("# "):
            out.append(f"<h1>{line.lstrip('# ').strip()}</h1>")
        elif s.startswith("## "):
            out.append(f"<h2>{line.lstrip('# ').strip()}</h2>")
        elif s.startswith("### "):
            out.append(f"<h3>{line.lstrip('# ').strip()}</h3>")
        elif s.startswith(("- ", "* ")):
            out.append(f"<p>• {line.strip()[2:]}</p>")
        else:
            out.append(f"<p>{line}</p>")

    if in_code:
        out.append("</pre>")
    return "\n".join(out)


def text_to_pdf():
    print("\n=== TEXT / MARKDOWN / HTML TO PDF ===")
    exts = (".txt", ".md", ".html", ".htm")
    jobs = collect_jobs(choose_mode(), "File path: ", exts, "text/markdown/html")
    if not jobs:
        if jobs == []:
            print("No matching files found.")
        return

    ensure_pkgs("pymupdf")
    import pymupdf

    for src, out in jobs:
        try:
            print("Converting:", src.name)
            raw = src.read_text(encoding="utf-8", errors="replace")
            ext = src.suffix.lower()

            if ext == ".txt":
                body = f"<pre>{html.escape(raw)}</pre>"
            elif ext == ".md":
                body = md_to_html(raw)
            else:
                body = raw

            content = f"<html><body style='font-family:sans-serif;font-size:11pt;'>{body}</body></html>"
            story = pymupdf.Story(html=content)

            out.parent.mkdir(parents=True, exist_ok=True)
            writer = pymupdf.DocumentWriter(str(out))
            box = pymupdf.paper_rect("a4")
            rect = box + (36, 36, -36, -36)

            more = True
            while more:
                dev = writer.begin_page(box)
                more, _ = story.place(rect)
                story.draw(dev)
                writer.end_page()
            writer.close()
            print("Done:", out.name)
        except Exception as e:
            print("Failed:", src.name, e)

    print("Conversion completed!")


def pdf_to_images():
    print("\n=== PDF TO IMAGES ===")
    mode = choose_mode()

    if mode == "1":
        src = get_path("PDF file path: ")
        if not src.is_file() or src.suffix.lower() != ".pdf":
            print("Invalid PDF file.")
            return
        files = [src]
    elif mode == "2":
        src_dir = get_path("Source folder: ")
        if not src_dir.is_dir():
            print("Folder not found.")
            return
        files = sorted(
            (f for f in src_dir.iterdir() if f.is_file() and f.suffix.lower() == ".pdf"),
            key=lambda f: f.name.lower(),
        )
    else:
        print("Invalid choice.")
        return

    if not files:
        print("No PDF files found.")
        return

    out_dir = get_path("Output folder: ")
    fmt = input("Format (png/jpg, default png): ").strip().lower() or "png"
    if fmt not in ("png", "jpg"):
        fmt = "png"

    dpi_in = input("DPI (default 300): ").strip()
    dpi = int(dpi_in) if dpi_in.isdigit() else 300

    ensure_pkgs("pymupdf")
    import pymupdf

    mat = pymupdf.Matrix(dpi / 72, dpi / 72)

    for f in files:
        doc = None
        try:
            dest = out_dir if len(files) == 1 else out_dir / f.stem
            dest.mkdir(parents=True, exist_ok=True)
            print("Converting:", f.name)
            doc = pymupdf.open(str(f))
            for i in range(doc.page_count):
                pix = doc.load_page(i).get_pixmap(matrix=mat, alpha=False)
                pix.save(str(dest / f"{f.stem}_page_{i + 1}.{fmt}"))
            print(f"Done: {doc.page_count} page(s) saved to {dest}")
        except Exception as e:
            print("Failed:", f.name, e)
        finally:
            if doc:
                doc.close()

    print("Conversion completed!")


def pdf_to_word():
    print("\n=== PDF TO WORD ===")
    jobs = collect_jobs(choose_mode(), "PDF file path: ", (".pdf",), "PDF", ".docx")
    if not jobs:
        if jobs == []:
            print("No PDF files found.")
        return

    ensure_pkgs("win32com")
    import win32com.client as win32

    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0

    try:
        for src, out in jobs:
            doc = None
            try:
                print("Converting:", src.name)
                doc = app.Documents.Open(str(src), ConfirmConversions=False, ReadOnly=True)
                doc.SaveAs2(str(out), FileFormat=16)
                doc.Close(False)
                doc = None
                print("Done:", out.name)
            except Exception as e:
                print("Failed:", src.name, e)
            finally:
                if doc:
                    doc.Close(False)
    finally:
        app.Quit()

    print("Conversion completed!")


def image_converter():
    print("\n=== IMAGE CONVERTER & COMPRESSOR ===")
    exts = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff")
    fmts = {
        "jpg": ("JPEG", ".jpg"),
        "jpeg": ("JPEG", ".jpg"),
        "png": ("PNG", ".png"),
        "webp": ("WEBP", ".webp"),
        "bmp": ("BMP", ".bmp"),
        "tiff": ("TIFF", ".tiff"),
        "tif": ("TIFF", ".tif"),
    }

    mode = choose_mode()
    if mode == "1":
        src = get_path("Image file path: ")
        if not src.is_file() or src.suffix.lower() not in exts:
            print("Invalid image file.")
            return
        files = [src]
    elif mode == "2":
        src_dir = get_path("Source folder: ")
        if not src_dir.is_dir():
            print("Folder not found.")
            return
        files = sorted(
            (f for f in src_dir.iterdir() if f.is_file() and f.suffix.lower() in exts),
            key=lambda f: f.name.lower(),
        )
    else:
        print("Invalid choice.")
        return

    if not files:
        print("No images found.")
        return

    out_dir = get_path("Output folder: ")
    target = input("Target format (jpg/png/webp/bmp/tiff, Enter to keep): ").strip().lower()
    if target and target not in fmts:
        print("Invalid format.")
        return

    q = input("Quality 1-100 (default 85): ").strip()
    quality = int(q) if q.isdigit() and 1 <= int(q) <= 100 else 85

    sz = input("Max width/height in px (Enter to skip): ").strip()
    max_size = int(sz) if sz.isdigit() and int(sz) > 0 else None

    ensure_pkgs("PIL")
    from PIL import Image

    out_dir.mkdir(parents=True, exist_ok=True)

    for f in files:
        try:
            key = target if target else f.suffix.lower().lstrip(".")
            fmt, ext = fmts[key]
            dst = out_dir / f"{f.stem}{ext}"

            with Image.open(f) as img:
                if max_size and max(img.size) > max_size:
                    img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)

                if fmt in ("JPEG", "BMP") and img.mode in ("RGBA", "LA", "P"):
                    bg = Image.new("RGB", img.size, (255, 255, 255))
                    rgba = img.convert("RGBA")
                    bg.paste(rgba, mask=rgba.split()[3])
                    img = bg
                elif fmt == "JPEG" and img.mode != "RGB":
                    img = img.convert("RGB")

                args = {"optimize": True}
                if fmt in ("JPEG", "WEBP"):
                    args["quality"] = quality

                img.save(str(dst), format=fmt, **args)

            print("Done:", dst.name)
        except Exception as e:
            print("Failed:", f.name, e)

    print("Conversion completed!")


def ocr_text():
    print("\n=== OCR (PDF / IMAGES TO TEXT) ===")
    mode = choose_mode()

    if mode == "1":
        src = get_path("File path: ")
    elif mode == "2":
        src = get_path("Source folder: ")
    else:
        print("Invalid choice.")
        return

    out_dir = get_path("Output folder: ")

    print("\n1. Auto")
    print("2. Printed")
    print("3. Handwritten")
    m = {"1": "auto", "2": "printed", "3": "handwritten"}.get(
        input("Choose (1/2/3): ").strip(), "auto"
    )
    if m == "printed":
        ensure_pkgs("pymupdf", "PIL", "pytesseract", "cv2", "numpy")
    else:
        ensure_pkgs("pymupdf", "PIL", "pytesseract", "cv2", "numpy", "torch", "transformers")
    from ocr import run_ocr

    run_ocr(src, out_dir, m)


def merge_pdfs():
    print("\n=== MERGE PDFs ===")
    print("1. Select PDF files")
    print("2. Merge entire folder")

    mode = input("Choose (1/2): ").strip()

    if mode == "1":
        raw = input("Enter PDF paths separated by commas:\n")
        files = [Path(p.strip().strip('"')).resolve() for p in raw.split(",")]

        if not all(f.is_file() and f.suffix.lower() == ".pdf" for f in files):
            print("Invalid PDF file(s).")
            return

        out = get_path("Output PDF path: ")
        if out.suffix.lower() != ".pdf":
            out = out.with_suffix(".pdf")

    elif mode == "2":
        src_dir = get_path("Source folder: ")
        out_dir = get_path("Output folder: ")

        if not src_dir.is_dir():
            print("Folder not found.")
            return

        out_dir.mkdir(parents=True, exist_ok=True)
        out = out_dir / "merged.pdf"

        files = sorted(
            (
                f for f in src_dir.iterdir()
                if f.is_file()
                and f.suffix.lower() == ".pdf"
                and f.resolve() != out.resolve()
            ),
            key=lambda f: f.name.lower(),
        )
    else:
        print("Invalid choice.")
        return

    if not files:
        print("No PDFs found.")
        return

    ensure_pkgs("pypdf")
    from pypdf import PdfWriter

    writer = PdfWriter()
    try:
        for pdf in files:
            print("Adding:", pdf.name)
            writer.append(str(pdf))

        out.parent.mkdir(parents=True, exist_ok=True)
        with open(out, "wb") as f:
            writer.write(f)
        print(f"\nSuccessfully merged {len(files)} PDFs!")
        print("Output:", out)
    except Exception as e:
        print("PDF merging failed:", e)
    finally:
        writer.close()


def split_pdf():
    print("\n=== SPLIT / EXTRACT PDF ===")
    src = get_path("PDF file path: ")
    if not src.is_file() or src.suffix.lower() != ".pdf":
        print("Invalid PDF file.")
        return

    print("\n1. Split into single-page PDFs")
    print("2. Extract page range (e.g. 1-3, 5, 8-10)")
    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        out_dir = get_path("Output folder: ")
    elif choice == "2":
        out = get_output_path(src, ".pdf", "_extracted")
    else:
        print("Invalid choice.")
        return

    ensure_pkgs("pypdf")
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(str(src))
    total = len(reader.pages)
    print(f"Total pages: {total}")

    try:
        if choice == "1":
            out_dir.mkdir(parents=True, exist_ok=True)

            for i, page in enumerate(reader.pages, 1):
                writer = PdfWriter()
                writer.add_page(page)
                dst = out_dir / f"{src.stem}_page_{i}.pdf"
                with open(dst, "wb") as f:
                    writer.write(f)
                writer.close()

            print(f"\nSplit into {total} files in {out_dir}")

        elif choice == "2":
            spec = input(f"Pages (1-{total}): ").strip()
            pages = parse_pages(spec, total)
            if not pages:
                print("No valid pages selected.")
                return

            writer = PdfWriter()
            for idx in pages:
                writer.add_page(reader.pages[idx])

            with open(out, "wb") as f:
                writer.write(f)
            writer.close()

            print(f"\nExtracted {len(pages)} page(s) to {out}")
    except Exception as e:
        print("Failed:", e)


def compress_pdf():
    print("\n=== COMPRESS PDF ===")
    jobs = collect_jobs(choose_mode(), "PDF file path: ", (".pdf",), "PDF", ".pdf", "_compressed")
    if not jobs:
        if jobs == []:
            print("No PDF files found.")
        return

    print("\n1. Light")
    print("2. Medium (default)")
    print("3. High")
    lvl = input("Level (1/2/3): ").strip() or "2"

    dpi = None if lvl == "1" else (100 if lvl == "3" else 150)
    quality = 55 if lvl == "3" else 75

    if dpi:
        ensure_pkgs("pymupdf", "PIL")
        from PIL import Image
    else:
        ensure_pkgs("pymupdf")
    import pymupdf

    for src, out in jobs:
        doc = None
        try:
            old_kb = src.stat().st_size / 1024
            print("Compressing:", src.name)
            doc = pymupdf.open(str(src))
            seen = set()

            if dpi:
                max_px = int(dpi * 11.7)
                for pg in doc:
                    for info in pg.get_images(full=True):
                        xref = info[0]
                        if xref in seen:
                            continue
                        seen.add(xref)
                        try:
                            raw = doc.extract_image(xref)
                            if not raw or not raw.get("image"):
                                continue
                            with Image.open(io.BytesIO(raw["image"])) as img:
                                if img.mode in ("RGBA", "LA", "P"):
                                    continue
                                img = img.convert("RGB")
                                if max(img.size) > max_px:
                                    img.thumbnail((max_px, max_px), Image.Resampling.LANCZOS)
                                buf = io.BytesIO()
                                img.save(buf, format="JPEG", quality=quality, optimize=True)
                                data = buf.getvalue()
                                if len(data) < len(raw["image"]):
                                    pg.replace_image(xref, stream=data)
                        except Exception:
                            pass

            same = src.resolve() == out.resolve()
            tmp = out.with_suffix(".tmp.pdf") if same else out
            doc.save(str(tmp), garbage=4, deflate=True, clean=True)
            doc.close()
            doc = None

            if same:
                tmp.replace(out)

            new_kb = out.stat().st_size / 1024
            print(f"Done: {out.name} ({old_kb:.0f} KB -> {new_kb:.0f} KB)")
        except Exception as e:
            print("Failed:", src.name, e)
        finally:
            if doc:
                doc.close()

    print("Compression completed!")


def protect_pdf():
    print("\n=== ENCRYPT / DECRYPT PDF ===")
    print("1. Encrypt (Add password)")
    print("2. Decrypt (Remove password)")
    act = input("Choose (1/2): ").strip()
    if act not in ("1", "2"):
        print("Invalid choice.")
        return

    jobs = collect_jobs(choose_mode(), "PDF file path: ", (".pdf",), "PDF")
    if not jobs:
        if jobs == []:
            print("No PDF files found.")
        return

    pw = input("Password: ").strip()
    if not pw:
        print("Password cannot be empty.")
        return

    ensure_pkgs("pymupdf")
    import pymupdf

    for src, out in jobs:
        doc = None
        try:
            doc = pymupdf.open(str(src))
            same = src.resolve() == out.resolve()
            tmp = out.with_suffix(".tmp.pdf") if same else out

            if act == "1":
                doc.save(
                    str(tmp),
                    encryption=pymupdf.PDF_ENCRYPT_AES_256,
                    user_pw=pw,
                    owner_pw=pw,
                )
                label = "Encrypted:"
            else:
                if doc.is_encrypted and not doc.authenticate(pw):
                    raise ValueError("Wrong password.")
                doc.save(str(tmp), encryption=pymupdf.PDF_ENCRYPT_NONE)
                label = "Decrypted:"

            doc.close()
            doc = None
            if same:
                tmp.replace(out)
            print(label, out.name)
        except Exception as e:
            print("Failed:", src.name, e)
        finally:
            if doc:
                doc.close()

    print("Done!")


def is_blank(pg):
    import numpy as np
    import pymupdf

    if pg.get_text().strip():
        return False
    pix = pg.get_pixmap(matrix=pymupdf.Matrix(1, 1), colorspace=pymupdf.csGRAY, alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8)
    return float(np.std(arr)) < 3.5 and float(np.mean(arr)) > 245.0


def edit_pages():
    print("\n=== EDIT PDF PAGES ===")
    src = get_path("PDF file path: ")
    if not src.is_file() or src.suffix.lower() != ".pdf":
        print("Invalid PDF file.")
        return

    print("\n1. Rotate pages")
    print("2. Reorder pages")
    print("3. Delete pages")
    print("4. Remove blank pages")
    act = input("Choose (1/2/3/4): ").strip()
    if act not in ("1", "2", "3", "4"):
        print("Invalid choice.")
        return

    out = get_output_path(src, ".pdf", "_edited")
    same = src.resolve() == out.resolve()
    tmp = out.with_suffix(".tmp.pdf") if same else out

    if act == "4":
        ensure_pkgs("pymupdf", "numpy")
    else:
        ensure_pkgs("pymupdf")
    import pymupdf

    doc = pymupdf.open(str(src))
    total = doc.page_count
    print(f"Total pages: {total}")

    try:
        if act == "1":
            spec = input("Pages to rotate (e.g. 1-3, 5 or all): ").strip() or "all"
            angle = int(input("Angle (90 / 180 / 270): ").strip())
            for i in parse_pages(spec, total):
                doc[i].set_rotation((doc[i].rotation + angle) % 360)

        elif act == "2":
            spec = input(f"New page order (e.g. 3, 1-2, 4-{total}): ").strip()
            order = parse_pages(spec, total)
            if not order:
                print("Invalid pages.")
                return
            doc.select(order)

        elif act == "3":
            spec = input("Pages to delete (e.g. 2, 4-5): ").strip()
            drop = set(parse_pages(spec, total))
            keep = [i for i in range(total) if i not in drop]
            if not keep:
                print("Cannot delete all pages.")
                return
            doc.select(keep)

        elif act == "4":
            keep = [i for i in range(total) if not is_blank(doc[i])]
            if not keep:
                keep = [0]
            print(f"Removed {total - len(keep)} blank page(s).")
            doc.select(keep)

        else:
            print("Invalid choice.")
            return

        doc.save(str(tmp), garbage=3, deflate=True)
    except Exception as e:
        print("Failed:", e)
        return
    finally:
        doc.close()

    if same:
        tmp.replace(out)
    print("Saved:", out)


def watermark_pdf():
    print("\n=== WATERMARK & PAGE NUMBERS ===")
    print("1. Watermark only")
    print("2. Page numbers only")
    print("3. Both")
    mode = input("Choose (1/2/3): ").strip()
    if mode not in ("1", "2", "3"):
        print("Invalid choice.")
        return

    jobs = collect_jobs(choose_mode(), "PDF file path: ", (".pdf",), "PDF", ".pdf", "_stamped")
    if not jobs:
        if jobs == []:
            print("No PDF files found.")
        return

    text = ""
    if mode in ("1", "3"):
        text = input("Watermark text: ").strip()
        if not text:
            print("Watermark text cannot be empty.")
            return

    nums = mode in ("2", "3")
    pos = "right"
    if nums:
        print("\n1. Bottom Right")
        print("2. Bottom Center")
        print("3. Bottom Left")
        p = input("Position (1/2/3): ").strip()
        pos = {"1": "right", "2": "center", "3": "left"}.get(p, "right")

    ensure_pkgs("pymupdf")
    import pymupdf

    for src, out in jobs:
        doc = None
        try:
            print("Stamping:", src.name)
            doc = pymupdf.open(str(src))
            total = doc.page_count
            same = src.resolve() == out.resolve()
            tmp = out.with_suffix(".tmp.pdf") if same else out

            for i, pg in enumerate(doc, 1):
                r = pg.rect

                if text:
                    sz = max(24, min(r.width, r.height) * 0.07)
                    w = pymupdf.get_text_length(text, fontname="helv", fontsize=sz)
                    mid = pymupdf.Point(r.width / 2, r.height / 2)
                    pt = pymupdf.Point(mid.x - w / 2, mid.y + sz / 3)
                    pg.insert_text(
                        pt,
                        text,
                        fontname="helv",
                        fontsize=sz,
                        color=(0.75, 0.75, 0.75),
                        morph=(mid, pymupdf.Matrix(45)),
                        overlay=True,
                    )

                if nums:
                    lbl = f"Page {i} of {total}"
                    w = pymupdf.get_text_length(lbl, fontname="helv", fontsize=10)
                    if pos == "left":
                        x = 36
                    elif pos == "center":
                        x = (r.width - w) / 2
                    else:
                        x = r.width - w - 36

                    pg.insert_text(
                        pymupdf.Point(x, r.height - 20),
                        lbl,
                        fontname="helv",
                        fontsize=10,
                        color=(0.25, 0.25, 0.25),
                        overlay=True,
                    )

            doc.save(str(tmp), garbage=2, deflate=True)
            doc.close()
            doc = None

            if same:
                tmp.replace(out)
            print("Done:", out.name)
        except Exception as e:
            print("Failed:", src.name, e)
        finally:
            if doc:
                doc.close()

    print("Completed!")


def main():
    actions = {
        "1": word_to_pdf,
        "2": ppt_to_pdf,
        "3": excel_to_pdf,
        "4": images_to_pdf,
        "5": text_to_pdf,
        "6": pdf_to_images,
        "7": pdf_to_word,
        "8": image_converter,
        "9": ocr_text,
        "10": merge_pdfs,
        "11": split_pdf,
        "12": compress_pdf,
        "13": protect_pdf,
        "14": edit_pages,
        "15": watermark_pdf,
    }

    while True:
        print("\n" + "=" * 40)
        print("       UNIVERSAL DOCUMENT CONVERTER")
        print("=" * 40)
        print("1. Word to PDF")
        print("2. PowerPoint to PDF")
        print("3. Excel to PDF")
        print("4. Images to PDF")
        print("5. Text / Markdown / HTML to PDF")
        print("6. PDF to Images")
        print("7. PDF to Word")
        print("8. Image Converter & Compressor")
        print("9. OCR (PDF/Image to Text)")
        print("10. Merge PDFs")
        print("11. Split / Extract PDF")
        print("12. Compress PDF")
        print("13. Encrypt / Decrypt PDF")
        print("14. Rotate / Reorder / Delete Pages")
        print("15. Watermark & Page Numbers")
        print("16. Exit")
        print("=" * 40)

        choice = input("Select operation: ").strip()

        if choice == "16":
            print("Goodbye!")
            break

        action = actions.get(choice)
        if action:
            action()
        else:
            print("Invalid selection.")

        input("\nPress Enter to return to main menu...")


if __name__ == "__main__":
    main()
