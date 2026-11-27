from pathlib import Path
from pypdf import PdfWriter
import win32com.client as win32
from ocr import run_ocr, imgs_to_ocr_pdf, make_pdf_searchable


def get_path(prompt):
    return Path(input(prompt).strip().strip('"')).expanduser().resolve()


def choose_mode():
    print("\n1. Single file")
    print("2. Entire folder")
    return input("Choose (1/2): ").strip()


def get_output_path(src):
    out = get_path("Output PDF path: ")
    if out.suffix.lower() != ".pdf":
        out = out / f"{src.stem}.pdf" if out.is_dir() else out.with_suffix(".pdf")
    out.parent.mkdir(parents=True, exist_ok=True)
    return out


def collect_office_jobs(mode, file_prompt, exts, label):
    if mode == "1":
        src = get_path(file_prompt)
        if not src.is_file() or src.suffix.lower() not in exts:
            print(f"Invalid {label} file.")
            return None
        return [(src, get_output_path(src))]

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
                if f.suffix.lower() in exts and not f.name.startswith("~$")
            ),
            key=lambda f: f.name.lower(),
        )
        return [(f, out_dir / f"{f.stem}.pdf") for f in files]

    print("Invalid choice.")
    return None


def word_to_pdf():
    print("\n=== WORD TO PDF ===")
    jobs = collect_office_jobs(
        choose_mode(), "Word file path: ", (".docx", ".doc"), "Word"
    )
    if jobs is None:
        return
    if not jobs:
        print("No Word files found.")
        return

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
                make_pdf_searchable(out)
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
    jobs = collect_office_jobs(
        choose_mode(), "PPTX file path: ", (".pptx", ".ppt"), "PowerPoint"
    )
    if jobs is None:
        return
    if not jobs:
        print("No PowerPoint files found.")
        return

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
                make_pdf_searchable(out)
                print("Done:", out.name)
            except Exception as e:
                print("Failed:", src.name, e)
            finally:
                if deck:
                    deck.Close()
    finally:
        app.Quit()

    print("PowerPoint conversion completed!")


def images_to_pdf():
    print("\n=== IMAGES TO PDF (WITH OCR) ===")
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
            (
                f for f in src_dir.iterdir()
                if f.is_file() and f.suffix.lower() in exts
            ),
            key=lambda f: f.name.lower(),
        )
        out = out_dir / "images.pdf"

    else:
        print("Invalid choice.")
        return

    if not images:
        print("No supported images found.")
        return

    try:
        imgs_to_ocr_pdf(images, out)
        print(f"Successfully converted {len(images)} images with selectable text!")
        print("Output:", out)
    except Exception as e:
        print("Image conversion failed:", e)


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

    writer = PdfWriter()
    try:
        for pdf in files:
            print("Adding:", pdf.name)
            writer.append(str(pdf))

        out.parent.mkdir(parents=True, exist_ok=True)
        with open(out, "wb") as f:
            writer.write(f)
    except Exception as e:
        print("PDF merging failed:", e)
        return
    finally:
        writer.close()

    try:
        print("Running OCR on scanned pages...")
        make_pdf_searchable(out)
        print(f"\nSuccessfully merged {len(files)} PDFs with selectable text!")
        print("Output:", out)
    except Exception as e:
        print("OCR step failed:", e)


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
    run_ocr(src, out_dir, m)


def main():
    actions = {
        "1": word_to_pdf,
        "2": ppt_to_pdf,
        "3": images_to_pdf,
        "4": merge_pdfs,
        "5": ocr_text,
    }

    while True:
        print("\n" + "=" * 40)
        print("       UNIVERSAL DOCUMENT CONVERTER")
        print("=" * 40)
        print("1. Word to PDF")
        print("2. PowerPoint to PDF")
        print("3. Images to PDF")
        print("4. Merge PDFs")
        print("5. OCR (PDF/Image to Text)")
        print("6. Exit")
        print("=" * 40)

        choice = input("Select operation: ").strip()

        if choice == "6":
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
