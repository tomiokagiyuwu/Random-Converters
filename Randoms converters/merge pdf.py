from pathlib import Path
from pypdf import PdfWriter
from ocr import make_pdf_searchable


def main():
    print("\n=== PDF MERGER (WITH OCR) ===")
    print("1. Merge selected PDF files")
    print("2. Merge all PDFs in a folder")

    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        raw = input("Enter PDF file paths separated by commas:\n").strip()
        files = [Path(p.strip().strip('"')) for p in raw.split(",")]

        if not all(f.is_file() and f.suffix.lower() == ".pdf" for f in files):
            print("One or more PDF files are invalid.")
            return

        out = Path(input("Output PDF path: ").strip().strip('"'))
        if out.suffix.lower() != ".pdf":
            out = out.with_suffix(".pdf")

    elif choice == "2":
        src = Path(input("Source folder: ").strip().strip('"'))
        out_dir = Path(input("Output folder: ").strip().strip('"'))

        if not src.is_dir():
            print("Source folder not found.")
            return

        out_dir.mkdir(parents=True, exist_ok=True)
        out = out_dir / "merged.pdf"

        files = sorted(
            (
                f for f in src.iterdir()
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
        print("No PDF files found.")
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
        print("Merge failed:", e)
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


if __name__ == "__main__":
    main()
