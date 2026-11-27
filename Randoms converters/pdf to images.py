from pathlib import Path
import pymupdf


def pdf_to_imgs(pdf, out_dir, fmt="png", dpi=300):
    doc = pymupdf.open(str(pdf))
    mat = pymupdf.Matrix(dpi / 72, dpi / 72)
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        for i in range(doc.page_count):
            pix = doc.load_page(i).get_pixmap(matrix=mat, alpha=False)
            out = out_dir / f"{pdf.stem}_page_{i + 1}.{fmt}"
            pix.save(str(out))
        return doc.page_count
    finally:
        doc.close()


def main():
    print("\n=== PDF TO IMAGES ===")
    print("1. Single file")
    print("2. Entire folder")

    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("PDF file path: ").strip().strip('"'))
        if not src.is_file() or src.suffix.lower() != ".pdf":
            print("File not found.")
            return
        files = [src]

    elif choice == "2":
        src = Path(input("Source folder: ").strip().strip('"'))
        if not src.is_dir():
            print("Folder not found.")
            return
        files = sorted(
            (f for f in src.iterdir() if f.is_file() and f.suffix.lower() == ".pdf"),
            key=lambda f: f.name.lower(),
        )
    else:
        print("Invalid choice.")
        return

    if not files:
        print("No PDF files found.")
        return

    out = Path(input("Output folder: ").strip().strip('"'))
    fmt = input("Format (png/jpg, default png): ").strip().lower() or "png"
    if fmt not in ("png", "jpg"):
        fmt = "png"

    dpi_in = input("DPI (default 300): ").strip()
    dpi = int(dpi_in) if dpi_in.isdigit() else 300

    for f in files:
        try:
            dest = out if len(files) == 1 else out / f.stem
            print("Converting:", f.name)
            n = pdf_to_imgs(f, dest, fmt, dpi)
            print(f"Done: {n} page(s) saved to {dest}")
        except Exception as e:
            print("Failed:", f.name, e)

    print("\nConversion completed!")


if __name__ == "__main__":
    main()
