from pathlib import Path
from pypdf import PdfReader, PdfWriter


def parse_ranges(spec, total):
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


def main():
    print("\n=== SPLIT / EXTRACT PDF ===")
    src = Path(input("PDF file path: ").strip().strip('"'))
    if not src.is_file() or src.suffix.lower() != ".pdf":
        print("File not found.")
        return

    reader = PdfReader(str(src))
    total = len(reader.pages)
    print(f"Total pages: {total}")

    print("\n1. Split into single-page PDFs")
    print("2. Extract page range (e.g. 1-3, 5, 8-10)")
    choice = input("Choose (1/2): ").strip()

    try:
        if choice == "1":
            out_dir = Path(input("Output folder: ").strip().strip('"'))
            out_dir.mkdir(parents=True, exist_ok=True)

            for i, page in enumerate(reader.pages, 1):
                writer = PdfWriter()
                writer.add_page(page)
                out = out_dir / f"{src.stem}_page_{i}.pdf"
                with open(out, "wb") as f:
                    writer.write(f)
                writer.close()

            print(f"\nSplit into {total} files in {out_dir}")

        elif choice == "2":
            spec = input(f"Pages (1-{total}): ").strip()
            pages = parse_ranges(spec, total)
            if not pages:
                print("No valid pages selected.")
                return

            out = Path(input("Output PDF path: ").strip().strip('"'))
            if out.suffix.lower() != ".pdf":
                out = out / f"{src.stem}_extracted.pdf" if out.is_dir() else out.with_suffix(".pdf")

            writer = PdfWriter()
            for idx in pages:
                writer.add_page(reader.pages[idx])

            out.parent.mkdir(parents=True, exist_ok=True)
            with open(out, "wb") as f:
                writer.write(f)
            writer.close()

            print(f"\nExtracted {len(pages)} page(s) to {out}")
        else:
            print("Invalid choice.")
    except Exception as e:
        print("Failed:", e)


if __name__ == "__main__":
    main()
