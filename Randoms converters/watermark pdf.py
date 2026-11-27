from pathlib import Path
import pymupdf


def stamp_pdf(src, dst, text="", page_nums=False, pos="right"):
    doc = pymupdf.open(str(src))
    total = doc.page_count
    same = src.resolve() == dst.resolve()
    tmp = dst.with_suffix(".tmp.pdf") if same else dst

    try:
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

            if page_nums:
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

        dst.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(tmp), garbage=2, deflate=True)
    finally:
        doc.close()

    if same:
        tmp.replace(dst)


def main():
    print("\n=== WATERMARK & PAGE NUMBERS ===")
    print("1. Watermark only")
    print("2. Page numbers only")
    print("3. Both")
    mode = input("Choose (1/2/3): ").strip()
    if mode not in ("1", "2", "3"):
        print("Invalid choice.")
        return

    print("\n1. Single file")
    print("2. Entire folder")
    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("PDF file path: ").strip().strip('"'))
        if not src.is_file() or src.suffix.lower() != ".pdf":
            print("File not found.")
            return
        out = Path(input("Output PDF path: ").strip().strip('"'))
        if out.suffix.lower() != ".pdf":
            out = out / f"{src.stem}_stamped.pdf" if out.is_dir() else out.with_suffix(".pdf")
        files = [(src, out)]

    elif choice == "2":
        src = Path(input("Source folder: ").strip().strip('"'))
        out = Path(input("Output folder: ").strip().strip('"'))
        if not src.is_dir():
            print("Folder not found.")
            return
        out.mkdir(parents=True, exist_ok=True)
        files = [(f, out / f.name) for f in src.glob("*.pdf")]
    else:
        print("Invalid choice.")
        return

    if not files:
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

    for src_file, dst_file in files:
        try:
            print("Stamping:", src_file.name)
            stamp_pdf(src_file, dst_file, text, nums, pos)
            print("Done:", dst_file.name)
        except Exception as e:
            print("Failed:", src_file.name, e)

    print("\nCompleted!")


if __name__ == "__main__":
    main()
