from pathlib import Path
import io
import pymupdf
from PIL import Image


def shrink_pdf(src, dst, dpi=150, quality=75):
    doc = pymupdf.open(str(src))
    seen = set()

    try:
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

        dst.parent.mkdir(parents=True, exist_ok=True)
        same = src.resolve() == dst.resolve()
        tmp = dst.with_suffix(".tmp.pdf") if same else dst
        doc.save(str(tmp), garbage=4, deflate=True, clean=True)
    finally:
        doc.close()

    if same:
        tmp.replace(dst)


def main():
    print("\n=== COMPRESS PDF ===")
    print("1. Single file")
    print("2. Entire folder")

    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("PDF file path: ").strip().strip('"'))
        if not src.is_file() or src.suffix.lower() != ".pdf":
            print("File not found.")
            return
        out = Path(input("Output PDF path: ").strip().strip('"'))
        if out.suffix.lower() != ".pdf":
            out = out / f"{src.stem}_compressed.pdf" if out.is_dir() else out.with_suffix(".pdf")
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

    print("\n1. Light")
    print("2. Medium (default)")
    print("3. High")
    lvl = input("Level (1/2/3): ").strip() or "2"

    dpi = None if lvl == "1" else (100 if lvl == "3" else 150)
    quality = 55 if lvl == "3" else 75

    for src_file, dst_file in files:
        try:
            old_kb = src_file.stat().st_size / 1024
            print("Compressing:", src_file.name)
            shrink_pdf(src_file, dst_file, dpi, quality)
            new_kb = dst_file.stat().st_size / 1024
            print(f"Done: {dst_file.name} ({old_kb:.0f} KB -> {new_kb:.0f} KB)")
        except Exception as e:
            print("Failed:", src_file.name, e)

    print("\nCompression completed!")


if __name__ == "__main__":
    main()
