from pathlib import Path
from PIL import Image

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


def convert_img(src, dst, fmt, quality=85, max_size=None):
    with Image.open(src) as img:
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

        dst.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(dst), format=fmt, **args)


def main():
    print("\n=== IMAGE CONVERTER & COMPRESSOR ===")
    print("1. Single image")
    print("2. Entire folder")

    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("Image file path: ").strip().strip('"'))
        if not src.is_file() or src.suffix.lower() not in exts:
            print("File not found.")
            return
        files = [src]

    elif choice == "2":
        src = Path(input("Source folder: ").strip().strip('"'))
        if not src.is_dir():
            print("Folder not found.")
            return
        files = sorted(
            (f for f in src.iterdir() if f.is_file() and f.suffix.lower() in exts),
            key=lambda f: f.name.lower(),
        )
    else:
        print("Invalid choice.")
        return

    if not files:
        print("No images found.")
        return

    out = Path(input("Output folder: ").strip().strip('"'))
    target = input("Target format (jpg/png/webp/bmp/tiff, Enter to keep): ").strip().lower()
    if target and target not in fmts:
        print("Invalid format.")
        return

    q = input("Quality 1-100 (default 85): ").strip()
    quality = int(q) if q.isdigit() and 1 <= int(q) <= 100 else 85

    sz = input("Max width/height in px (Enter to skip): ").strip()
    max_size = int(sz) if sz.isdigit() and int(sz) > 0 else None

    for f in files:
        try:
            key = target if target else f.suffix.lower().lstrip(".")
            fmt, ext = fmts[key]
            dst = out / f"{f.stem}{ext}"
            convert_img(f, dst, fmt, quality, max_size)
            print("Done:", dst.name)
        except Exception as e:
            print("Failed:", f.name, e)

    print("\nConversion completed!")


if __name__ == "__main__":
    main()
