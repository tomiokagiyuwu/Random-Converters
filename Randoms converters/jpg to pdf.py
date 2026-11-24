from pathlib import Path
import img2pdf


def main():
    print("\n=== IMAGES TO PDF ===")
    print("1. Single image")
    print("2. Entire folder (combine into one PDF)")

    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("Image file path: ").strip().strip('"'))
        if not src.is_file() or src.suffix.lower() not in (".jpg", ".jpeg"):
            print("Valid JPG/JPEG file not found.")
            return

        out = Path(input("Output PDF path: ").strip().strip('"'))
        if out.suffix.lower() != ".pdf":
            out = out / f"{src.stem}.pdf" if out.is_dir() else out.with_suffix(".pdf")

        images = [src]

    elif choice == "2":
        src = Path(input("Source folder: ").strip().strip('"'))
        out = Path(input("Output folder: ").strip().strip('"'))

        if not src.is_dir():
            print("Folder not found.")
            return

        out.mkdir(parents=True, exist_ok=True)
        images = sorted(
            (
                f for f in src.iterdir()
                if f.is_file() and f.suffix.lower() in (".jpg", ".jpeg")
            ),
            key=lambda f: f.name.lower(),
        )
        out = out / "output.pdf"

    else:
        print("Invalid choice.")
        return

    if not images:
        print("No JPG/JPEG images found.")
        return

    try:
        with open(out, "wb") as pdf:
            pdf.write(img2pdf.convert([str(f) for f in images]))

        print(f"\nSuccessfully converted {len(images)} image(s).")
        print("PDF saved:", out)
    except Exception as e:
        print("Conversion failed:", e)


if __name__ == "__main__":
    main()
