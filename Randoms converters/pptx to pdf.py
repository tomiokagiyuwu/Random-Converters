from pathlib import Path
import win32com.client as win32


def main():
    print("\n=== POWERPOINT TO PDF ===")
    print("1. Single file")
    print("2. Entire folder")

    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("PPTX file path: ").strip().strip('"'))
        if not src.is_file():
            print("File not found.")
            return

        out = Path(input("Output PDF path: ").strip().strip('"'))
        if out.suffix.lower() != ".pdf":
            out = out / f"{src.stem}.pdf" if out.is_dir() else out.with_suffix(".pdf")

        files = [(src, out)]

    elif choice == "2":
        src = Path(input("Source folder: ").strip().strip('"'))
        out = Path(input("Output folder: ").strip().strip('"'))

        if not src.is_dir():
            print("Folder not found.")
            return

        out.mkdir(parents=True, exist_ok=True)
        files = [
            (f, out / f"{f.stem}.pdf")
            for f in src.glob("*.pptx")
            if not f.name.startswith("~$")
        ]
    else:
        print("Invalid choice.")
        return

    ppt = win32.DispatchEx("PowerPoint.Application")

    try:
        for src_file, dst_file in files:
            deck = None
            try:
                print("Converting:", src_file.name)
                deck = ppt.Presentations.Open(
                    str(src_file.resolve()),
                    ReadOnly=True,
                    Untitled=False,
                    WithWindow=False,
                )
                deck.SaveAs(str(dst_file.resolve()), 32)
                deck.Close()
                deck = None
                print("Done:", dst_file.name)
            except Exception as e:
                print("Failed:", src_file.name, e)
            finally:
                if deck:
                    deck.Close()
    finally:
        ppt.Quit()

    print("\nConversion completed!")


if __name__ == "__main__":
    main()
