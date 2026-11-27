from pathlib import Path
import win32com.client as win32


def main():
    print("\n=== WORD TO PDF ===")
    print("1. Single file")
    print("2. Entire folder")

    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("Word file path: ").strip().strip('"'))
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
            for f in src.glob("*.docx")
            if not f.name.startswith("~$")
        ]
    else:
        print("Invalid choice.")
        return

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        for src_file, dst_file in files:
            doc = None
            try:
                print("Converting:", src_file.name)
                doc = word.Documents.Open(str(src_file.resolve()), ReadOnly=True)
                doc.ExportAsFixedFormat(str(dst_file.resolve()), 17)
                doc.Close(False)
                doc = None
                print("Done:", dst_file.name)
            except Exception as e:
                print("Failed:", src_file.name, e)
            finally:
                if doc:
                    doc.Close(False)
    finally:
        word.Quit()

    print("\nConversion completed!")


if __name__ == "__main__":
    main()
