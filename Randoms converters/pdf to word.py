from pathlib import Path
import win32com.client as win32
from ocr import make_pdf_searchable


def main():
    print("\n=== PDF TO WORD ===")
    print("1. Single file")
    print("2. Entire folder")

    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("PDF file path: ").strip().strip('"'))
        if not src.is_file() or src.suffix.lower() != ".pdf":
            print("File not found.")
            return

        out = Path(input("Output Word path: ").strip().strip('"'))
        if out.suffix.lower() != ".docx":
            out = out / f"{src.stem}.docx" if out.is_dir() else out.with_suffix(".docx")

        files = [(src, out)]

    elif choice == "2":
        src = Path(input("Source folder: ").strip().strip('"'))
        out = Path(input("Output folder: ").strip().strip('"'))

        if not src.is_dir():
            print("Folder not found.")
            return

        out.mkdir(parents=True, exist_ok=True)
        files = [
            (f, out / f"{f.stem}.docx")
            for f in src.glob("*.pdf")
        ]
    else:
        print("Invalid choice.")
        return

    if not files:
        print("No PDF files found.")
        return

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        for src_file, dst_file in files:
            doc = None
            try:
                print("Converting:", src_file.name)
                make_pdf_searchable(src_file.resolve())
                doc = word.Documents.Open(
                    str(src_file.resolve()), ConfirmConversions=False, ReadOnly=True
                )
                dst_file.parent.mkdir(parents=True, exist_ok=True)
                doc.SaveAs2(str(dst_file.resolve()), FileFormat=16)
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
