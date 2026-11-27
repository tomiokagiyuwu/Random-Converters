from pathlib import Path
import win32com.client as win32

exts = (".xlsx", ".xls", ".csv")


def main():
    print("\n=== EXCEL TO PDF ===")
    print("1. Single file")
    print("2. Entire folder")

    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("Excel file path: ").strip().strip('"'))
        if not src.is_file() or src.suffix.lower() not in exts:
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
            for f in src.iterdir()
            if f.is_file() and f.suffix.lower() in exts and not f.name.startswith("~$")
        ]
    else:
        print("Invalid choice.")
        return

    if not files:
        print("No Excel files found.")
        return

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        for src_file, dst_file in files:
            wb = None
            try:
                print("Converting:", src_file.name)
                wb = excel.Workbooks.Open(str(src_file.resolve()), ReadOnly=True)
                for ws in wb.Worksheets:
                    try:
                        ws.PageSetup.Zoom = False
                        ws.PageSetup.FitToPagesWide = 1
                        ws.PageSetup.FitToPagesTall = False
                    except Exception:
                        pass
                dst_file.parent.mkdir(parents=True, exist_ok=True)
                wb.ExportAsFixedFormat(0, str(dst_file.resolve()))
                wb.Close(False)
                wb = None
                print("Done:", dst_file.name)
            except Exception as e:
                print("Failed:", src_file.name, e)
            finally:
                if wb:
                    wb.Close(False)
    finally:
        excel.Quit()

    print("\nConversion completed!")


if __name__ == "__main__":
    main()
