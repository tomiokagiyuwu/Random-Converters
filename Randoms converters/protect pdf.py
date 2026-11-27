from pathlib import Path
import pymupdf


def lock_pdf(src, dst, pw):
    doc = pymupdf.open(str(src))
    same = src.resolve() == dst.resolve()
    tmp = dst.with_suffix(".tmp.pdf") if same else dst
    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        doc.save(
            str(tmp),
            encryption=pymupdf.PDF_ENCRYPT_AES_256,
            user_pw=pw,
            owner_pw=pw,
        )
    finally:
        doc.close()

    if same:
        tmp.replace(dst)


def unlock_pdf(src, dst, pw):
    doc = pymupdf.open(str(src))
    same = src.resolve() == dst.resolve()
    tmp = dst.with_suffix(".tmp.pdf") if same else dst
    try:
        if doc.is_encrypted and not doc.authenticate(pw):
            raise ValueError("Wrong password.")
        dst.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(tmp), encryption=pymupdf.PDF_ENCRYPT_NONE)
    finally:
        doc.close()

    if same:
        tmp.replace(dst)


def main():
    print("\n=== ENCRYPT / DECRYPT PDF ===")
    print("1. Encrypt (Add password)")
    print("2. Decrypt (Remove password)")
    act = input("Choose (1/2): ").strip()
    if act not in ("1", "2"):
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
            out = out / src.name if out.is_dir() else out.with_suffix(".pdf")
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

    pw = input("Password: ").strip()
    if not pw:
        print("Password cannot be empty.")
        return

    for src_file, dst_file in files:
        try:
            if act == "1":
                lock_pdf(src_file, dst_file, pw)
                print("Encrypted:", dst_file.name)
            else:
                unlock_pdf(src_file, dst_file, pw)
                print("Decrypted:", dst_file.name)
        except Exception as e:
            print("Failed:", src_file.name, e)

    print("\nDone!")


if __name__ == "__main__":
    main()
