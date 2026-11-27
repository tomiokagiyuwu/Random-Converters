from pathlib import Path
import numpy as np
import pymupdf


def parse_pages(spec, total):
    if spec.strip().lower() == "all":
        return list(range(total))
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


def is_blank(pg):
    if pg.get_text().strip():
        return False
    pix = pg.get_pixmap(matrix=pymupdf.Matrix(1, 1), colorspace=pymupdf.csGRAY, alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8)
    return float(np.std(arr)) < 3.5 and float(np.mean(arr)) > 245.0


def main():
    print("\n=== EDIT PDF PAGES ===")
    src = Path(input("PDF file path: ").strip().strip('"'))
    if not src.is_file() or src.suffix.lower() != ".pdf":
        print("File not found.")
        return

    doc = pymupdf.open(str(src))
    total = doc.page_count
    print(f"Total pages: {total}")

    print("\n1. Rotate pages")
    print("2. Reorder pages")
    print("3. Delete pages")
    print("4. Remove blank pages")
    act = input("Choose (1/2/3/4): ").strip()

    out = Path(input("Output PDF path: ").strip().strip('"'))
    if out.suffix.lower() != ".pdf":
        out = out / f"{src.stem}_edited.pdf" if out.is_dir() else out.with_suffix(".pdf")

    same = src.resolve() == out.resolve()
    tmp = out.with_suffix(".tmp.pdf") if same else out

    try:
        if act == "1":
            spec = input(f"Pages to rotate (e.g. 1-3, 5 or all): ").strip() or "all"
            angle = int(input("Angle (90 / 180 / 270): ").strip())
            for i in parse_pages(spec, total):
                doc[i].set_rotation((doc[i].rotation + angle) % 360)

        elif act == "2":
            spec = input(f"New page order (e.g. 3, 1-2, 4-{total}): ").strip()
            order = parse_pages(spec, total)
            if not order:
                print("Invalid pages.")
                return
            doc.select(order)

        elif act == "3":
            spec = input("Pages to delete (e.g. 2, 4-5): ").strip()
            drop = set(parse_pages(spec, total))
            keep = [i for i in range(total) if i not in drop]
            if not keep:
                print("Cannot delete all pages.")
                return
            doc.select(keep)

        elif act == "4":
            keep = [i for i in range(total) if not is_blank(doc[i])]
            if not keep:
                keep = [0]
            print(f"Removed {total - len(keep)} blank page(s).")
            doc.select(keep)

        else:
            print("Invalid choice.")
            return

        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(tmp), garbage=3, deflate=True)
    except Exception as e:
        print("Failed:", e)
        return
    finally:
        doc.close()

    if same:
        tmp.replace(out)
    print("Saved:", out)


if __name__ == "__main__":
    main()
