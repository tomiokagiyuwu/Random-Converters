from pathlib import Path
import html
import re
import pymupdf

exts = (".txt", ".md", ".html", ".htm")


def md_to_html(text):
    out = []
    in_code = False

    for line in text.splitlines():
        s = line.strip()
        if s.startswith("```"):
            out.append("<pre>" if not in_code else "</pre>")
            in_code = not in_code
            continue

        if in_code:
            out.append(html.escape(line))
            continue

        if not s:
            out.append("<br>")
            continue

        line = html.escape(line)
        line = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", line)
        line = re.sub(r"\*(.+?)\*", r"<i>\1</i>", line)
        line = re.sub(r"`(.+?)`", r"<code>\1</code>", line)

        if s.startswith("# "):
            out.append(f"<h1>{line.lstrip('# ').strip()}</h1>")
        elif s.startswith("## "):
            out.append(f"<h2>{line.lstrip('# ').strip()}</h2>")
        elif s.startswith("### "):
            out.append(f"<h3>{line.lstrip('# ').strip()}</h3>")
        elif s.startswith(("- ", "* ")):
            out.append(f"<p>• {line.strip()[2:]}</p>")
        else:
            out.append(f"<p>{line}</p>")

    if in_code:
        out.append("</pre>")
    return "\n".join(out)


def text_to_pdf(src, dst):
    raw = src.read_text(encoding="utf-8", errors="replace")
    ext = src.suffix.lower()

    if ext == ".txt":
        body = f"<pre>{html.escape(raw)}</pre>"
    elif ext == ".md":
        body = md_to_html(raw)
    else:
        body = raw

    content = f"<html><body style='font-family:sans-serif;font-size:11pt;'>{body}</body></html>"
    story = pymupdf.Story(html=content)

    dst.parent.mkdir(parents=True, exist_ok=True)
    writer = pymupdf.DocumentWriter(str(dst))
    box = pymupdf.paper_rect("a4")
    rect = box + (36, 36, -36, -36)

    more = True
    while more:
        dev = writer.begin_page(box)
        more, _ = story.place(rect)
        story.draw(dev)
        writer.end_page()

    writer.close()


def main():
    print("\n=== TEXT / MARKDOWN / HTML TO PDF ===")
    print("1. Single file")
    print("2. Entire folder")

    choice = input("Choose (1/2): ").strip()

    if choice == "1":
        src = Path(input("File path: ").strip().strip('"'))
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
            if f.is_file() and f.suffix.lower() in exts
        ]
    else:
        print("Invalid choice.")
        return

    if not files:
        print("No valid files found.")
        return

    for src_file, dst_file in files:
        try:
            print("Converting:", src_file.name)
            text_to_pdf(src_file, dst_file)
            print("Done:", dst_file.name)
        except Exception as e:
            print("Failed:", src_file.name, e)

    print("\nConversion completed!")


if __name__ == "__main__":
    main()
