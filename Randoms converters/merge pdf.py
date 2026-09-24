from pypdf import PdfWriter
import os
#bash pip install pypdf

folder = r"enter the directory of folder"

writer = PdfWriter()

pdf_files = [
    f for f in os.listdir(folder)
    if f.lower().endswith(".pdf") and f != "merged.pdf" #the merged pdf will be there in the same fodler as the original one
]

pdf_files.sort()

for pdf in pdf_files:
    path = os.path.join(folder, pdf)
    print(f"Adding: {pdf}")
    writer.append(path)

output = os.path.join(folder, "merged.pdf")

with open(output, "wb") as f:
    writer.write(f)

print(f"\nMerged {len(pdf_files)} PDFs successfully!")
print(f"Output: {output}")