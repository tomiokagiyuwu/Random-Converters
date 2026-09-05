import os
from pypdf import PdfMerger

folder_path = r"enter target pdf"

input_pdfs = [
    os.path.join(folder_path, f)
    for f in os.listdir(folder_path)
    if f.lower().endswith(".pdf")
]
input_pdfs.sort()

output_pdf = os.path.join(folder_path, "merged_output.pdf")

merger = PdfMerger()
for pdf in input_pdfs:
  merger.append(pdf)

merger.write(output_pdf)
merger.close()

print(f"Successfully merged {len(input_pdfs)} PDFs into: {output_pdf}")