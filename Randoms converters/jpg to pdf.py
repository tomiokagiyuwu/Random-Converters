import os
import img2pdf

directory = r"enter the target directory"

images = [
    os.path.join(directory, f)
    for f in os.listdir(directory)
    if f.lower().endswith((".jpg", ".jpeg"))
]
images.sort()

output_pdf = os.path.join(directory, "output.pdf")

with open(output_pdf, "wb") as f:
  f.write(img2pdf.convert(images))

print(f"Successfully converted {len(images)} images to {output_pdf}")