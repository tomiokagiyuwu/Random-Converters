import os
from comtypes import client


def convert_pptx_to_pdf(input_path, output_path):
  powerpoint = client.CreateObject("PowerPoint.Application")
  powerpoint.Visible = 1
  deck = powerpoint.Presentations.Open(input_path)
  deck.SaveAs(output_path, 32)  
  deck.Close()
  powerpoint.Quit()


pptx_file = r"target file"
pdf_file = r"output file"

convert_pptx_to_pdf(pptx_file, pdf_file)
print("Successfully converted PPTX to PDF!")

