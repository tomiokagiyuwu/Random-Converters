import os
import win32com.client as win32
# bash this in terminal to download win32 thing  "python -m pip install pywin32"

def convert_pptx_to_pdf(input_path):
    powerpoint = win32.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1

    if os.path.isfile(input_path):
        files = [input_path]

    elif os.path.isdir(input_path):
        files = [
            os.path.join(input_path, f)
            for f in os.listdir(input_path)
            if f.lower().endswith(".pptx")
        ]

    else:
        print("Error: File or folder not found.")
        powerpoint.Quit()
        return

    for pptx_file in files:
        pdf_file = os.path.splitext(pptx_file)[0] + ".pdf"

        print(f"Converting: {os.path.basename(pptx_file)}")

        deck = powerpoint.Presentations.Open(
            os.path.abspath(pptx_file),
            WithWindow=False
        )

        deck.SaveAs(os.path.abspath(pdf_file), 32)
        deck.Close()

        print(f"Done: {os.path.basename(pdf_file)}")

    powerpoint.Quit()

    print("Successfully converted!")


# Put either a .pptx FILE or a FOLDER here
input_path = r"C:\Users\Mayank Raj.LAPTOP-R5JI6HDP\Downloads\FOFA"

convert_pptx_to_pdf(input_path)