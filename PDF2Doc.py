# Install: pip install pdf2docx

import os
from pdf2docx import Converter
from datetime import datetime

# Function to convert PDF to DOCX
def pdf_to_docx(pdf_file, docx_file):
    cv = Converter(pdf_file)
    cv.convert(docx_file, start=0, end=None)
    cv.close()

# Specify the input PDF file path here (change Path = copy and paste PDF Path)
pdf_file = r'E:\01_PXpler_File\P002_PDF_2_Doc\Bundeskleingartengesetz.pdf' 

# Get the current date and time
current_datetime = datetime.now().strftime("%Y%m%d%H%M%S")

# Create the output DOCX file path with os.path.join
output_dir = r'E:\01_PXpler_File\P002_PDF_2_Doc' # change Path = copy and paste Path
docx_file = os.path.join(output_dir, f'output_{current_datetime}.docx')

pdf_to_docx(pdf_file, docx_file)

print(f'PDF converted to DOCX: {docx_file}') # it works Tested and debug by me. Does not work with PDF picture, just docs.

"""
Install: pip install pdf2docx

I've used the r prefix before the file paths to indicate that they are raw string literals. 
This ensures that backslashes in the file path are treated as literal characters and not as escape characters. 
Now, when you run this code, it should correctly convert the PDF located at the specified path
to a DOCX file and save it to the output path you provided.

With this code, every time you run it, the output DOCX file will have a unique name that includes the current date and time, 
ensuring that you don't overwrite previous versions.
"""