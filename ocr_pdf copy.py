import glob
import ocrmypdf
import pdfplumber
import os
import pandas as pd
from pathlib import Path

PDF_DIR = r"C:\Users\OLB5JVL\Downloads\KeepTrue outros\KeepTrue05032024\RBPE"
RESULTS_DIR = r"results"
CONVERTED_DIR = r"converted_pdfs"

if not os.path.exists(RESULTS_DIR):
    os.makedirs(RESULTS_DIR)

array_pdfs = glob.glob(os.path.join(PDF_DIR, "*.pdf"))


for file in array_pdfs: #go through every file in the folder
    try:
        with pdfplumber.open(file) as pdf:
            if len(pdf.pages) > 0:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text.isprintable() == False: #checks if the pdf is searchable or not
                        ocrmypdf.ocr(file, CONVERTED_DIR, skip_text=True, deskew=True, force_ocr=True)
                        print("converted file")
                    else:
                        extracted_lines = []
                        
                        file_name = os.path.splitext(os.path.basename(file))[0]

                        suffix = ".xlsx"
                        for page in pdf.pages:
                            text = page.extract_text(x_tolerance=2).split('\n')
                            extracted_lines.extend(text)

                            data = [[item] for item in extracted_lines]

                            output = Path(RESULTS_DIR,file_name).with_suffix(suffix)
                            df = pd.DataFrame(data)

                            df.to_excel(output)
            else:
                print(f"Skipping empty PDF file: {file}")
    except pdfplumber.errors.PdfError as e:
        print(f"Error processing PDF file: {file} - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

