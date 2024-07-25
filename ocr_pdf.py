import glob
import pdfplumber
import os
import pandas as pd
from pathlib import Path

PDF_DIR = r"C:\Users\OLB5JVL\Downloads\KeepTrue outros\KeepTrue05032024\RBPE"
RESULTS_DIR = r"results"

if not os.path.exists(RESULTS_DIR):
    os.makedirs(RESULTS_DIR)

array_pdfs = glob.glob(os.path.join(PDF_DIR, "*.pdf"))


for file in array_pdfs:
    try:
        with pdfplumber.open(file) as pdf:
            if len(pdf.pages) > 0:
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

