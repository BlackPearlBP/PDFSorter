import glob
import imagepdf_converter
import os
import pandas as pd
import pathlib
from pathlib import Path
import pdfplumber
import pypdfium2
import xlsx_searcher

EXCEL_DIR = r"results"
CONVERTED_DIR = r"converted_pdfs"
directory1_path = None

def process_pdfs(file: str) -> None:
    try:
        with pdfplumber.open(file) as pdf:
            if len(pdf.pages) > 0:
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text.strip() or text.isprintable():
                        pdf_to_convert_path = os.path.abspath(file)
                        file_name = os.path.splitext(os.path.basename(file))[0]
                        convert_pdf_to_jpg(pdf_to_convert_path, file_name)
                    else:
                        extracted_lines = []
                        for page in pdf.pages:
                            text = page.extract_text(x_tolerance=2).split('\n')
                            extracted_lines.extend(text)
                        convert_pdf_to_excel(file, extracted_lines)
            else:
                print(f"Skipping empty PDF file: {file}")
    except IOError as e:
        print(f"Error processing PDF file: {file} - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

def convert_pdf_to_jpg(pdf_to_convert_path, file_name) -> None:
    to_convert = pypdfium2.PdfDocument(pdf_to_convert_path)
    for i in range(len(to_convert)):
        page = to_convert[i]
        image = page.render(scale=4).to_pil()
        output_path = pathlib.Path(CONVERTED_DIR) / f"{file_name}_{i}.jpg"
        output_path.parent.mkdir(parents=True,exist_ok=True)
        image.save(output_path)
        print("Converted file")

def convert_pdf_to_excel(file: str, extracted_lines: list) -> None:
    file_name = os.path.splitext(os.path.basename(file))[0]
    suffix = ".xlsx"
    output = Path(EXCEL_DIR,file_name).with_suffix(suffix)
    df = pd.DataFrame([item] for item in extracted_lines)
    df.to_excel(output)

def main() -> None:
    pdf_files = glob.glob(os.path.join(directory1_path, "*.pdf"))
    for file in pdf_files:
        process_pdfs(file)
    
    imagepdf_converter.main()
    xlsx_searcher.main()