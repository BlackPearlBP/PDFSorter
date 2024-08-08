import os
import pypdfium2
import glob
import pdfplumber
import pandas as pd
import imagepdf_converter
from pathlib import Path

PDF_DIR = r"C:\Users\OLB5JVL\Downloads\Keep True - RBPE\KP - RBPE"
RESULTS_DIR = r"results"
CONVERTED_DIR = r"converted_pdfs"

def process_pdf(file: str) -> None:
    try:
        with pdfplumber.open(file) as pdf:
            if len(pdf.pages) > 0:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text.isprintable():
                        pdf_to_convert_path = os.path.abspath(file)
                        file_name = os.path.splitext(os.path.basename(file))[0]
                        perform_ocr(pdf_to_convert_path, file_name)
                    else:
                        extracted_lines = []
                        for page in pdf.pages:
                            text = page.extract_text(x_tolerance=2).split('\n')
                            extracted_lines.extend(text)
                        convert_to_excel(file, extracted_lines)
            else:
                print(f"Skipping empty PDF file: {file}")
    except pdfplumber.errors.PdfError as e:
        print(f"Error processing PDF file: {file} - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

def perform_ocr(pdf_to_convert_path, file_name) -> None:
    to_convert = pypdfium2.PdfDocument(pdf_to_convert_path)
    for i in range(len(to_convert)):
        page = to_convert[i]
        image = page.render(scale=4).to_pil()
        image.save(f"\\converted_pdfs\\{file_name}.jpg")
        print("Converted file")

def convert_to_excel(file: str, extracted_lines: list) -> None:
    file_name = os.path.splitext(os.path.basename(file))[0]
    suffix = ".xlsx"
    output = Path(RESULTS_DIR,file_name).with_suffix(suffix)
    df = pd.DataFrame([item] for item in extracted_lines)
    df.to_excel(output)

def main() -> None:
    pdf_files = glob.glob(os.path.join(PDF_DIR, "*.pdf"))
    for file in pdf_files:
        process_pdf(file)

if __name__ == "__main__":
    main()