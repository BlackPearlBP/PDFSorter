import easyocr
import os
import pathlib
import pandas as pd
import logging

CONVERTED_DIR = r"converted_pdfs"
RESULTS_DIR = r"results"
LOG_FILE = r"ocr_log.txt"

logging.basicConfig(filename=LOG_FILE, level=logging.INFO)

def convert_to_excel(file: str, text: list) -> None:
    if not text:
        logging.info(f"Skipping empty file {file}")
        return

    file_name = os.path.splitext(os.path.basename(file))[0]
    suffix = ".xlsx"
    output = pathlib.Path(RESULTS_DIR).joinpath(file_name).with_suffix(suffix)

    if not os.path.exists(RESULTS_DIR):
        os.makedirs(RESULTS_DIR)

    df = pd.DataFrame([item] for item in text)
    df.to_excel(output)

def delete_converted(CONVERTED_DIR):
    for file in os.listdir(CONVERTED_DIR):
        os.remove(os.path.join(CONVERTED_DIR, file))

def main():
    reader = easyocr.Reader(['en', 'es'],detector="craft")

    for file in os.listdir(CONVERTED_DIR):
        if file.endswith(".jpg"):
            try:
                file_path = os.path.join(CONVERTED_DIR, file)
                if not os.path.isfile(file_path):
                    logging.warning(f"Skipping non-file {file_path}")
                    continue
                text = reader.readtext(file_path, detail=0, width_ths=1.5, height_ths=1, batch_size=10)
                logging.info(f"Extracted text from {file}: {text}")
                convert_to_excel(file, text)
            except easyocr.exceptions.TesseractConfigError as e:
                logging.error(f"Tesseract config error processing file {file}: {e}")
            except Exception as e:
                logging.error(f"Error processing file {file}: {e}")
    delete_converted(CONVERTED_DIR)

if __name__ == "__main__":
    main()