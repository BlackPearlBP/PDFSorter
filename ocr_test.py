import ocrmypdf
FILE = r"C:\Users\OLB5JVL\Downloads\KeepTrue outros\KeepTrue05032024\RBPE\12214_20240305133501.511_X.pdf"
CONVERTED_DIR = r"converted_pdfs"
ocrmypdf.ocr(FILE, CONVERTED_DIR, skip_text=True, deskew=True)