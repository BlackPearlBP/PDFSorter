import glob
import pandas
import pdfplumber
import os
from pathlib import Path

array_pdfs = (glob.glob(r"C:\Users\OLB5JVL\Downloads\KeepTrue outros\KeepTrue05032024\RBPE\*.pdf"))

for file in array_pdfs:
    with pdfplumber.open(file) as pdf:
        page = pdf.pages[0]

        brute_archive_data = page.extract_tables()
        
        file_name = os.path.splitext(os.path.basename(file))[0]

        suffix = ".xlsx"

        dir_name = r"C:\Users\OLB5JVL\Desktop\Leitor PDF\results"

        output = Path(dir_name,file_name).with_suffix(suffix)

        #archive_data = {'Text': [brute_archive_data]}
        
        df = pandas.DataFrame(data=brute_archive_data)
        df.to_excel(output)







