import glob
import pandas
import pdfplumber
import os

# reader = PdfReader(r"C:\Users\OLB5JVL\Downloads\KeepTrue outros\KeepTrue05032024\RBPC\01_01_0000002031_001_0000.pdf")
# print(len(reader.pages))
# page = reader.pages[0]
# print(page.extract_text())

array_pdfs = (glob.glob(r"C:\Users\OLB5JVL\Downloads\KeepTrue outros\KeepTrue05032024\RBPC\*.pdf"))

#output = r"C:\Users\OLB5JVL\Desktop\Leitor PDF\results"

for file in array_pdfs:
    with pdfplumber.open(file) as pdf:
        page = pdf.pages[0]
        data = page.extract_table()
        file_name = data[43:48]
        output = r"C:\Users\OLB5JVL\Desktop\Leitor PDF\results\{file_name}.xlsx".format(file_name)
        #os.rename(file, file_name)
        df = pandas.DataFrame(data=data)
        df.to_excel(output)

# pdf = pdfplumber.open(r"C:\Users\OLB5JVL\Downloads\KeepTrue outros\KeepTrue05032024\RBPC\01_01_0000002031_001_0000.pdf")
# page = pdf.pages[0]
# tables = page.extract_table()
#pdf_data = tables[0:]