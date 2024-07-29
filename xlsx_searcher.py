import pandas as pd
import re
import os

RESULTS_DIR = r"results"

#OK
def search_ruc_peru(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.']
        foundinfo = df[df[0].str.contains('|'.join(searchfor))]
        ruc = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)
        ruc = ruc[ruc.str.len() == 11]
        
        ruc = ruc.apply(lambda x: 'RUC BOSCH: ' + x if x == '20524506166' else 'RUC FORNECEDOR: ' + x if len(x) == 11 else x)
        return ruc.values.tolist()
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#WIP
def search_amounts(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['TOTAL','SUBTOTAL','Gravada','GRAVADA','Total','Subtotal']
        foundinfo = df[df[0].str.contains('|'.join(searchfor))]
        values = foundinfo[0].str.extract(r'([0-9-,.$]+)', expand=False)
        return values
        #print(values)
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#WIP
def search_dates(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['Fecha','FECHA','Date','DATE']
        foundinfo = df[df[0].str.contains('|'.join(searchfor))]
        dates = foundinfo[0].str.extract(r'([0-9-/]+)', expand=False)
        return dates
        #print(dates)
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_country(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['PERU', 'Peru', 'Paraguay', 'PARAGUAY', 'URUGUAY', 'Uruguay', 'LIMA', 'Lima', 'Asunción', 'ASUNCIÓN', 'Ciudad del Este', 'Guayas', 'Ecuador', 'ECUABOSCH', 'PANAMA', 'Panama','Montevideo','MONTEVIDEO']
        pattern = re.compile('|'.join(searchfor), re.IGNORECASE)
        foundinfo = df[0].apply(lambda x: pattern.search(str(x)))
        country_rows = [match.group() for match in foundinfo if match]
        for country in country_rows:
            if country == 'LIMA' or 'Lima' or 'PERU' or 'Peru':
                result = 'PERU'
                return result
            elif country == 'PARAGUAY' or 'Paraguay' or 'ASUNCIÓN' or 'Asunción':
                result = 'PARAGUAY'
                return result
            elif country == 'URUGUAY' or 'Uruguay' or 'Montevideo' or 'MONTEVIDEO':
                result = 'URUGUAY'
                return result
            elif country == 'ECUADOR' or 'Ecuador' or 'Quito' or 'QUITO':
                result = 'ECUADOR'
                return result
            elif country == 'PANAMA' or 'Panama':
                result = 'PANAMA'
                return result
            else:
                print(f"Error while sorting country")
        #print(country_rows)
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#WIP
def search_currency(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['SOLES','Soles','$','Dolares','DOLARES','Moneda']
        foundinfo = df[df[0].str.contains('|'.join(searchfor))]
        currency = foundinfo[0].str.extract(expand=False)
        return currency
        #print(currency)
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#WIP
def search_order(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['Orden','PO','Order']
        foundinfo = df[df[0].str.contains('|'.join(searchfor))]
        order = foundinfo[0].str.extract(r'([0-9-/]+)',expand=False)
        return order
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

def find_data(country, file_path):
    match country:
        case 'PERU':
            list = search_ruc_peru(file_path)
            return list
        case 'PARAGUAY':
            return "Not found"
        case 'URUGUAY':
            return "I'm a teapot"
        case 'ECUADOR':
            return 'abc'
        case 'PANAMA':
            return 'abc'
        case _:
            return "Country not found"


for root, _, files in os.walk(RESULTS_DIR):
    for file in files:
        if file.endswith(".xlsx"):
            file_path = os.path.join(root, file)
            print(file)
            name = file
            country = search_country(file_path)
            IdNumbers = find_data(country, file_path)
            #order = search_order(file_path)
            
            important_data = [name, country, IdNumbers]
            print(important_data)

            #print("DOC: "+name+"\nRUC/CUIT: "+ruc+"\nMoney Spent: "+str(amounts)+"\nCountry: "+str(country)+"\nCurrency: "+str(currency)+"\nOrder: "+str(order)+"\n")
