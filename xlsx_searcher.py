import pandas as pd
from datetime import datetime
import numpy as np
import re
import os

RESULTS_DIR = r"results"

#OK
def search_ruc_peru(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº']
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        if df[0].isnull().any():
            print("Warning: missing values found in column!")
        
        ruc = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False)

        ruc = ruc.apply(lambda x: x if ein_mask.loc[ruc.index[ruc == x].tolist()[0]] and len(x) == 9 or len(x) == 11 else None)

        ruc = ruc.dropna()

        ruc = ruc.apply(lambda x: 'RUC BOSCH: ' + x if x == '20524506166' else 'RUC FORNECEDOR: ' + x)
        return ruc.values.tolist()
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_ruc_paraguay(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº']
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        if df[0].isnull().any():
            print("Warning: missing values found in column!")
       
        ruc = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False)

        ruc = ruc.apply(lambda x: x if ein_mask.loc[ruc.index[ruc == x].tolist()[0]] and len(x) == 8 or len(x) == 11 else None)

        ruc = ruc.dropna()

        ruc = ruc.apply(lambda x: 'RUC BOSCH: ' + x if x == '80097300-3' else 'RUC FORNECEDOR: ' + x)
        return ruc.values.tolist()
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_rut_uruguay(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº']
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        if df[0].isnull().any():
            print("Warning: missing values found in column!")
       
        rut = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False)

        rut = rut.apply(lambda x: x if ein_mask.loc[rut.index[rut == x].tolist()[0]] and len(x) == 9 or len(x) == 12 else None)

        rut = rut.dropna()

        rut = rut.apply(lambda x: 'RUT BOSCH: ' + x if x == '217706890016' else 'RUT FORNECEDOR: ' + x)
        return rut.values.tolist()
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_rut_ecuador(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº']
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        if df[0].isnull().any():
            print("Warning: missing values found in column!")
       
        ruc = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False)

        ruc = ruc.apply(lambda x: x if ein_mask.loc[ruc.index[ruc == x].tolist()[0]] and len(x) == 13 else None)

        ruc = ruc.dropna()

        ruc = ruc.apply(lambda x: 'RUC BOSCH: ' + x if x == '0992862467001' else 'RUC FORNECEDOR: ' + x)
        return ruc.values.tolist()
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_ruc_panama(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº']
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        if df[0].isnull().any():
            print("Warning: missing values found in column!")
       
        ruc = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False)

        ruc = ruc.apply(lambda x: x if ein_mask.loc[ruc.index[ruc == x].tolist()[0]] and len(x) == 16 else None)

        ruc = ruc.dropna()

        ruc = ruc.apply(lambda x: 'RUC BOSCH: ' + x if x == '155607295-2-2015' else 'RUC FORNECEDOR: ' + x)
        return ruc.values.tolist()
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_cuit_argentina(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº']
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        if df[0].isnull().any():
            print("Warning: missing values found in column!")
       
        cuit = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False)

        cuit = cuit.apply(lambda x: x if ein_mask.loc[cuit.index[cuit == x].tolist()[0]] and len(x) == 11 or 13 else None)

        cuit = cuit.dropna()

        cuit = cuit.apply(lambda x: 'CUIT BOSCH: ' + x if x == '30677423141' or '30-67742314-1' else 'CUIT FORNECEDOR: ' + x)
        return cuit.values.tolist()
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

#OK
def search_dates(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['Fecha','FECHA','Date','DATE','EMISION','Emisión','EMISIÓN']
        
        df = df.dropna(subset=[0])
        
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        
        dates = []
        for text in foundinfo[0]:
            line_dates = re.findall(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{1,2}[-]\w{3}[-]\d{2,4}', str(text))
            for date in line_dates:
                if 'llegada' not in text.lower():
                    
                    date_formats = ['%d/%m/%Y', '%d-%m-%Y', '%d-%m-%y', '%d/%m/%y', '%d-%b-%Y', '%d/%b/%Y']
                    for date_format in date_formats:
                        try:
                            date_obj = datetime.strptime(date, date_format)
                            dates.append(date_obj)
                            break
                        except ValueError:
                            pass
        
        if dates:
            earliest_date = min(dates)
            return ["Document date: " + earliest_date.strftime('%d/%m/%Y')]
        else:
            return []
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_country(file_path):
    try:
        df = pd.read_excel(file_path)
        searchfor = ['PERU', 'Peru', 'Paraguay', 'PARAGUAY', 'URUGUAY', 'Uruguay', 'LIMA', 'Lima', 'Asunción', 'ASUNCIÓN', 'Ciudad del Este', 'Guayas', 'Ecuador', 'ECUABOSCH', 'PANAMA', 'Panama','Montevideo','MONTEVIDEO','ARGENTINA','Argentina','Buenos Aires']
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
            elif country == 'ARGENTINA' or 'Argentina' or 'Buenos Aires':
                result = 'ARGENTINA'
                return result
            else:
                print(f"Error while sorting country")
        #print(country_rows)
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_currency(file_path):
    try:
        df = pd.read_excel(file_path)
        currencies = ['DOLARES', 'SOLES', 'DOLAR', 'Soles', 'Dolares', 'Dolar', 'DÓLARES', 'Dólares', 'DÓLAR', 'UYU', 'Peso Uruguayo', 'USD', '$', 'PYG', 'Gs', 'Guarani', 'GS', 'US$', 'Dollar', 'Dólares Americanos']
        pattern = re.compile('|'.join(currencies), re.IGNORECASE)

        df_str = df[0].astype(str).dropna()

        foundinfo = df_str.apply(lambda x: pattern.search(x) if x else None)
        matches = [match.group() for match in foundinfo if match]

        currency_map = {
            'SOLES': 'PEN',
            'DOLARES': 'USD',
            'UYU': 'UYU',
            'PYG': 'PYG',
            'USD': 'USD',
            'DÓLARES': 'USD',
            'DÓLAR': 'USD',
            'DÓLARES AMERICANOS': 'USD',
            'DÓLARES AMERICANOS': 'USD',
            'Soles': 'PEN',
            'DOLAR': 'USD',
            'Dólares': 'USD',
            'Dolares': 'USD',
            'Dólares Americanos': 'USD',
            'Dolar': 'USD',
            'Peso Uruguayo': 'UYU',
            '$': 'USD',
            'Gs': 'PYG',
            'Guarani': 'PYG',
            'GS': 'PYG',
            'US$': 'USD',
            'Dollar': 'USD'
        }

        result = ''
        for match in matches:
            if match in currency_map:
                result = currency_map[match]
                break

        return result
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
        return ''

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

#OK
def search_reference(file_path):
    try:
        df = pd.read_excel(file_path)
        
        patterns = [
            r'F\d{3}-\d{8}',  # FXXX-XXXXXXXX
            r'F\d{3} - \d{8}',  # FXXX - XXXXXXXX
            r'F[A-Z]\d{2}-\d{7}',  # FYXX-XXXXXXX
            r'F[A-Z]\d{2}-\d{8}',  # FYXX-XXXXXXXX
            r'F[A-Z]\d{2} - \d{7}',  # FYXX - XXXXXXX
            r'F[A-Z]\d{2} - \d{8}',  # FYXX - XXXXXXXX
            r'F\d{3}_\d{8}',  # FXXX_XXXXXXXX
            r'F\d{3} N° \d{8}',  # FXXX   N° XXXXXXXX
            r'F\d{3}-\d{6}',  # FXXX-XXXXXX
            r'F\d{3}-\d{7}',  # FXXX-XXXXXXX
            r'F\d{3}-\d{5}',  # FXXX-XXXXX
            r'E\d{3}-\d{3}',  # EXXX-XXX
            r'E\d{3}-\d{4}',  # EXXX-XXXXl
            r'F\d{3} - \d{6}',  # FXXX - XXXXXX
            r'F\d{3} - \d{7}',  # FXXX - XXXXXXX
            r'F\d{3} - \d{5}',  # FXXX - XXXXX
            r'E\d{3} - \d{3}',  # EXXX - XXX
            r'E\d{3} - \d{4}',  # EXXX - XXXX
            r'Número: \d{6}',  # Número: XXXXXX
            r'Nro. F\d{3}-\d{8}',  # Nro. FXXX-XXXXXXXX
            r'\d{3}-\d{7}',  # XXX-XXXXXXX
             r'Nro. F\d{3} - \d{8}',  # Nro. FXXX - XXXXXXXX
            r'\d{3} - \d{7}',  # XXX - XXXXXXX
            r'Invoice Date Customer Supplier',  # Invoice -> first numbers on the next line
            r'Série \| Número.*\n.*A\d{5}',  # Serie | Numero -> on the next line, using these models: -> AXXXXX
            r'F\d{7}',  # FXXXXXXX
            r'No. \d{6}-\d{8}',  # No. XXXXXX-XXXXXXXXX
            r'No. \d{3}-\d{3}-\d{8}',  # No. XXX-XXX-XXXXXXXXX
            r'V-\d{6}',  # No = V-XXXXXX
            r'No. \d{6} - \d{8}',  # No. XXXXXX - XXXXXXXXX
            r'No. \d{3} - \d{3} - \d{8}',  # No. XXX - XXX - XXXXXXXXX
            r'V - \d{6}',  # No = V - XXXXXX,
            r'F\d{3} N° \d{8}', # FXXX N° XXXXXXXX
            r'F\d{3} Nº \d{8}', # FXXX Nº XXXXXXXX
            r'F\d{3} \d{8}', # FXXX XXXXXXXX
            r'\d{6} \d{2}[/]', # XXXXXX XX/
            r'FAC\d{1}-\d{8}' #YYY-XXXXXXXX 
        ]
        
        regex_patterns = [re.compile(pattern) for pattern in patterns]
        
        matches = []
        invoice_index = None
        for index, text in enumerate(df[0]):
            for pattern in regex_patterns:
                match = pattern.search(str(text))
                if match:
                    if match.group() == 'Invoice Date Customer Supplier':
                        invoice_index = index
                    else:
                        if pattern.pattern == r'\d{6} \d{2}[/]': 
                            matches.append(match.group()[:6])
                        else:
                            matches.append(match.group())
            if invoice_index is not None and index == invoice_index + 1:
                numbers = re.findall(r'\d{6}', str(text))
                if numbers:
                    matches.append(numbers[0])
                invoice_index = None
        
        # Remove repeated matches
        matches = list(set(matches))
        
        # Filter matches to keep only the longest match
        if matches:
            max_len = max(len(match) for match in matches)
            matches = [match for match in matches if len(match) == max_len]
        
        return matches
    
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def find_data(country, file_path):
    match country:
        case 'PERU':
            list = search_ruc_peru(file_path)
            return list
        case 'PARAGUAY':
            list = search_ruc_paraguay(file_path)
            return list
        case 'URUGUAY':
            list = search_rut_uruguay(file_path)
            return list
        case 'ECUADOR':
            list = search_rut_ecuador(file_path)
            return list
        case 'PANAMA':
            list = search_ruc_panama(file_path)
            return list
        case 'ARGENTINA':
            list = search_cuit_argentina(file_path)
            return list
        case _:
            return "Country not found"


for root, _, files in os.walk(RESULTS_DIR):
    for file in files:
        if file.endswith(".xlsx"):
            file_path = os.path.join(root, file)

            name = file
            country = search_country(file_path)
            reference = search_reference(file_path)
            print(name)
            print(reference)
            idNumbers = find_data(country, file_path)
            #order = search_order(file_path)
            currency = search_currency(file_path)
            date = search_dates(file_path)
            important_data = [name, date, country, idNumbers, currency]
            #print(important_data)

#old def search_ruc_peru(file_path):
#     try:
#         df = pd.read_excel(file_path)
#         searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº']
#         foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
#         if df[0].isnull().any():
#             print("Warning: missing values found in column!")
#         print(file_path)
#         print(foundinfo)
#         ruc = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)
#         ruc = ruc[ruc.str.len() == 11]
        
#         ruc = ruc.apply(lambda x: 'RUC BOSCH: ' + x if x == '20524506166' else 'RUC FORNECEDOR: ' + x if len(x) == 11 else x)
#         return ruc.values.tolist()
#     except Exception as e:
#         print(f"Error processing file: {file_path} - {str(e)}")
#     return None
