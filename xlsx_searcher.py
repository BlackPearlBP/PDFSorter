from decimal import Decimal
import pandas as pd
from datetime import datetime
import pathlib
import locale
import re
import os

locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

# DIRECTORY WHERE THE CODE WILL FIND THE FILES FOR SEARCH
RESULTS_DIR = r"results"
directory1_path = None
directory2_path = None

# BOSCH'S LEGAL ENTITY REGISTER NUMBERS ARE INSIDE THE FUNCTIONS THAT USE IT

# ALL BOSCH'S LATIN AMERICA LEGAL ENTITY NUMBERS:
'''
    20524506166 - PERU
    80097300-3 - PARAGUAY
    0992862467001 - ECUADOR
    217706890016 - URUGUAY
    30677423141 - ARGENTINA
    30-67742314-1 - ARGENTINA ALT
    155607295-2-2015 - PANAMA
'''

#OK
def search_ruc_peru(file_path):
    
    try:
        # RUC PERU
        RUC_PE = "20524506166"

        df = pd.read_excel(file_path) # Opens file
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº'] # Patterns to be found
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)] # search possible matches and put them in a list (na=False ignores missing values)
        if df[0].isnull().any(): # If the the list is empty, the code returns a warning message
            print("Warning: missing values found in column!")
        
        ruc = foundinfo[0].str.extract(r'([0-9-]+)', expand=False) # extracts only the information specified inside the brackets (numbers from 0-9 and hyphens) 

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False) # searches if the match has the parameter EIN (american standard)

        ruc = ruc.apply(lambda x: x if ein_mask.loc[ruc.index[ruc == x].tolist()[0]] and len(x) == 9 or len(x) == 11 else None) # applies the mask 9 for private limited companies and 11 for legal entities

        ruc = ruc.dropna() # remove missing values if any of them passed through 

        ruc = ruc.apply(lambda x: 'RUC BOSCH: ' + x if x == RUC_PE else 'RUC FORNECEDOR: ' + x) # applies labels to differentiate between Bosch's Id and supplier's Id
        return ruc.values.tolist() # return found values
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_ruc_paraguay(file_path):
    # RUC PARAGUAY
    RUC_PY = "80097300-3"
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº']
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        if df[0].isnull().any():
            print("Warning: missing values found in column!")
       
        ruc = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False)

        ruc = ruc.apply(lambda x: x if ein_mask.loc[ruc.index[ruc == x].tolist()[0]] and len(x) == 8 or len(x) == 9 or 10 or 11 else None)

        ruc = ruc.dropna()

        ruc = ruc.apply(lambda x: 'RUC BOSCH: ' + x if x == RUC_PY else 'RUC FORNECEDOR: ' + x)
        return ruc.values.tolist()
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_rut_uruguay(file_path):
    # RUT URUGUAY
    RUT_UY = "217706890016"
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

        rut = rut.apply(lambda x: 'RUT BOSCH: ' + x if x == RUT_UY else 'RUT FORNECEDOR: ' + x)
        return rut.values.tolist()
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_rut_ecuador(file_path):
    # RUT ECUADOR
    RUT_EC = "0992862467001"
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº', 'RUC / CI:','R.U.C.']
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        if df[0].isnull().any():
            print("Warning: missing values found in column!")
       
        ruc = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False)

        ruc = ruc.apply(lambda x: x if ein_mask.loc[ruc.index[ruc == x].tolist()[0]] and len(x) == 12 or 13 else None)

        ruc = ruc.dropna()

        ruc = ruc.apply(lambda x: 'RUC BOSCH: ' + x if x == RUT_EC else 'RUC FORNECEDOR: ' + x)
        return ruc.values.tolist()
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def search_ruc_panama(file_path):
    # RUC PANAMA
    RUC_PC = "155607295-2-2015"
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº','Identificación Fiscal']
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        if df[0].isnull().any():
            print("Warning: missing values found in column!")
       
        ruc = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False)

        ruc = ruc.apply(lambda x: x if ein_mask.loc[ruc.index[ruc == x].tolist()[0]] and len(x) == 15 or 16 or 17 else None)

        ruc = ruc.dropna()

        ruc = ruc.apply(lambda x: 'RUC BOSCH: ' + x if x == RUC_PC else 'RUC FORNECEDOR: ' + x)
        return ruc.values.tolist()
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#WIP
def search_cuit_argentina(file_path):
    # CUIT ARGENTINA
    CUIT_BOSCH = "30677423141"
    CUIT_BOSCH_ALT = "30-67742314-1"
    try:
        df = pd.read_excel(file_path)
        searchfor = ['RUC','R.U.C.','C.U.I.T.','CUIT','TIN','T.I.N.','CNPJ','EIN','E.I.N.', 'R.U.C. N°', 'R.U.C. Nº']
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)]
        if df[0].isnull().any():
            print("Warning: missing values found in column!")
       
        cuit = foundinfo[0].str.extract(r'([0-9-]+)', expand=False)

        ein_mask = foundinfo[0].str.contains('EIN|E.I.N.', na=False)

        cuit = cuit.apply(lambda x: x if ein_mask.loc[cuit.index[cuit == x].tolist()[0]] and len(x) == 11 or 12 or 13 else None)

        cuit = cuit.dropna()
        if not cuit.empty:
            cuit = cuit.apply(lambda x: 'CUIT BOSCH: ' + x if x.replace('-', '') == CUIT_BOSCH.replace('-', '') or x.replace('-', '') == CUIT_BOSCH_ALT.replace('-', '') else 'CUIT FORNECEDOR: ' + x)
            return cuit.values.tolist()
        else:
            return []
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OLD (Older method, overcomplicated)
""" def parse_monetary_value(value_str):
    # Remove non-numeric characters, except for negative signs and decimal separators
    value_str = re.sub(r'[^\\d\\.\\-,]+', '', value_str)

    # Handle multiple dots or commas
    if ',' in value_str and '.' in value_str:
        # Check if the comma is a thousands separator or a decimal separator
        if value_str.count(',') > value_str.count('.'):
            # Format X,XXX.XX
            value_str = value_str.replace(',', '')
        else:
            # Format X.XXX,XX
            value_str = value_str.replace('.', '')
            value_str = value_str.replace(',', '')
    elif ',' in value_str:
        # Format X,XXX.XX
        value_str = value_str.replace(',', '')
    elif '.' in value_str:
        # Format X.XXX,XX
        pass
    else:
        # No dot or comma, assume it's a valid number
        pass

    # Convert to float
    value = float(value_str)
    return value """

#OK
def parse_monetary_value(value_str):
    return '{:,.2f}'.format(value_str)

#OK-ISH (Can't catch specific patterns, needs rework)
def search_amounts(file_path):
    """
    Searches for total amount values in an Excel file.

    Args:
        file_path (str): The path to the Excel file to search.

    Returns:
        list[float]: A list of total amount values found in the file. If an error occurs, returns None.

    Notes:
        This function searches for specific patterns in the first column of the Excel file to identify total amount values.
        It handles different number formats, including commas and dots as thousands separators or decimal separators.

    Raises:
        Exception: If an error occurs while processing the file.
    """
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        patterns = [r'Importe Total : S/',r'TOTAL\. S/',r'TOTAL\.',r'Importe Total: US$',r'TOTAL',r'TOTAL US$',r'Importe Total:',r'IMPORTE TOTAL: US$',r'Total',r'Importe Total',r'TOTAL S/',r'OP\. EXONERADA OP\. INAFECTA OP\. GRAVADA TOT\. DSCTO\. I\.S\.C I\.G\.V\. IMPORTE TOTAL',r'IMPORTE TOTAL S/',r'TOTAL DOCUMENTO US$',r'Importe total:',r'Importe Total USD',r'TOTAL VENTA US$',r'TOTAL: PEN',r'Importe total de la venta S/',r'Sub Total: % Tax: Sales Tax: Total Amount Due:']
        values = []
        for text in df[0]:
            match = re.search('|'.join(patterns), str(text))
            if match:
                if match.group() == r'Sub Total: % Tax: Sales Tax: Total Amount Due:':
                    next_row = df.iloc[df.index.get_loc(match.end()) + 1][0]                 
                    value_match = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d+)?', str(next_row))
                    if value_match:
                        value_str = value_match.group(0).replace('.', '.').replace(',', '').replace('$','')
                        value = parse_monetary_value(float(value_str))
                        values.append(value)
                elif match.group() == r'OP\. EXONERADA OP\. INAFECTA OP\. GRAVADA TOT\. DSCTO\. I\.S\.C I\.G\.V\. IMPORTE TOTAL':
                    next_row = df.iloc[df.index.get_loc(match.start()) + 1][0]
                    value_match = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d+)?', str(next_row))
                    if value_match:
                        value_str = value_match[0].replace('.', '.').replace(',', '')
                        values.append(float(value_str))
                else:
                    value_match = re.search(r'([0-9.,]+)', str.strip(text[match.end():]))
                    if value_match:
                        #value_str = value_match.group(0).replace(',','').replace('.','.').replace('$','')
                        value_str = Decimal(locale.atof(value_match.group(0)))
                        #value = parse_monetary_value(float(value_str))
                        values.append(value_str)
                        #values.append(value)
        if values:
            real_total = max(values)
            total_return = "Total: {:.2f}".format(real_total) 
            return total_return
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

'''
if match.group() in [r"Sub Total: % Tax: Sales Tax: Total Amount Due:"]:
                    next_row = df.iloc[df.index.get_loc(match.start()) + 1][0]
                    value = re.search(r'([0-9.,]+)', str(next_row.iloc[-1:]))
                    if value:
                        value_str = value.group(0).replace('.', '.').replace(',', '')
                        values.append(float(value_str))
'''

#OLD (Too complex and buggy)
""" def search_amounts(file_path):
    
    Searches for total amount values in an Excel file.

    Args:
        file_path (str): The path to the Excel file to search.

    Returns:
        list[float]: A list of total amount values found in the file. If an error occurs, returns None.

    Notes:
        This function searches for specific patterns in the first column of the Excel file to identify total amount values.
        It handles different number formats, including commas and dots as thousands separators or decimal separators.

    Raises:
        Exception: If an error occurs while processing the file.
    
    try:
        df = pd.read_excel(file_path)
        patterns = ['TOTAL','Total','Importe Total','TOTAL S/','TOTAL. S/',r'OP. EXONERADA OP. INAFECTA OP. GRAVADA TOT. DSCTO. I.S.C I.G.V. IMPORTE TOTAL','IMPORTE TOTAL S/','TOTAL DOCUMENTO US$','Importe total:','Importe Total USD','TOTAL VENTA US$', 'TOTAL: PEN','Importe total de la venta S/',r'Sub Total: % Tax: Sales Tax: Total Amount Due:']
        values = []
        for text in df[0]:
            match = re.search('|'.join(patterns), str(text))
            if match:
                if match.group() in [r"Sub Total: % Tax: Sales Tax: Total Amount Due:", r"OP. EXONERADA OP. INAFECTA OP. GRAVADA TOT. DSCTO. I.S.C I.G.V. IMPORTE TOTAL"]:
                    next_row_index = df.index.get_loc(match.start()) + 1
                    if next_row_index < len(df):
                        next_row = df.iloc[next_row_index][0]
                        value_match = re.search(r'([0-9.,]+)', str(next_row))
                        if value_match:
                            value_str = value_match.group(0).replace('.', '.').replace(',', '')
                            values.append(float(value_str))
                else:
                    value_match = re.search(r'([0-9.,]+)', str(text[match.end():]))
                    if value_match:
                        value_str = value_match.group(0)
                        if ',' in value_str and '.' in value_str:
                            # Check if the comma is a thousands separator or a decimal separator
                            if value_str.count(',') > value_str.count('.'):
                                # Format X,XXX.XX
                                value_str = value_str.replace(',', '.')
                            else:
                                # Format X.XXX,XX
                                value_str = value_str.replace('.', '.')
                                value_str = value_str.replace(',', '')
                        elif ',' in value_str:
                            # Format X,XXX.XX
                            value_str = value_str.replace(',', '')
                        elif '.' in value_str:
                            # Format X.XXX,XX
                            pass
                        else:
                            # No dot or comma, assume it's a valid number
                            pass
                        value_str = re.sub(r'[^\\d\\.]+', '', value_str)  # Remove everything except digits and dot
                        value = float(value_str)
                        values.append(value)
        return values
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None """

#OK
def search_dates(file_path):
    try:
        df = pd.read_excel(file_path) # Opens excel file
        searchfor = ['Fecha','FECHA','Date','DATE','EMISION','Emisión','EMISIÓN', 'emisión'] # Pattens to be found
        
        df = df.dropna(subset=[0]) # Remove null values from the first column
        
        foundinfo = df[df[0].str.contains('|'.join(searchfor), na=False)] # Adds found values to a list containing every possible value
        
        dates = [] # list for the formatted dates
        for text in foundinfo[0]:
            line_dates = re.findall(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{1,2}[-]\w{3}[-]\d{2,4}', str(text)) # regex patterns (explanation bellow)
            
                #\d is for matches where the string contains numbers
                #{1,2} is for matches where the string contains 1 or 2 numbers
                #[/-] is for matches where the string contains a slash or a hyphen
                #\w is for matches where the string contains a word
                
                #this way we can find dates in every format DD/MM/YYYY, MM/DD/YYYY, DD-MM-YYYY, DD-MONTH-YYYY (...)
            
            for date in line_dates:
                if 'llegada' not in text.lower(): # The found dates should only be the Issue date. Expiration and arrival dates will be ignored
                    
                    date_formats = ['%d/%m/%Y', '%d-%m-%Y', '%d-%m-%y', '%d/%m/%y', '%d-%b-%Y', '%d/%b/%Y'] # Supported date formats. Attention: no support for american format, needs validation for this
                    for date_format in date_formats:
                        try:
                            date_obj = datetime.strptime(date, date_format) # Creates a datetime object with the given date and format
                            dates.append(date_obj) # Adds the found dates to the list
                            break
                        except ValueError:
                            pass
        
        # To guarantee that the returned date is the Issue one, this snippet retrieves the earliest date possible from the list
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
        searchfor = ['PERU', 'Peru', 'peru', 'Paraguay','paraguay', 'PARAGUAY', 'URUGUAY', 'Uruguay', 'LIMA', 'Lima', 'Asunción', 'ASUNCIÓN', 'Ciudad del Este', 'Guayas', 'Ecuador', 'ECUABOSCH', 'PANAMA', 'Panama','Montevideo','MONTEVIDEO','ARGENTINA','Argentina','Buenos Aires','BUENOS AIRES','argentina']
        pattern = re.compile('|'.join(searchfor), re.IGNORECASE)
        foundinfo = df[0].apply(lambda x: pattern.search(str(x)))
        country_rows = [match.group() for match in foundinfo if match]
        results = []
        for country in country_rows:
            if country in ['LIMA', 'Lima', 'PERU', 'Peru', 'peru']:
                results.append('PERU')
            elif country in ['PARAGUAY', 'Paraguay', 'ASUNCIÓN', 'Asunción','paraguay']:
                results.append('PARAGUAY')
            elif country in ['URUGUAY', 'Uruguay', 'Montevideo', 'MONTEVIDEO']:
                results.append('URUGUAY')
            elif country in ['ECUADOR', 'Ecuador', 'Quito', 'QUITO','ECUABOSCH','Ecuabosch','ecuador','GUAYAS','Guayas']:
                results.append('ECUADOR')
            elif country in ['PANAMA', 'Panama']:
                results.append('PANAMA')
            elif country in ['ARGENTINA', 'Argentina', 'Buenos Aires','BUENOS AIRES','argentina']:
                results.append('ARGENTINA')
            else:
                print(f"Error while sorting country: {country}")
        if not results:
            print(f"No country found in file: {file_path}")
        
        results = list(set(results))
        return results
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
        return None

#OK
def search_currency(file_path):
    try:
        df = pd.read_excel(file_path)
        currencies = ['DOLARES', 'SOLES', 'DOLAR', 'Soles', 'Dolares', 'Dolar', 'DÓLARES', 'Dólares', 'DÓLAR', 'UYU', 'Peso Uruguayo', 'USD', '$', 'PYG', 'Gs', 'Guarani', 'GS', 'US$', 'Dollar', 'Dólares Americanos']
        pattern = re.compile('|'.join(currencies), re.IGNORECASE) # Creates patterns that wll be used for searching

        df_str = df[0].astype(str).dropna() # Converts the row where the text was found to string and removes null values

        foundinfo = df_str.apply(lambda x: pattern.search(x) if x else None) # Applies patterns and searches for them
        matches = [match.group() for match in foundinfo if match] # List of matches

        # How to catalog currency based on match input
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
            '$': 'Ambiguous/multiple values, manual search is recommended',
            'Gs': 'PYG',
            'Guarani': 'PYG',
            'GS': 'PYG',
            'US$': 'USD',
            'Dollar': 'USD',
            'PESOS': 'ARS',
            'ARS': 'ARS',
            'AR$': 'ARS',
            'GUARANÍES': 'PYG'
        }

        # For every match found, verify if the match is present on the currency map. If so, the result is returned
        result = ''
        for match in matches:
            if match in currency_map:
                result = currency_map[match]
                break # Invoices usually use only one currency, so after the first result there's no reason to continue searching
            else:
                result = 'Ambiguous/multiple values, manual search is recommended'

        return result
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
        return 'Currency not found!'

#OK
def search_order(file_path): # Not working properly. Sometimes it returns dates or unrelated numbers (Needs fixing)
    try:
        df = pd.read_excel(file_path)

        patterns = [
            r'Orden',
            r'PO',
            r'Order',
            r'PO',
            r'OC NRO',
            r'OC',
            r'Nro de Orden de compra',
            r'ORDEN DE COMPRA',
            r'OC\(PO {10}\)',
            r'Pedido',
            r'ORDCOMPRA',
            r'PED',
            r'PED:',
            r'Orden de Compra:',
            r'Nro. Orden',
            ]
        
        orders = []
        for text in df[0]:
            for pattern in patterns:
                match = re.search(pattern, str(text))
                if match:
                    if not (re.search(r'RUC.*' + re.escape(match.group()), str(text), re.IGNORECASE) or re.search(r'RUC\s*\d+', str(text), re.IGNORECASE) or re.search(r'R.U.C.\s*\d+', str(text), re.IGNORECASE) or re.search(r'RUC / CI:', str(text), re.IGNORECASE) or re.search(r'BCP\s*\d+', str(text), re.IGNORECASE)):
                    # Extract the order number from the text
                        order_match = re.search(r'\d{5,}', str(text[match.end():]))
                        if order_match:
                            orders.append(order_match.group())

            # Remove repeated order numbers
            orders = list(set(orders))

            if orders:
                return orders[0]
        else:
            return "Order number not found, manual search is recommended"
    except Exception as e:
        raise ValueError(f"Error processing file: {file_path} - {str(e)}")

#OK
def search_tax(file_path):
    try:
        df = pd.read_excel(file_path)

        patterns = [
            r'Impuesto',
            r'Tax',
            r'tax',
            r'I.G.V.',
            r'IGV 18%',
            r'IGV $',
            r'Igv $',
            r'IGV \( %\):',
            r'IGV \( %\) :',
            r'IGV 18% USD',
            r'NET AMOUNTS VAT AMOUNTS',
            r'IGV : S/',
            r'I.G.V. 18%',
            r'TOTAL IGV',
            r'TOTAL IGV USD',
            r'Total IGV USD',
            r'Sumatoria IGV',
            r'OP\. EXONERADA OP\. INAFECTA OP\. GRAVADA TOT\. DSCTO\. I\.S\.C I\.G\.V\. IMPORTE TOTAL',
            r'IGV 18.00% 1 PEN',
            r'IGV:',
            r'IVA',
            r'IVA Resp. insc. 21%',
            r'I.V.A. INSC.% 21',
            r'IVA.INSC.21.00%',
            r'IVA 21%',
            r'IVA 21,00',
            r'Iva Insc. 21,00%',
            r'TOTAL IVA',
            r'IVA:',
            r'10%',
            r'ITBMS\(7%\)',
            r'ITBMS',
            r'Total Impuesto',
            r'Total Impuestos',
            r'Impuestos',
            r'IVA 12%',
            r'I.G.V. 18.00%',
            r'IVA: 12%',
            r'I.V.A',
            r'Iva',
            r'Total iva\(22%\)',
            r'TOTAL IGV \(18%\) S\/'
            r'Sin Impuestos'
            r'IVA 12% :'
        ]

        values = []
        for text in df[0]:
            for pattern in patterns:
                match = re.search(pattern, str(text))
                if pattern == r'Sin Impuestos':
                    pass
                else:
                    if match:
                        if pattern == r'OP\. EXONERADA OP\. INAFECTA OP\. GRAVADA TOT\. DSCTO\. I\.S\.C I\.G\.V\. IMPORTE TOTAL':
                            next_row = df.iloc[df.index.get_loc(match.start()) + 1][0]
                            value = re.search(r'([0-9.,]+)', str(next_row))
                        elif pattern == r'NET AMOUNTS VAT AMOUNTS':
                            value = re.search(r'([0-9.,]+)', str(text[match.end():]))
                        elif pattern == r'IGV 18.00% 1 PEN':
                            matches = re.findall(r'([0-9.,]+)', str(text[match.end():]))
                            if len(matches) > 1:
                                value = matches[1]
                            else:
                                value = "No tax"
                        else:
                            value = re.search(r'([0-9.,]+)', str(text[match.end():]))
                        if value:
                            value_str = value.group(0).replace('.', '.').replace(',', '')
                            if value_str not in ['18','18.0','18,0','18.0','12.0','12,0','12','0','0.0','0,0','0,00','0.00','1']:
                                values.append(f"Tax: {float(value_str)}")
                        else:
                            value = "No tax"
                            values.append(value)
        
        if values:
            numerical_values = [float(value.split(": ")[1]) for value in values]
            return f"Tax: {min(numerical_values)}"
        else:
            return "No tax"

    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return "Error"

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
            r'No. \d{6}-\d{9}',  # No. XXXXXX-XXXXXXXXXX
            r'No. \d{3}-\d{3}-\d{8}',  # No. XXX-XXX-XXXXXXXXX
            r'V-\d{6}',  # No = V-XXXXXX
            r'No. \d{6} - \d{8}',  # No. XXXXXX - XXXXXXXXX
            r'No. \d{3} - \d{3} - \d{8}',  # No. XXX - XXX - XXXXXXXXX
            r'V - \d{6}',  # No = V - XXXXXX,
            r'F\d{3} N° \d{8}', # FXXX N° XXXXXXXX
            r'F\d{3} Nº \d{8}', # FXXX Nº XXXXXXXX
            r'F\d{3} \d{8}', # FXXX XXXXXXXX
            r'\d{6} \d{2}[/]', # XXXXXX XX/
            r'FAC\d{1}-\d{8}', #FACX-XXXXXXXX
            r'FACTURA N° \d{6}', #FACTURA N° XXXXXX
            r'Número: \d{10}', #Número: XXXXXXXXXX
            r'\d{3}-\d{3}-\d{7}', #XXX-XXX-XXXXXXX
            r'No. Factura \d{4} - \d{8}', #No. Factura XXXX - XXXXXXXX
            r'\d{4}-\d{8}', #XXXX-XXXXXXXX
            r'N° A-\d{4}-\d{8}', #A-XXXX-XXXXXXXX
            r'N° A-\d{5}-\d{8}', #AXXXXX-XXXXXXXX
            r'Punto de Venta: \d{5} Comp. Nro: \d{8}', #Punto de Venta: XXXX Comp. Nro. XXXXXXXX
            r'Factura \d{4}A\d{8}', #Factura XXXXAXXXX
            r'FACTURA \d{5} - \d{8}', #FACTURA XXXX - XXXXXXXX
            r'N° \d{3}-\d{3}-\d{7}[/]SUNAT' #N° XXX-XXX-XXXXXXX/SUNAT
        ]

        regex_patterns = [re.compile(pattern) for pattern in patterns]
        
        matches = []
        invoice_index = None
        for index, text in enumerate(df[0]):
            for pattern in regex_patterns:
                match = pattern.search(str(text))
                if match:
                    if (pattern.pattern == r'N° \d{3}-\d{3}-\d{7}[/]SUNAT' or re.search(r'\bAutorizado mediante\b.*' + re.escape(match.group()), str(text), re.IGNORECASE) or re.search(r'\bCta.Cte\b.*' + re.escape(match.group()), str(text), re.IGNORECASE) or re.search(r'\bBANCO\b.*' + re.escape(match.group()), str(text), re.IGNORECASE) or re.search(r'\bDerechos\b.*' + re.escape(match.group()), str(text), re.IGNORECASE) or re.search(r'\bBCP\b.*' + re.escape(match.group()), str(text), re.IGNORECASE) or re.search(r'\bMon\b.*' + re.escape(match.group()), str(text), re.IGNORECASE) or re.search(r'\bDAM\b.*' + re.escape(match.group()), str(text), re.IGNORECASE) or re.search(r'\bSUNAT\b.*' + re.escape(match.group()), str(text), re.IGNORECASE) or re.search(r'\bTEXTOS/\b.*' + re.escape(match.group()), str(text), re.IGNORECASE) or re.search(r'\bEmisor electrónico\b.*' + re.escape(match.group()), str(text), re.IGNORECASE)):
                        continue
                    elif match.group() == 'Invoice Date Customer Supplier':
                        invoice_index = index
                    else:
                        if pattern.pattern == r'\d{6} \d{2}[/]': # Extract the  first 6 numbers, as the last two are unrelated
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
        
        # Filter matches to keep only the longest match (avoids the emergence of undesired numbers that may have the same pattern but are unrelated to the reference)
        if matches:
            max_len = max(len(match) for match in matches)
            matches = [match for match in matches if len(match) == max_len]
        
        return matches
    
    except Exception as e:
        print(f"Error processing file: {file_path} - {str(e)}")
    return None

#OK
def find_data(country, file_path):
    """
    Searches for data in files associated with the specified countries.

    Args:
        country (str or list): A string or list of country names.
        file_path (str): The path to the file containing the data.

    Returns:
        list: A list of data found in the files associated with the specified countries.

    Raises:
        ValueError: If country is not a string or a list of strings.
        FileNotFoundError: If the file associated with a country does not exist.
        IOError: If the file associated with a country cannot be read.
    """

    # Check if country is a string or a list of strings
    if not isinstance(country, (str,type([]))):
        raise ValueError("country must be a string or a list of strings")

    # Convert string with multiple countries to a list
    if isinstance(country, str):
        country = country.split(',')

    for c in [x.strip() for x in country]:
        if c:
            match c.upper():
                case 'PERU':
                    try:
                        found_info = search_ruc_peru(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case 'PARAGUAY':
                    try:
                        found_info = search_ruc_paraguay(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case 'URUGUAY':
                    try:
                        found_info = search_rut_uruguay(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case 'ECUADOR':
                    try:
                        found_info = search_rut_ecuador(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case 'PANAMA':
                    try:
                        found_info = search_ruc_panama(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case 'ARGENTINA':
                    try:
                        found_info = search_cuit_argentina(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case _:
                    pass
    return None

#OK
def convert_to_csv(name: str,important_data: list) -> None:

    file_name = os.path.splitext(os.path.basename(name))[0]
    suffix = ".csv"
    output = pathlib.Path(directory2_path).joinpath(file_name).with_suffix(suffix)

    #if not os.path.exists(CSV_FILES):
    #    os.makedirs(CSV_FILES)

    df = pd.DataFrame(important_data).T # Transpose the list to create a DataFrame with the correct number of rows
    df.to_excel(output)

def delete_converted(directory):
    for file in os.listdir(directory):
        os.remove(os.path.join(directory, file))

# For every file present at the directory (independent of the OS), search the ones that are .xslx
def main():
    for root, _, files in os.walk(RESULTS_DIR): 
        for file in files:
            if not file.startswith("~"):
                if file.endswith(".xlsx"):
                    file_path = os.path.join(root, file) # Defines the path of the given file

                    name = file # Name of the file, for manual search if needed
                    country = search_country(file_path) # Searches what country the invoice is from 
                    reference = search_reference(file_path) # Searches for its reference number 
                    idNumbers = find_data(country, file_path) # Searches for the Legal Entity Register Numbers (RUC, CUIT, RUT, EIN, CNPJ...) using the country's respective LERN patterns
                    order = search_order(file_path) # Searches for its order number, if given any
                    total = search_amounts(file_path)
                    tax = search_tax(file_path) # Searches for its tax cost
                    if tax is not None:
                        tax_info = tax
                    currency = search_currency(file_path) # Searches for the currency used
                    date = search_dates(file_path) # Searches for its issue date
                    important_data = [name,str(date), country, idNumbers, currency, reference, order, total, tax_info] # Places every found information on a list that will be converted into a .CSV file later
                    convert_to_csv(name, important_data)
    delete_converted(RESULTS_DIR)


if __name__ == "__main__":
    main()