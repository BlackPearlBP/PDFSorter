from decimal import Decimal
import locale
import pandas as pd
import re

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