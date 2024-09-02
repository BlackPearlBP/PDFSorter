from datetime import datetime
import legal_entity_finder
import pandas as pd
import re

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
def find_id(country, file_path):
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
                        found_info = legal_entity_finder.search_ruc_peru(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case 'PARAGUAY':
                    try:
                        found_info = legal_entity_finder.search_ruc_paraguay(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case 'URUGUAY':
                    try:
                        found_info = legal_entity_finder.search_rut_uruguay(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case 'ECUADOR':
                    try:
                        found_info = legal_entity_finder.search_rut_ecuador(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case 'PANAMA':
                    try:
                        found_info = legal_entity_finder.search_ruc_panama(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case 'ARGENTINA':
                    try:
                        found_info = legal_entity_finder.search_cuit_argentina(file_path)
                        return found_info
                    except FileNotFoundError:
                        print(f"File not found for country: {c}")
                    except IOError:
                        print(f"Error reading file for country: {c}")
                case _:
                    pass
    return None
