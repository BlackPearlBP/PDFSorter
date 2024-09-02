import pandas as pd
import re

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