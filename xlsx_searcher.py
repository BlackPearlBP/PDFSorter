import details_finder
import pandas as pd
import finance_searcher
import pathlib
import locale
import os

locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

# DIRECTORY WHERE THE CODE WILL FIND THE FILES FOR SEARCH
EXCEL_DIR = r"results"
directory2_path = None

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
    for root, _, files in os.walk(EXCEL_DIR): 
        for file in files:
            if not file.startswith("~"):
                if file.endswith(".xlsx"):
                    file_path = os.path.join(root, file) # Defines the path of the given file

                    name = file # Name of the file, for manual search if needed
                    country = details_finder.search_country(file_path) # Searches what country the invoice is from 
                    reference = details_finder.search_reference(file_path) # Searches for its reference number 
                    idNumbers = details_finder.find_id(country, file_path) # Searches for the Legal Entity Register Numbers (RUC, CUIT, RUT, EIN, CNPJ...) using the country's respective LERN patterns
                    order = details_finder.search_order(file_path) # Searches for its order number, if given any
                    total = finance_searcher.search_amounts(file_path)
                    tax = finance_searcher.search_tax(file_path) # Searches for its tax cost
                    if tax is not None:
                        tax_info = tax
                    currency = finance_searcher.search_currency(file_path) # Searches for the currency used
                    date = details_finder.search_dates(file_path) # Searches for its issue date
                    important_data = [name,str(date), country, idNumbers, currency, reference, order, total, tax_info] # Places every found information on a list that will be converted into a .CSV file later
                    convert_to_csv(name, important_data)
    delete_converted(EXCEL_DIR)


if __name__ == "__main__":
    main()