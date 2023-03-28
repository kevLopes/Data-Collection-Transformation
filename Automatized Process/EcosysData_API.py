import requests
import pandas as pd
import json
from datetime import datetime

def ecosys_poheader_lines_data_api(API, username, password):
    try:
        # Make API request and get JSON response
        response = requests.get(API, auth=(username, password), verify=False)
        response.raise_for_status()
        json_data = json.loads(response.content)

        # Convert JSON data to Pandas DataFrame and select desired columns
        df = pd.DataFrame(json_data['CostObjectList'])

        # select desired columns
        df = df[['HierarchyPathID', 'PONumber', 'PODescription', 'POValue',
                 'ProjectCurrency', 'POIssueDate', 'POCurrencyID', 'ProductID', 'ProductDescription',
                 'SupplierNumber', 'SupplierName', 'CurrentPORevisionID']]

        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        # create an Excel writer object
        writer = pd.ExcelWriter(f'C:\\Users\\keven.deoliveiralope\\Documents\\Data Analyze '
                                f'Automatization\\Scripts\\Data-Collection-Transformation-kevLopes-DCT\\Automatized '
                                f'Process\\Data Pool\\Ecosys API Data\\PO Headers\\Ecosys POHea  ders Lines_'
                                f'{timestamp}.xlsx')

        # write the DataFrame to the Excel file
        df.to_excel(writer, index=False)

        # save the Excel file
        writer.save()
        print('Loading Data . . .')
        print('Data saved to Excel file successfully.')

    except requests.exceptions.RequestException as re:
        print('RequestException:', re)
    except Exception as e:
        print('Error:', e)


#--------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------

def ecosys_sun_lines_data_api(API, username, password):
    try:
        # Make API request and get JSON response
        response = requests.get(API, auth=(username, password), verify=False)
        response.raise_for_status()
        json_data = json.loads(response.content)

        # Convert JSON data to Pandas DataFrame and select desired columns
        df = pd.DataFrame(json_data['WorkingForecastTransactionList'])

        # select desired columns
        df = df[['CostObjectHierarchyPathID', 'ExternalKey', 'TransactionReference', 'Description',
                 'InvoiceReference', 'PONumber', 'AccountCodeID', 'AccountCodeName', 'BusinessUnitID',
                 'CostCodeID', 'JournalTypeID', 'JournalNumber', 'JournalLine', 'JournalSourceID',
                 'CompanyID', 'SunSystemsPeriodID', 'TransactionDate', 'SunSystemsTransactionDate',
                 'Amount', 'Currency', 'ProjectAmount', 'ProjectCurrencyID', 'EntryDate',
                 'InvoiceHyperlink', 'SunSystemsPostingDate', 'SunSystemsBaseAmount', 'SunSystemsCpyCurrency']]

        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        # create an Excel writer object
        writer = pd.ExcelWriter(f'C:\\Users\\keven.deoliveiralope\\Documents\\Data Analyze '
                                f'Automatization\\Scripts\\Data-Collection-Transformation-kevLopes-DCT\\Automatized '
                                f'Process\\Data Pool\\Ecosys API Data\\SUN Transactions\\Ecosys SUN Transactions_'
                                f'{timestamp}.xlsx')

        # write the DataFrame to the Excel file
        df.to_excel(writer, index=False)

        # save the Excel file
        writer.save()
        print('Loading Data . . .')
        print('Data saved to Excel file successfully.')

    except requests.exceptions.RequestException as re:
        print('RequestException:', re)
    except Exception as e:
        print('Error:', e)


#--------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------

def ecosys_po_lines_data_api(api, username, password):
    try:
        # Make API request and get JSON response
        response = requests.get(api, auth=(username, password), verify=False)
        response.raise_for_status()
        json_data = json.loads(response.content)

        # Convert JSON data to Pandas DataFrame and select desired columns
        df = pd.DataFrame(json_data['WorkingForecastTransactionList'])
        df = df[['CostObjectHierarchyPathID', 'CostObjectID', 'CostBreakdownStructureHierarchyPathID',
                 'PORevisionID', 'TransactionDate', 'POLineNumber', 'PODescription', 'UnitofMeasureID',
                 'CostCostObjectCurrency', 'CurrencyCostObjectCode', 'CostTransactionCurrency',
                 'CurrencyTransactionCode', 'TagNumber', 'CostElementROSDate']]

        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        # Save DataFrame to Excel file
        writer = pd.ExcelWriter(f'C:\\Users\\keven.deoliveiralope\\Documents\\Data Analyze '
                                f'Automatization\\Scripts\\Data-Collection-Transformation-kevLopes-DCT\\Automatized '
                                f'Process\\Data Pool\\Ecosys API Data\\PO Lines\\Ecosys PO Lines Data_'
                                f'{timestamp}.xlsx')
        df.to_excel(writer, index=False)
        writer.save()

        print('Loading Data . . .')
        print("Excel file saved successfully!")

    except requests.exceptions.RequestException as e:
        print("Request error:", e)
    except (KeyError, ValueError) as e:
        print("Error parsing JSON data:", e)
    except Exception as e:
        print("An error occurred:", e)