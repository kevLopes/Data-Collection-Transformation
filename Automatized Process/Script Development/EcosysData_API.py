import requests
import pandas as pd
import json
from datetime import datetime
import os


def ecosys_poheader_data_api(api, username, password, project_number):

    try:
        # Make API request and get JSON response
        response = requests.get(api, auth=(username, password), verify=False)
        response.raise_for_status()
        json_data = json.loads(response.content)

        # Convert JSON data to Pandas DataFrame and select desired columns
        df = pd.DataFrame(json_data['CostObjectList'])

        # select desired columns
        df = df[['HierarchyPathID', 'PONumber', 'PODescription', 'POValue',
                 'ProjectCurrency', 'POIssueDate', 'POCurrencyID', 'ProductID', 'ProductDescription',
                 'SupplierNumber', 'SupplierName', 'CurrentPORevisionID']]

        # Map existing column names to new column names
        column_name_mapping = {
            'HierarchyPathID': 'Hierarchy Path ID',
            'PONumber': 'PO Number',
            'PODescription': 'PO Description',
            'POValue': 'PO Cost',
            'ProjectCurrency': 'Project Currency',
            'POIssueDate': 'PO Issue Date',
            'POCurrencyID': 'PO Currency',
            'ProductID': 'Product Code',
            'ProductDescription': 'Description',
            'SupplierNumber': 'Supplier Number',
            'SupplierName': 'Supplier Name',
            'CurrentPORevisionID': 'PO Revision'
        }

        # Rename columns in the DataFrame
        df = df.rename(columns=column_name_mapping)

        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        # create an Excel writer object
        writer = pd.ExcelWriter(f'..\\Data Pool\\Ecosys API Data\\PO Headers\\'f'MP{project_number}_Ecosys POHeaders_'
                                f'{timestamp}.xlsx')

        # write the DataFrame to the Excel file
        df.to_excel(writer, index=False)

        # save the Excel file
        writer.save()
        print('PO Header file extracted from Ecosys successfully!')

    except requests.exceptions.RequestException as re:
        error_msg = f"RequestException: {re}"
        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        print(error_msg)
        log_error(error_msg, timestamp)

    except Exception as e:
        error_msg = f"Error: {e}"
        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        print(error_msg)
        log_error(error_msg, timestamp)


# --------------------------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------------------------

def log_error(error_msg, timestamp):
    error_logs_folder = "Error Logs"
    os.makedirs(error_logs_folder, exist_ok=True)
    error_file_name = f"Error_Logs_{timestamp}.txt"
    error_file_path = os.path.join(error_logs_folder, error_file_name)

    with open(error_file_path, "w") as error_file:
        error_file.write(error_msg)

# --------------------------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------------------------


def ecosys_sun_lines_data_api(api, username, password, project_number):

    try:
        # Make API request and get JSON response
        response = requests.get(api, auth=(username, password), verify=False)
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
        writer = pd.ExcelWriter(f'..\\Data Pool\\Ecosys API Data\\SUN Transactions\\MP{project_number}_Ecosys SUN Transactions_'
                                f'{timestamp}.xlsx')

        # write the DataFrame to the Excel file
        df.to_excel(writer, index=False)

        # save the Excel file
        writer.save()
        print('SUN Transactions file extracted from Ecosys successfully!')

    except requests.exceptions.RequestException as re:
        error_msg = f"RequestException: {re}"
        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        print(error_msg)
        log_error(error_msg, timestamp)

    except Exception as e:
        error_msg = f"Error: {e}"
        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        print(error_msg)
        log_error(error_msg, timestamp)


# --------------------------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------------------------


def ecosys_po_lines_data_api(api, username, password, project_number):

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
                 'CurrencyTransactionCode', 'TagNumber', 'CostElementROSDate', 'Units']]

        # Map existing column names to new column names
        column_name_mapping = {
            'CostObjectHierarchyPathID': 'Hierarchy Path ID',
            'CostObjectID': 'Cost Object ID',
            'CostBreakdownStructureHierarchyPathID': 'Product Code',
            'PORevisionID': 'PO Revision',
            'TransactionDate': 'Transaction Date',
            'POLineNumber': 'PO Line Number',
            'PODescription': 'PO Description',
            'UnitofMeasureID': 'UOM',
            'CostCostObjectCurrency': 'Cost Project Currency',
            'CurrencyCostObjectCode': 'Project Currency',
            'CostTransactionCurrency': 'Cost Transaction Currency',
            'CurrencyTransactionCode': 'Transaction Currency',
            'TagNumber': 'Tag Number',
            'CostElementROSDate': 'ROSDate',
            'Units': 'Quantity'
        }

        # Rename columns in the DataFrame
        df = df.rename(columns=column_name_mapping)

        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        # Save DataFrame to Excel file
        writer = pd.ExcelWriter(f'..\\Data Pool\\Ecosys API Data\\PO Lines\\MP{project_number}_Ecosys PO Lines_{timestamp}.xlsx')
        df.to_excel(writer, index=False)
        writer.save()
        print("PO Lines file extracted from Ecosys successfully!")

    except requests.exceptions.RequestException as re:
        error_msg = f"RequestException: {re}"
        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        print(error_msg)
        log_error(error_msg, timestamp)

    except Exception as e:
        error_msg = f"Error: {e}"
        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        print(error_msg)
        log_error(error_msg, timestamp)
