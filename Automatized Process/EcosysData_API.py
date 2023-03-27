import os
import requests
import pandas as pd
from requests.auth import HTTPBasicAuth
import xml.etree.ElementTree as ET
import json
import openpyxl
from datetime import datetime

def EcosysPOHeadersAPIData( api_url, username,password):
    # Create directory if it doesn't exist
    if not os.path.exists('Ecosys Data'):
        os.makedirs('Ecosys Data')

    # Set API URL and credentials
    #api_url = 'https://ecosys-stg.sbmoffshore.com/ecosys/api/restjson/EcosysPOLinesData_DCTAPI_KOL/?RootCostObject={RootCostObject}'
    #username = 'your_username'
    #password = 'your_password'

    # Set headers for JSON format
    headers = {'Content-type': 'application/json', 'Accept': 'application/json'}

    # Make API request with basic authentication
    response = requests.get(api_url, headers=headers, auth=HTTPBasicAuth(username, password))

    # Convert response from JSON to Python dictionary
    data = json.loads(response.content)

    # Get XML file from API response
    xml_string = data['xml']

    # Parse XML string into an ElementTree object
    root = ET.fromstring(xml_string)

    # Create Excel workbook and sheet
    wb = openpyxl.Workbook()
    ws = wb.active

    # Get the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    # Write XML data to Excel sheet
    for child in root:
        ws.append([child.tag, child.attrib])

    # Save Excel workbook to file
    wb.save(f'Ecosys Data/Ecosys_POHeader{timestamp}.xlsx')

#--------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------

def EcosysPOLineAPIData(api_url,username,password):
    # Create directory if it doesn't exist
    if not os.path.exists('Ecosys Data'):
        os.makedirs('Ecosys Data')

    # Set API URL and credentials
    #api_url = 'https://example.com/api'
    #username = 'your_username'
    #password = 'your_password'

    # Set headers for JSON format
    headers = {'Content-type': 'application/json', 'Accept': 'application/json'}

    # Make API request with basic authentication
    response = requests.get(api_url, headers=headers, auth=HTTPBasicAuth(username, password))

    # Convert response from JSON to Python dictionary
    data = json.loads(response.content)

    # Get XML file from API response
    xml_string = data['xml']

    # Parse XML string into an ElementTree object
    root = ET.fromstring(xml_string)

    # Create Excel workbook and sheet
    wb = openpyxl.Workbook()
    ws = wb.active

    # Get the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    # Write XML data to Excel sheet
    for child in root:
        ws.append([child.tag, child.attrib])

    # Save Excel workbook to file
    wb.save(f'Ecosys Data/Ecosys_POLine{timestamp}.xlsx')

    def EcosysPOLineAPIData(api_url, username, password):
        # Create directory if it doesn't exist
        if not os.path.exists('Ecosys Data'):
            os.makedirs('Ecosys Data')

        # Set API URL and credentials
        # api_url = 'https://example.com/api'
        # username = 'your_username'
        # password = 'your_password'

        # Set headers for JSON format
        headers = {'Content-type': 'application/json', 'Accept': 'application/json'}

        # Make API request with basic authentication
        response = requests.get(api_url, headers=headers, auth=HTTPBasicAuth(username, password))

        # Convert response from JSON to Python dictionary
        data = json.loads(response.content)

        # Get XML file from API response
        xml_string = data['xml']

        # Parse XML string into an ElementTree object
        root = ET.fromstring(xml_string)

        # Create Excel workbook and sheet
        wb = openpyxl.Workbook()
        ws = wb.active

        # Get the current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        # Write XML data to Excel sheet
        for child in root:
            ws.append([child.tag, child.attrib])

        # Save Excel workbook to file
        wb.save(f'Ecosys Data/Ecosys_POLine{timestamp}.xlsx')

#--------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------

def EcosysSUNAPIData(api_url,username,password):
    # Create directory if it doesn't exist
    if not os.path.exists('Ecosys Data'):
        os.makedirs('Ecosys Data')

    # Set API URL and credentials
    #api_url = 'https://example.com/api'
    #username = 'your_username'
    #password = 'your_password'

    # Set headers for JSON format
    headers = {'Content-type': 'application/json', 'Accept': 'application/json'}

    # Make API request with basic authentication
    response = requests.get(api_url, headers=headers, auth=HTTPBasicAuth(username, password))

    # Convert response from JSON to Python dictionary
    data = json.loads(response.content)

    # Get XML file from API response
    xml_string = data['xml']

    # Parse XML string into an ElementTree object
    root = ET.fromstring(xml_string)

    # Create Excel workbook and sheet
    wb = openpyxl.Workbook()
    ws = wb.active

    # Get the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    # Write XML data to Excel sheet
    for child in root:
        ws.append([child.tag, child.attrib])

    # Save Excel workbook to file
    wb.save(f'Ecosys Data/Ecosys_SUN{timestamp}.xlsx')

#--------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------

def save_api_data_to_excel(API, username, password):
    try:
        # make a request to the API and retrieve the data
        response = requests.get(API, auth=(username, password), verify=False)
        response.raise_for_status()  # raise exception if response status code is not 200

        # parse the response data as JSON
        data = response.json()

        # convert the data to a pandas DataFrame
        df = pd.DataFrame(data)

        # create an Excel writer object
        writer = pd.ExcelWriter('Ecosys PO Lines Data.xlsx')

        # write the DataFrame to the Excel file
        df.to_excel(writer, index=False)

        # save the Excel file
        writer.save()

        print('Data saved to Excel file successfully.')

    except requests.exceptions.RequestException as re:
        print('RequestException:', re)
    except Exception as e:
        print('Error:', e)