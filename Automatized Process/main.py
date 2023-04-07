#Automatic Process
import DataCollector
import EcosysData_API
import AnalyzeProcess
import AnalyzeProcessByMaterial

if __name__ == '__main__':

    # Example
    folder_path = "C:\\Users\\keven.deoliveiralope\\Documents\Data Analyze Automatization\\Scripts\\Data-Collection-Transformation-kevLopes-DCT\\Automatized Process\\Data Pool\\Data Organize"
    project_number = 17033
    material_type = "Piping"
    distinct_product_codes = AnalyzeProcessByMaterial.extract_distinct_product_codes(folder_path, project_number, material_type)
    print("Analyze done for all Distinct Material codes:", distinct_product_codes)


    while False:
        user_input = input("Do you want to call Ecosys PO Line API Data function? (Y/N)").lower()

        if user_input == "y":
            user_inputProj = input("For which project do you want to read the data (project number always start with "
                                   "MP)")
            EcosysData_API.ecosys_poheader_lines_data_api(f"https://ecosys-stg.sbmoffshore.com/ecosys/api/restjson/EcosysPOHeadersData_DCTAPI_KOL/?RootCostObject={user_inputProj}","keven.deOliveiralope", "My-SBM#code23")

            user_inputProj = input("For which project do you want to read the data (project number always start with MP)")
            EcosysData_API.ecosys_po_lines_data_api(f"https://ecosys-stg.sbmoffshore.com/ecosys/api/restjson/EcosysPOLinesData_DCTAPI_KOL/?RootCostObject={user_inputProj}","keven.deOliveiralope","My-SBM#code23")

            user_inputProj = input("For which project do you want to read the data (project number always start with MP)")
            EcosysData_API.ecosys_sun_lines_data_api(f"https://ecosys-stg.sbmoffshore.com/ecosys/api/restjson/EcosysSUNData_DCTAPI_KOL/?RootCostObject={user_inputProj}","keven.deOliveiralope", "My-SBM#code23")
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")

    while False:
        user_input = input("Do you want to call Pipping Data Collector function? (Y/N)").lower()

        if user_input == "y":
            DataCollector.data_collector_piping()
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")

    while False:
        user_input = input("Do you want to call Valve Data Collector function? (Y/N)").lower()

        if user_input == "y":
            DataCollector.data_collector_valve()
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")

    while False:
        user_input = input("Do you want to call Bolt Data Collector function? (Y/N)").lower()

        if user_input == "y":
            DataCollector.data_collector_bolt()
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")

    while False:
        user_input = input("Do you want to the complete analyze of the data collected? (Y/N)").lower()

        if user_input == "y":
            AnalyzeProcess.data_analyze_process()
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")
