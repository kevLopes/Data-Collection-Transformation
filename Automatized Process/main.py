#Automatic Process
import DataCollector
import EcosysData_API
import AnalyzeProcess

if __name__ == '__main__':
    while True:
        user_input = input("Do you want to call Ecosys PO Line API Data function? (Y/N)").lower()

        if user_input == "y":
            user_inputProj = input("For which project do you want to read the data (project number always start with MP)")
            EcosysData_API.ecosys_po_lines_data_api(f"https://ecosys-stg.sbmoffshore.com/ecosys/api/restjson/EcosysPOLinesData_DCTAPI_KOL/?RootCostObject={user_inputProj}","keven.deOliveiralope","My-SBM#code23")

            user_inputProj = input("For which project do you want to read the data (project number always start with MP)")
            EcosysData_API.EcosysSUNAPIData(f"https://ecosys-stg.sbmoffshore.com/ecosys/api/restjson/EcosysSUNData_DCTAPI_KOL/?RootCostObject={user_inputProj}","keven.deOliveiralope", "My-SBM#code23")
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")

    while True:
        user_input = input("Do you want to call Pipping Data Collector function? (Y/N)").lower()

        if user_input == "y":
            DataCollector.dataCollectorFuncPiping()
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")

    while True:
        user_input = input("Do you want to call Valve Data Collector function? (Y/N)").lower()

        if user_input == "y":
            DataCollector.dataCollectorFuncValve()
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")

    while True:
        user_input = input("Do you want to call Bolt Data Collector function? (Y/N)").lower()

        if user_input == "y":
            DataCollector.dataCollectorFuncBolt()
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")



    while True:
        user_input = input("Do you want to the complete analyze of the data collected? (Y/N)").lower()

        if user_input == "y":
            AnalyzeProcess.AnalyzeDataProcessFunc()
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")