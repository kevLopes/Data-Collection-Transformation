#Automatic Process
import DataCollector
import EcosysData_API

if __name__ == '__main__':
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
        user_input = input("Do you want to call Ecosys PO Line API Data function? (Y/N)").lower()

        if user_input == "y":
            EcosysData_API.EcosysPOLineAPIData("https://Ecosys.sbm","keven.deOliveiralope","My-SBM#code23")
            break
        elif user_input == "n":
            break
        else:
            print("Invalid input. Please enter 'Y' or 'N'.")

