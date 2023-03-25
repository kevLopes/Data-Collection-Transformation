#Automatic Process
import DataCollector
import EcosysData_API

if __name__ == '__main__':

    DataCollector.dataCollectorFuncPiping()

    EcosysData_API.EcosysPOLineAPIData("https://Ecosys.sbm","keven.deOliveiralope","My-SBM#code23")

