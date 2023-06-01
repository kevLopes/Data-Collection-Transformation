import EcosysData_API
import CostAnalyzeByTagNumber
import CostAnalyzeProcessByMaterial
import DataCollector


def action_for_ecosys_api(project_number):
    print("Analyze on going. Data are being refreshed from Ecosys")
    # Ecosys PO Lines
    EcosysData_API.ecosys_po_lines_data_api(
        f"https://ecosys-stg.sbmoffshore.com/ecosys/api/restjson/EcosysPOLinesData_DCTAPI_KOL"
        f"/?RootCostObject=MP{project_number}", "keven.deOliveiralope", "My-SBM#code23", project_number)

    # PO Header and SUN transactions API
    EcosysData_API.ecosys_poheader_data_api(
        f"https://ecosys-stg.sbmoffshore.com/ecosys/api"f"/restjson/EcosysPOHeadersData_DCTAPI_KOL"
        f"/?RootCostObject=MP{project_number}", "keven.deOliveiralope", "My-SBM#code23", project_number)

    EcosysData_API.ecosys_sun_lines_data_api(
        f"https://ecosys-stg.sbmoffshore.com/ecosys/api/restjson"f"/EcosysSUNData_DCTAPI_KOL"
        f"/?RootCostObject=MP{project_number}", "keven.deOliveiralope", "My-SBM#code23", project_number)


def action_for_material_analyze(project_number, material_type):
    folder_path_valve = "../Data Pool/Material Data Organized/Valve"  # self.folder_path_input.get()
    folder_path_piping = "../Data Pool/Material Data Organized/Piping"
    folder_path_bolt = "../Data Pool/Material Data Organized/Bolt"

    # Material Type Piping
    if material_type == "Piping":
        # Organize Data from Data Hub
        DataCollector.data_collector_piping(project_number, material_type)
        # Do something with the input and type (e.g., save to a file or database)
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_piping(folder_path_piping, project_number, material_type)
        CostAnalyzeByTagNumber.extract_distinct_tag_numbers_piping(folder_path_piping, project_number, material_type)
        print("Analyze done for all Piping Material Codes")
        flag = True

    # Material Type Valve
    elif material_type == "Valve":
        # Organize Data from Data Hub
        DataCollector.data_collector_valve(project_number, material_type)
        # Do something with the input and type (e.g., save to a file or database)
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_valve(folder_path_valve, project_number, material_type)
        print("Analyze done for all Valve Material Codes")
        flag = True

    # Material Type Bolt
    elif material_type == "Bolt":
        # Organize Data from Data Hub
        DataCollector.data_collector_bolt(project_number, material_type)
        # Do something with the input and type (e.g., save to a file or database)
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_bolt(folder_path_bolt, project_number, material_type)
        print("Analyze done for all Bolt Material Codes")
        flag = True

    # Material Type Structure
    elif material_type == "Structure":
        print("Implementation not done yet")

    # Material Type Bend
    elif material_type == "Bend":
        print("Implementation not done yet")

    # All Materials
    elif material_type == "All Materials":
        # print("Implementation not done yet")
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_piping(folder_path_piping, project_number, "Piping")
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_valve(folder_path_valve, project_number, "Valve")
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_bolt(folder_path_bolt, project_number, "Bolt")
        print("Cost Analyze done based on distinct Material types")
        flag = True