import EcosysData_API
import CostAnalyzeByTagNumber
import CostAnalyzeProcessByMaterial
import DataCollector
import AnalyzeProcessByYard
import AnalyzeProcessByMTO


def action_for_ecosys_api(project_number):
    print("Analyze on going. Data are being refreshed from Ecosys")
    # Ecosys PO Lines
    EcosysData_API.ecosys_po_lines_data_api("keven.deOliveiralope", "My-SBM#code23", project_number)

    # PO Header and SUN transactions API
    EcosysData_API.ecosys_poheader_data_api("keven.deOliveiralope", "My-SBM#code23", project_number)
    EcosysData_API.ecosys_sun_lines_data_api("keven.deOliveiralope", "My-SBM#code23", project_number)


#SBM Scope Analyze
def action_for_material_analyze(project_number, material_type):
    print("Analyze on going without collecting data from Ecosys")
    folder_path_valve = "../Data Pool/Material Data Organized/Valve"  # self.folder_path_input.get()
    folder_path_piping = "../Data Pool/Material Data Organized/Piping"
    folder_path_bolt = "../Data Pool/Material Data Organized/Bolt"
    folder_path_sp = "../Data Pool/Material Data Organized/Special Piping"
    folder_path_structure = "../Data Pool/Material Data Organized/Structure"

    # Material Type Piping
    if material_type == "Piping":
        # Organize Data from Data Hub
        #DataCollector.data_collector_piping(project_number, material_type)
        # Do something with the input and type
        #CostAnalyzeProcessByMaterial.extract_distinct_product_codes_piping(folder_path_piping, project_number, material_type)
        #CostAnalyzeByTagNumber.extract_distinct_tag_numbers_piping(folder_path_piping, project_number, material_type)
        AnalyzeProcessByMTO.analyze_by_material_type_piping(folder_path_piping, project_number, material_type)
        print("Analyze done for all SBM Scope Piping Materials")
        return True

    # Material Type Valve
    elif material_type == "Valve":
        # Organize Data from Data Hub
        DataCollector.data_collector_valve(project_number, material_type)
        # Do something with the input and type
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_valve(folder_path_valve, project_number, material_type)
        print("Analyze done for all SBM Scope Valve Materials")
        return True

    # Material Type Bolt
    elif material_type == "Bolt":
        # Organize Data from Data Hub
        DataCollector.data_collector_bolt(project_number, material_type)
        # Do something with the input and type
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_bolt(folder_path_bolt, project_number, material_type)
        print("Analyze done for all SBM Scope Bolt Materials")
        return True

    # Material Type Structure
    elif material_type == "Structure":
        # Organize Data from Data Hub
        DataCollector.data_collector_structure(project_number, material_type)
        # Do something with the input and type
        CostAnalyzeProcessByMaterial.analyze_structure_materials(folder_path_structure, project_number, material_type)
        print("Analyze done for Structure Materials")
        return True

    # Material Type Bend
    elif material_type == "Bend":
        # Organize Data from Data Hub
        DataCollector.data_collector_bend(project_number, material_type)
        print("Analyze done for all SBM Scope Bend Materials")
        return True

    # Material Type Special Piping
    elif material_type == "Special Piping":
        # Organize Data from Data Hub
        DataCollector.data_collector_specialpip(project_number, "Special PIP")
        # Do something with the input and type
        CostAnalyzeByTagNumber.extract_distinct_tag_numbers_special_piping(folder_path_sp, project_number, "SPC Piping")
        print("Analyze done for all SBM Scope Special Piping Materials")
        return True

    # All Materials
    elif material_type == "All Materials":
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_piping(folder_path_piping, project_number, "Piping")
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_valve(folder_path_valve, project_number, "Valve")
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_bolt(folder_path_bolt, project_number, "Bolt")
        CostAnalyzeByTagNumber.extract_distinct_tag_numbers_piping(folder_path_piping, project_number, "Piping")
        CostAnalyzeByTagNumber.extract_distinct_tag_numbers_special_piping(folder_path_sp, project_number, "SPC Piping")
        print("Cost Analyze done based on all distinct Material types")
        return True

    return False


#YARD Scope Analyze
def action_for_material_analyze_by_yard(project_number, material_type):
    folder_path_valve = "../Data Pool/Material Data Organized/Valve"
    folder_path_piping = "../Data Pool/Material Data Organized/Piping"
    folder_path_bolt = "../Data Pool/Material Data Organized/Bolt"
    folder_path_sp = "../Data Pool/Material Data Organized/Special Piping"

    # Material Type Piping
    if material_type == "Piping":
        # Organize Data from Data Hub
        DataCollector.data_collector_piping_yard(project_number, material_type)
        # Do something with the input and type
        AnalyzeProcessByYard.yard_piping_material_type_analyze(folder_path_piping, project_number, material_type)
        return True

    # Material Type Valve
    elif material_type == "Valve":
        # Organize Data from Data Hub
        DataCollector.data_collector_valve_yard(project_number, material_type)
        # Do something with the input and type
        AnalyzeProcessByYard.yard_valve_material_type_analyze(folder_path_valve, project_number, material_type)
        return True

    # Material Type Bolt
    elif material_type == "Bolt":
        # Organize Data from Data Hub
        DataCollector.data_collector_bolt_yard(project_number, material_type)
        # Do something with the input and type
        AnalyzeProcessByYard.yard_bolt_material_type_analyze(folder_path_bolt, project_number, material_type)
        return True

    # Material Type Structure
    elif material_type == "Structure":
        # Organize Data from Data Hub
        return True

    # Material Type Bend
    elif material_type == "Bend":
        # Organize Data from Data Hub
        return True

    # Material Type Special Piping
    elif material_type == "Special Piping":
        # Organize Data from Data Hub
        DataCollector.data_collector_specialpip_yard(project_number, "Special PIP")
        # Do something with the input and type
        return True

    # All Materials
    elif material_type == "All Materials":
        return True

    return False
