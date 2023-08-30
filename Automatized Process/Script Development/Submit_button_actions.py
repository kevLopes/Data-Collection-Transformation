import EcosysData_API
import CostAnalyzeByTagNumber
import CostAnalyzeProcessByMaterial
import DataCollector
import AnalyzeProcessByYard
import AnalyzeProcessByMTO
import Complete_MTO_Process


#Ecosys API data
def action_for_ecosys_api(project_number):
    try:
        print("Analyze on going. Data are being refreshed from Ecosys")
        # Ecosys PO Lines
        EcosysData_API.ecosys_po_lines_data_api("keven.deOliveiralope", "My-SBM@Codigo23", project_number)
        # Ecosys PO Header
        EcosysData_API.ecosys_poheader_data_api("keven.deOliveiralope", "My-SBM@Codigo23", project_number)
        # Ecosys SUN transactions
        EcosysData_API.ecosys_sun_lines_data_api("keven.deOliveiralope", "My-SBM@Codigo23", project_number)
        # Ecosys etreg transactions
        EcosysData_API.ecosys_etreg_data_api("keven.deOliveiralope", "My-SBM@Codigo23", project_number)
        print("Analyze done for Ecosys API data")
        return True
    except Exception as e:
        print(f"Error occurred during Ecosys API analysis: {str(e)}")
        return False


#SBM Scope Analyze
def action_for_material_analyze(project_number, material_type):
    print("Analyze on going!!! Collecting and Transforming the Data")
    folder_path_valve = "../Data Pool/Material Data Organized/Valve"
    folder_path_piping = "../Data Pool/Material Data Organized/Piping"
    folder_path_bolt = "../Data Pool/Material Data Organized/Bolt"
    folder_path_sp = "../Data Pool/Material Data Organized/Special Piping"
    folder_path_structure = "../Data Pool/Material Data Organized/Structure"

    # Material Type Piping
    if material_type == "Piping":
        # Organize Data from Data Hub
        DataCollector.data_collector_piping(project_number, material_type)
        # Do something with the input and type
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_piping(folder_path_piping, project_number, material_type)
        CostAnalyzeByTagNumber.extract_distinct_tag_numbers_piping(folder_path_piping, project_number, material_type)
        AnalyzeProcessByMTO.analyze_by_material_type_piping(folder_path_piping, project_number, material_type)
        AnalyzeProcessByMTO.sbm_scope_piping_report(folder_path_piping, project_number, material_type)
        print("Analyze done for all SBM Scope Piping Materials")
        return True

    # Material Type Valve
    elif material_type == "Valve":
        # Organize Data from Data Hub
        DataCollector.data_collector_valve(project_number, material_type)
        # Do something with the input and type
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_valve(folder_path_valve, project_number, material_type)
        AnalyzeProcessByMTO.sbm_scope_valve_report(folder_path_valve, project_number, material_type)
        print("Analyze done for all SBM Scope Valve Materials")
        return True

    # Material Type Bolt
    elif material_type == "Bolt":
        # Organize Data from Data Hub
        DataCollector.data_collector_bolt(project_number, material_type)
        # Do something with the input and type
        CostAnalyzeProcessByMaterial.extract_distinct_product_codes_bolt(folder_path_bolt, project_number, material_type)
        AnalyzeProcessByMTO.sbm_scope_bolt_report(folder_path_bolt, project_number, material_type)
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
    elif material_type == "All Equipment":
        # UNITY
        if project_number == "17033":
            # Execute functions for each material type in sequence
            '''try:
                DataCollector.data_collector_piping(project_number, "Piping")
                CostAnalyzeProcessByMaterial.extract_distinct_product_codes_piping(folder_path_piping, project_number, "Piping")
                CostAnalyzeByTagNumber.extract_distinct_tag_numbers_piping(folder_path_piping, project_number, "Piping")
                AnalyzeProcessByMTO.analyze_by_material_type_piping(folder_path_piping, project_number, "Piping")
                AnalyzeProcessByMTO.sbm_scope_piping_report(folder_path_piping, project_number, "Piping")
            except Exception as e:
                print(f"Error occurred during Piping analysis: {str(e)}")

            try:
                DataCollector.data_collector_valve(project_number, "Valve")
                CostAnalyzeProcessByMaterial.extract_distinct_product_codes_valve(folder_path_valve, project_number, "Valve")
                AnalyzeProcessByMTO.sbm_scope_valve_report(folder_path_valve, project_number, "Valve")
            except Exception as e:
                print(f"Error occurred during Valve analysis: {str(e)}")

            try:
                DataCollector.data_collector_bolt(project_number, "Bolt")
                CostAnalyzeProcessByMaterial.extract_distinct_product_codes_bolt(folder_path_bolt, project_number, "Bolt")
                AnalyzeProcessByMTO.sbm_scope_bolt_report(folder_path_bolt, project_number, "Bolt")
            except Exception as e:
                print(f"Error occurred during Bolt analysis: {str(e)}")

            try:
                DataCollector.data_collector_structure(project_number, "Structure")
                CostAnalyzeProcessByMaterial.analyze_structure_materials(folder_path_structure, project_number, "Structure")
            except Exception as e:
                print(f"Error occurred during Structure analysis: {str(e)}")

            try:
                DataCollector.data_collector_bend(project_number, "Bend")
            except Exception as e:
                print(f"Error occurred during Bend analysis: {str(e)}")

            try:
                DataCollector.data_collector_specialpip(project_number, "Special PIP")
                CostAnalyzeByTagNumber.extract_distinct_tag_numbers_special_piping(folder_path_sp, project_number, "SPC Piping")
            except Exception as e:
                print(f"Error occurred during Special Piping analysis: {str(e)}")

            action_for_material_analyze_by_yard(project_number, material_type)'''

            try:
                # Run Complete MTO Process
                Complete_MTO_Process.complete_mto_data_analyze(project_number)
                print("Cost Analyze done based on all distinct Material types")
            except Exception as e:
                print(f"Error occurred during Complete MTO Process: {str(e)}")
        # PROSPERITY
        elif project_number == "17043":
            try:
                # Run Complete MTO Process
                Complete_MTO_Process.complete_mto_data_analyze(project_number)
                print("Cost Analyze done based on all distinct Material types")
            except Exception as e:
                print(f"Error occurred during Complete MTO Process: {str(e)}")

        return True

    return False


#YARD Scope Analyze
def action_for_material_analyze_by_yard(project_number, material_type):
    folder_path_valve = "../Data Pool/Material Data Organized/Valve/Yard"
    folder_path_piping = "../Data Pool/Material Data Organized/Piping/Yard"
    folder_path_bolt = "../Data Pool/Material Data Organized/Bolt/Yard"

    # Material Type Piping
    if material_type == "Piping":
        # Organize Data from Data Hub
        DataCollector.data_collector_piping_yard(project_number, material_type)
        # Do something with the input and type
        AnalyzeProcessByYard.yard_piping_material_type_analyze(folder_path_piping, project_number, material_type)
        AnalyzeProcessByMTO.yard_scope_piping_report(folder_path_piping, project_number, material_type)
        return True

    # Material Type Valve
    elif material_type == "Valve":
        # Organize Data from Data Hub
        DataCollector.data_collector_valve_yard(project_number, material_type)
        # Do something with the input and type
        AnalyzeProcessByYard.yard_valve_material_type_analyze(folder_path_valve, project_number, material_type)
        AnalyzeProcessByMTO.yard_scope_valve_report(folder_path_valve, project_number, material_type)
        return True

    # Material Type Bolt
    elif material_type == "Bolt":
        # Organize Data from Data Hub
        DataCollector.data_collector_bolt_yard(project_number, material_type)
        # Do something with the input and type
        AnalyzeProcessByYard.yard_bolt_material_type_analyze(folder_path_bolt, project_number, material_type)
        AnalyzeProcessByMTO.yard_scope_bolt_report(folder_path_bolt, project_number, material_type)
        return True

    # Material Type Special Piping
    elif material_type == "Special Piping":
        # Organize Data from Data Hub and Graphic Design
        DataCollector.data_collector_specialpip_yard(project_number, "Special PIP")
        return True

    # All Equipment Type
    elif material_type == "All Equipment":
        # Execute functions for each material type in sequence
        try:
            DataCollector.data_collector_piping_yard(project_number, "Piping")
            AnalyzeProcessByYard.yard_piping_material_type_analyze(folder_path_piping, project_number, "Piping")
            AnalyzeProcessByMTO.yard_scope_piping_report(folder_path_piping, project_number, "Piping")
        except Exception as e:
            print(f"Error occurred during YARD Piping analysis: {str(e)}")

        try:
            DataCollector.data_collector_valve_yard(project_number, "Valve")
            AnalyzeProcessByYard.yard_valve_material_type_analyze(folder_path_valve, project_number, "Valve")
            AnalyzeProcessByMTO.yard_scope_valve_report(folder_path_valve, project_number, "Valve")
        except Exception as e:
            print(f"Error occurred during YARD Valve analysis: {str(e)}")

        try:
            DataCollector.data_collector_bolt_yard(project_number, "Bolt")
            AnalyzeProcessByYard.yard_bolt_material_type_analyze(folder_path_bolt, project_number, "Bolt")
            AnalyzeProcessByMTO.yard_scope_bolt_report(folder_path_bolt, project_number, "Bolt")
        except Exception as e:
            print(f"Error occurred during YARD Bolt analysis: {str(e)}")

        try:
            DataCollector.data_collector_specialpip_yard(project_number, "Special PIP")
        except Exception as e:
            print(f"Error occurred during Special Piping analysis: {str(e)}")

        return True

    return False
