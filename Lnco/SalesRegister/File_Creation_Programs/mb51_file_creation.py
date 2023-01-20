import pandas as pd
import logging


def mb51_file_creation(config_main, mb51_client_dataframe, json_data_list, filtered_mb51_file_saving_path, filtered_mb51_sheet_name):
    try:
        mb51_json_data = json_data_list[0]
        mb51_new_dataframe = pd.DataFrame()

        material_default_name = mb51_json_data['Material']['default_column_name']
        material_client_name = mb51_json_data['Material']['client_column_name']
        mb51_new_dataframe[material_default_name] = mb51_client_dataframe[material_client_name]
        config_main['material_default_name'] = material_default_name
        config_main[material_default_name] = material_client_name

        material_description_default_name = mb51_json_data['Material_description']['default_column_name']
        material_description_client_name = mb51_json_data['Material_description']['client_column_name']
        mb51_new_dataframe[material_description_default_name] = mb51_client_dataframe[material_description_client_name]
        config_main['material_description_default_name'] = material_description_default_name
        config_main[material_description_default_name] = material_description_client_name

        quantity_default_name = mb51_json_data['quantity']['default_column_name']
        quantity_client_name = mb51_json_data['quantity']['client_column_name']
        mb51_new_dataframe[quantity_default_name] = mb51_client_dataframe[
            quantity_client_name]
        config_main['quantity_default_name'] = quantity_default_name
        config_main[quantity_default_name] = quantity_client_name

    except Exception as mb51_json_exception:
        logging.error(
            "Exception occurred while getting column names from the JSON data in 'input file configuration' datatable")
        raise mb51_json_exception
    try:
        mb51_new_dataframe.columns = ["Material", "Material description", "Quantity"]
        mb51_new_dataframe[["Material"]] = mb51_new_dataframe[["Material"]].fillna('').astype(str, errors='ignore')
        mb51_new_dataframe[["Material description"]] = mb51_new_dataframe[["Material description"]].fillna('').astype(str, errors='ignore')
        mb51_new_dataframe[["Quantity"]] = mb51_new_dataframe[["Quantity"]].fillna(0.0).astype(float, errors='ignore')

    except Exception as datatype_conversion_exception:
        logging.error("Exception occurred while converting datatypes of Inventory mapping file")
        raise datatype_conversion_exception

    # create new Excel file in ID folder in Input folder
    try:
        with pd.ExcelWriter(filtered_mb51_file_saving_path, engine="openpyxl") as writer:
            mb51_new_dataframe.to_excel(writer, sheet_name=filtered_mb51_sheet_name, index=False)
            return mb51_new_dataframe
    except Exception as filtered_mb51_error:
        logging.error("Exception occurred while creating filtered purchase register previous quarter file")
        raise filtered_mb51_error


if __name__ == '__main__':
    pass
