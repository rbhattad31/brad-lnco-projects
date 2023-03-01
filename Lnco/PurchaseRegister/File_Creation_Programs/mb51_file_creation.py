import pandas as pd
import logging
import gc


def mb51_file_creation(config_main, mb51_client_dataframe, json_data_list, filtered_mb51_file_saving_path, filtered_mb51_sheet_name):
    try:
        mb51_json_data = json_data_list[0]
        mb51_new_dataframe = pd.DataFrame()
        material_document_default_name = mb51_json_data['Material_Document']['default_column_name']
        material_document_client_name = mb51_json_data['Material_Document']['client_column_name']
        mb51_new_dataframe[material_document_default_name] = mb51_client_dataframe[material_document_client_name]
        config_main['material_document_default_name'] = material_document_default_name
        config_main[material_document_default_name] = material_document_client_name

        quantity_in_unit_of_entry_default_name = mb51_json_data['Qty_in_unit_of_entry']['default_column_name']
        quantity_in_unit_of_entry_client_name = mb51_json_data['Qty_in_unit_of_entry']['client_column_name']
        mb51_new_dataframe[quantity_in_unit_of_entry_default_name] = mb51_client_dataframe[
            quantity_in_unit_of_entry_client_name]

        config_main['quantity_in_unit_of_entry_default_name'] = quantity_in_unit_of_entry_default_name
        config_main[quantity_in_unit_of_entry_default_name] = quantity_in_unit_of_entry_client_name

        # Movement type
        movement_type_default_name = "Movement type"
        movement_type_client_name = "Movement type"
        mb51_new_dataframe[movement_type_default_name] = mb51_client_dataframe[movement_type_client_name]

        config_main['movement_type_default_name'] = movement_type_default_name
        config_main[movement_type_default_name] = movement_type_client_name

    except Exception as mb51_json_exception:
        logging.error(
            "Exception occurred while getting column names from the JSON data in 'input file configuration' datatable")
        logging.exception(mb51_json_exception)
        raise mb51_json_exception
    else:
        logging.info("Column names extracted for MB51 and created new dataframe with the required columns")
    try:
        mb51_new_dataframe.columns = ["Material Document", "Qty in unit of entry", "Movement type"]
        mb51_new_dataframe[["Material Document"]] = mb51_new_dataframe[["Material Document"]].fillna(0).astype(int, errors='ignore')
        mb51_new_dataframe[["Qty in unit of entry"]] = mb51_new_dataframe[["Qty in unit of entry"]].fillna(0).astype(float, errors='ignore')
        mb51_new_dataframe[["Movement type"]] = mb51_new_dataframe[["Movement type"]].fillna(0).astype(int, errors='ignore')

    except Exception as datatype_conversion_exception:
        logging.error("Exception occurred while converting datatypes of Inventory mapping file")
        raise datatype_conversion_exception
    else:
        logging.info("Changed Datatypes of Columns in MB51")
    # create new Excel file in ID folder in Input folder
    try:
        logging.info("saving Filtered MB51 file in input folder of the request")
        print("saving Filtered MB51 file in output folder of the request")
        gc.collect()
        print("Garbage collection done for saving file for clearing up space.")

        with pd.ExcelWriter(filtered_mb51_file_saving_path, engine="openpyxl") as writer:
            mb51_new_dataframe.to_excel(writer, sheet_name=filtered_mb51_sheet_name, index=False)

    except Exception as filtered_mb51_error:
        logging.error("Exception occurred while creating filtered purchase register previous quarter file")
        raise filtered_mb51_error
    else:
        logging.info("Successfully saved Filtered MB51 file in input folder of the request")
        print("Successfully saved Filtered MB51 file in input folder of the request")
        return [mb51_new_dataframe, config_main]


if __name__ == '__main__':
    pass
