import pandas as pd
import logging


def vendor_file_creation(vendor_file_client_dataframe, json_data_list, filtered_vendor_file_saving_path, filtered_vendor_file_sheet_name):
    try:
        vendor_json_data = json_data_list[1]
        vendor_new_dataframe = pd.DataFrame()
        vendor_code_default_name = vendor_json_data['Vendor_Code']['default_column_name']
        vendor_code_client_name = vendor_json_data['Vendor_Code']['client_column_name']
        vendor_new_dataframe[vendor_code_default_name] = vendor_file_client_dataframe[vendor_code_client_name]

        vendor_name_default_name = vendor_json_data['Vendor_Name']['default_column_name']
        vendor_name_client_name = vendor_json_data['Vendor_Name']['client_column_name']
        vendor_new_dataframe[vendor_name_default_name] = vendor_file_client_dataframe[vendor_name_client_name]

        tax_number_default_name = vendor_json_data['Tax_Number']['default_column_name']
        tax_number_client_name = vendor_json_data['Tax_Number']['client_column_name']
        vendor_new_dataframe[tax_number_default_name] = vendor_file_client_dataframe[tax_number_client_name]

    except Exception as vendor_json_exception:
        logging.error(
            "Exception occurred while getting column names from the JSON data in 'input file configuration' datatable")
        logging.exception(vendor_json_exception)
        raise vendor_json_exception

    try:
        vendor_new_dataframe.columns = ['Vendor Code', 'Vendor Name', 'Tax Number']
        vendor_new_dataframe[["Vendor Code"]] = vendor_new_dataframe[["Vendor Code"]].fillna('').astype(int, errors='ignore')
        vendor_new_dataframe[["Vendor Name"]] = vendor_new_dataframe[["Vendor Name"]].fillna('').astype(str, errors='ignore')
        vendor_new_dataframe[['Tax Number']] = vendor_new_dataframe[['Tax Number']].fillna('').astype(str, errors='ignore')

    except Exception as datatype_conversion_exception:
        logging.error("Exception occurred while converting datatypes of vendor file")
        raise datatype_conversion_exception

    # create new Excel file in ID folder in Config folder
    try:
        with pd.ExcelWriter(filtered_vendor_file_saving_path, engine="openpyxl") as writer:
            vendor_new_dataframe.to_excel(writer, sheet_name=filtered_vendor_file_sheet_name, index=False)
            return vendor_new_dataframe
    except Exception as filtered_vendor_file_error:
        logging.error("Exception occurred while creating filtered purchase register previous quarter file")
        logging.exception(filtered_vendor_file_error)
        raise filtered_vendor_file_error


if __name__ == '__main__':
    pass
