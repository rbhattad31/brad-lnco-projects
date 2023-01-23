import pandas as pd
import logging


def hsn_codes_file_creation(config_main, hsn_codes_file_dataframe, json_data_list, filtered_hsn_code_file_saving_path, filtered_hsn_codes_file_sheet_name):
    try:
        hsn_codes_json_data = json_data_list[1]
        hsn_codes_new_dataframe = pd.DataFrame()
        hsn_codes_default_name = hsn_codes_json_data['HSN_Codes']['default_column_name']
        hsn_codes_client_name = hsn_codes_json_data['HSN_Codes']['client_column_name']
        hsn_codes_new_dataframe[hsn_codes_default_name] = hsn_codes_file_dataframe[hsn_codes_client_name]
        config_main['hsn_codes_default_name'] = hsn_codes_default_name
        config_main[hsn_codes_default_name] = hsn_codes_client_name

        gst_rate_default_name = hsn_codes_json_data['GST_Rate']['default_column_name']
        gst_rate_client_name = hsn_codes_json_data['GST_Rate']['client_column_name']
        hsn_codes_new_dataframe[gst_rate_default_name] = hsn_codes_file_dataframe[gst_rate_client_name]
        config_main['gst_rate_default_name'] = gst_rate_default_name
        config_main[gst_rate_default_name] = gst_rate_client_name

    except Exception as hsn_codes_json_exception:
        logging.error(
            "Exception occurred while getting column names from the JSON data in 'input file configuration' datatable")
        logging.exception(hsn_codes_json_exception)
        raise hsn_codes_json_exception

    try:
        hsn_codes_new_dataframe.columns = ['HSN Codes', 'GST Rate']
        hsn_codes_new_dataframe[["HSN Codes"]] = hsn_codes_new_dataframe[["HSN Codes"]].fillna('').astype(int, errors='ignore')
        # print("Check GST Rate Column")
        # print(hsn_codes_new_dataframe)
        hsn_codes_new_dataframe[["GST Rate"]] = hsn_codes_new_dataframe[["GST Rate"]].fillna('').astype(float, errors='ignore')
        # print(hsn_codes_new_dataframe)
    except Exception as datatype_conversion_exception:
        logging.error("Exception occurred while converting datatypes of vendor file")
        raise datatype_conversion_exception

    # create new Excel file in ID folder in Input folder
    try:
        with pd.ExcelWriter(filtered_hsn_code_file_saving_path, engine="openpyxl") as writer:
            hsn_codes_new_dataframe.to_excel(writer, sheet_name=filtered_hsn_codes_file_sheet_name, index=False)
            return [hsn_codes_new_dataframe, config_main]
    except Exception as filtered_hsn_code_file_error:
        logging.error("Exception occurred while creating filtered hsn codes file")
        logging.exception(filtered_hsn_code_file_error)
        raise filtered_hsn_code_file_error


if __name__ == '__main__':
    pass
