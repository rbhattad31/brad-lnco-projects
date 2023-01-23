import pandas as pd
import logging


def open_po_file_creation(config_main, open_po_client_dataframe, json_data_list, filtered_open_po_file_saving_path,
                          filtered_open_po_sheet_name):
    try:
        open_po_json_data = json_data_list[4]
        open_po_new_dataframe = open_po_client_dataframe

        order_date_default_name = open_po_json_data['Order_Date']['default_column_name']
        order_date_client_name = open_po_json_data['Order_Date']['client_column_name']
        # open_po_new_dataframe[order_date_default_name] = open_po_client_dataframe[order_date_client_name]
        config_main['order_date_default_name'] = order_date_default_name
        config_main[order_date_default_name] = order_date_client_name

        po_date_default_name = open_po_json_data['PO_Date']['default_column_name']
        po_date_client_name = open_po_json_data['PO_Date']['client_column_name']
        # open_po_new_dataframe[po_date_default_name] = open_po_client_dataframe[po_date_client_name]
        config_main['po_date_default_name'] = po_date_default_name
        config_main[po_date_default_name] = po_date_client_name

    except Exception as open_po_json_exception:
        logging.error(
            "Exception occurred while getting column names from the JSON data in 'input file configuration' datatable")
        raise open_po_json_exception
    try:
        # open_po_new_dataframe.columns = ['Order Date', 'PO Date']
        open_po_new_dataframe.rename(columns={order_date_client_name: 'Order Date', po_date_client_name: 'PO Date'}, inplace=True)
        open_po_new_dataframe['Order Date'] = pd.to_datetime(open_po_new_dataframe['Order Date'], errors='coerce')
        open_po_new_dataframe['PO Date'] = pd.to_datetime(open_po_new_dataframe['PO Date'], errors='coerce')

    except Exception as datatype_conversion_exception:
        logging.error("Exception occurred while converting datatypes of open po file")
        raise datatype_conversion_exception

    # create new Excel file in ID folder in Input folder
    try:
        with pd.ExcelWriter(filtered_open_po_file_saving_path, engine="openpyxl") as writer:
            open_po_new_dataframe.to_excel(writer, sheet_name=filtered_open_po_sheet_name, index=False)
            return [open_po_new_dataframe, config_main]
    except Exception as filtered_open_po_error:
        logging.error("Exception occurred while creating filtered open po file")
        raise filtered_open_po_error


if __name__ == '__main__':
    pass
