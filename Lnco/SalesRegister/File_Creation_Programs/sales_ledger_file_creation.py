import pandas as pd
import logging


def sales_ledger_file_creation(config_main, sales_ledger_client_dataframe, json_data_list, filtered_sales_ledger_file_saving_path, filtered_sales_ledger_sheet_name):
    try:
        sales_ledger_json_data = json_data_list[0]
        sales_ledger_new_dataframe = pd.DataFrame()

        credit_default_name = sales_ledger_json_data['Credit']['default_column_name']
        credit_client_name = sales_ledger_json_data['Credit']['client_column_name']
        sales_ledger_new_dataframe[credit_default_name] = sales_ledger_client_dataframe[credit_client_name]
        config_main['credit_default_name'] = credit_default_name
        config_main[credit_default_name] = credit_client_name

        debit_default_name = sales_ledger_json_data['Debit']['default_column_name']
        debit_client_name = sales_ledger_json_data['Debit']['client_column_name']
        sales_ledger_new_dataframe[debit_default_name] = sales_ledger_client_dataframe[debit_client_name]
        config_main['debit_default_name'] = debit_default_name
        config_main[debit_default_name] = debit_client_name

    except Exception as sales_ledger_json_exception:
        logging.error(
            "Exception occurred while getting column names from the JSON data in 'input file configuration' datatable")
        raise sales_ledger_json_exception
    try:
        sales_ledger_new_dataframe.columns = ['Credit', 'Debit']
        sales_ledger_new_dataframe[['Credit', 'Debit']] = sales_ledger_new_dataframe[['Credit', 'Debit']].fillna(0.0).astype(float, errors='ignore')

    except Exception as datatype_conversion_exception:
        logging.error("Exception occurred while converting datatypes of sales ledger file")
        raise datatype_conversion_exception

    # create new Excel file in ID folder in Input folder
    try:
        with pd.ExcelWriter(filtered_sales_ledger_file_saving_path, engine="openpyxl") as writer:
            sales_ledger_new_dataframe.to_excel(writer, sheet_name=filtered_sales_ledger_sheet_name, index=False)
            return [sales_ledger_new_dataframe, config_main]
    except Exception as filtered_sales_ledger_error:
        logging.error("Exception occurred while creating filtered sales ledger file")
        raise filtered_sales_ledger_error


if __name__ == '__main__':
    pass
