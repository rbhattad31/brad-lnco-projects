import logging
import pandas as pd


def purchase_previous_quarter_file_creation(purchase_previous_client_dataframe, json_data_list, filtered_purchase_previous_file_saving_path, filtered_purchase_previous_sheet_name):
    try:
        purchase_columns_json_data = json_data_list[2]
        purchase_previous_new_dataframe = pd.DataFrame()

        plant_default_name = purchase_columns_json_data['Plant']['default_column_name']
        plant_client_name = purchase_columns_json_data['Plant']['client_column_name']
        purchase_previous_new_dataframe[plant_default_name] = purchase_previous_client_dataframe[plant_client_name]

        gr_document_number_default_name = purchase_columns_json_data['GR_Document_Number']['default_column_name']
        gr_document_number_client_name = purchase_columns_json_data['GR_Document_Number']['client_column_name']
        purchase_previous_new_dataframe[gr_document_number_default_name] = purchase_previous_client_dataframe[
            gr_document_number_client_name]

        gr_posting_date_default_name = purchase_columns_json_data['GR_Posting_Date']['default_column_name']
        gr_posting_date_client_name = purchase_columns_json_data['GR_Posting_Date']['client_column_name']
        purchase_previous_new_dataframe[gr_posting_date_default_name] = purchase_previous_client_dataframe[
            gr_posting_date_client_name]

        valuation_class_default_name = purchase_columns_json_data['Valuation_Class']['default_column_name']
        valuation_class_client_name = purchase_columns_json_data['Valuation_Class']['client_column_name']
        purchase_previous_new_dataframe[valuation_class_default_name] = purchase_previous_client_dataframe[
            valuation_class_client_name]

        valuation_class_text_default_name = purchase_columns_json_data['Valuation_Class_Text']['default_column_name']
        valuation_class_text_client_name = purchase_columns_json_data['Valuation_Class_Text']['client_column_name']
        purchase_previous_new_dataframe[valuation_class_text_default_name] = purchase_previous_client_dataframe[
            valuation_class_text_client_name]
        material_number_default_name = purchase_columns_json_data['Material_Number']['default_column_name']
        material_number_client_name = purchase_columns_json_data['Material_Number']['client_column_name']
        purchase_previous_new_dataframe[material_number_default_name] = purchase_previous_client_dataframe[
            material_number_client_name]

        material_description_default_name = purchase_columns_json_data['Material_Desc']['default_column_name']
        material_description_client_name = purchase_columns_json_data['Material_Desc']['client_column_name']
        purchase_previous_new_dataframe[material_description_default_name] = purchase_previous_client_dataframe[
            material_description_client_name]

        vendor_number_default_name = purchase_columns_json_data['Vendor_Number']['default_column_name']
        vendor_number_client_name = purchase_columns_json_data['Vendor_Number']['client_column_name']
        purchase_previous_new_dataframe[vendor_number_default_name] = purchase_previous_client_dataframe[
            vendor_number_client_name]

        vendor_name_default_name = purchase_columns_json_data['Vendor_Name']['default_column_name']
        vendor_name_client_name = purchase_columns_json_data['Vendor_Name']['client_column_name']
        purchase_previous_new_dataframe[vendor_name_default_name] = purchase_previous_client_dataframe[
            vendor_name_client_name]

        gr_quantity_default_name = purchase_columns_json_data['GR_Qty']['default_column_name']
        gr_quantity_client_name = purchase_columns_json_data['GR_Qty']['client_column_name']
        purchase_previous_new_dataframe[gr_quantity_default_name] = purchase_previous_client_dataframe[
            gr_quantity_client_name]

        gr_amount_in_loc_cur_default_name = purchase_columns_json_data['GR_Amt_in_loc_curr']['default_column_name']
        gr_amount_in_loc_cur_client_name = purchase_columns_json_data['GR_Amt_in_loc_curr']['client_column_name']
        purchase_previous_new_dataframe[gr_amount_in_loc_cur_default_name] = purchase_previous_client_dataframe[
            gr_amount_in_loc_cur_client_name]

        unit_price_default_name = purchase_columns_json_data['Unit_Price']['default_column_name']
        unit_price_client_name = purchase_columns_json_data['Unit_Price']['client_column_name']
        purchase_previous_new_dataframe[unit_price_default_name] = purchase_previous_client_dataframe[
            unit_price_client_name]

        currency_key_default_name = purchase_columns_json_data['Currency_Key']['default_column_name']
        currency_key_client_name = purchase_columns_json_data['Currency_Key']['client_column_name']
        purchase_previous_new_dataframe[currency_key_default_name] = purchase_previous_client_dataframe[
            currency_key_client_name]
    except Exception as purchase_json_exception:
        logging.error(
            "Exception occurred while getting column names from the JSON data in 'input file configuration' datatable")
        raise purchase_json_exception

    # create new Excel file in ID folder in Input folder
    try:
        with pd.ExcelWriter(filtered_purchase_previous_file_saving_path, engine="openpyxl") as writer:
            purchase_previous_new_dataframe.to_excel(writer, sheet_name=filtered_purchase_previous_sheet_name,
                                                     index=False)
            return purchase_previous_new_dataframe
    except Exception as filtered_purchase_previous_error:
        logging.error("Exception occurred while creating filtered purchase register previous quarter file")
        raise filtered_purchase_previous_error


if __name__ == '__main__':
    pass
