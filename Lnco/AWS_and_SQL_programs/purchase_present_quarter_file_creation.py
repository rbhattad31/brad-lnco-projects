import logging
import pandas as pd
import openpyxl


def purchase_present_quarter_file_creation(purchase_present_client_dataframe, json_data_list,
                                           filtered_purchase_present_file_saving_path,
                                           filtered_purchase_present_sheet_name):
    try:
        purchase_columns_json_data = json_data_list[2]
        purchase_present_new_dataframe = pd.DataFrame()

        plant_default_name = purchase_columns_json_data['Plant']['default_column_name']
        plant_client_name = purchase_columns_json_data['Plant']['client_column_name']
        purchase_present_new_dataframe[plant_default_name] = purchase_present_client_dataframe[plant_client_name]

        gr_document_number_default_name = purchase_columns_json_data['GR_Document_Number']['default_column_name']
        gr_document_number_client_name = purchase_columns_json_data['GR_Document_Number']['client_column_name']
        purchase_present_new_dataframe[gr_document_number_default_name] = purchase_present_client_dataframe[
            gr_document_number_client_name]

        gr_posting_date_default_name = purchase_columns_json_data['GR_Posting_Date']['default_column_name']
        gr_posting_date_client_name = purchase_columns_json_data['GR_Posting_Date']['client_column_name']
        purchase_present_new_dataframe[gr_posting_date_default_name] = purchase_present_client_dataframe[
            gr_posting_date_client_name]

        valuation_class_default_name = purchase_columns_json_data['Valuation_Class']['default_column_name']
        valuation_class_client_name = purchase_columns_json_data['Valuation_Class']['client_column_name']
        purchase_present_new_dataframe[valuation_class_default_name] = purchase_present_client_dataframe[
            valuation_class_client_name]

        valuation_class_text_default_name = purchase_columns_json_data['Valuation_Class_Text']['default_column_name']
        valuation_class_text_client_name = purchase_columns_json_data['Valuation_Class_Text']['client_column_name']
        purchase_present_new_dataframe[valuation_class_text_default_name] = purchase_present_client_dataframe[
            valuation_class_text_client_name]
        material_number_default_name = purchase_columns_json_data['Material_Number']['default_column_name']
        material_number_client_name = purchase_columns_json_data['Material_Number']['client_column_name']
        purchase_present_new_dataframe[material_number_default_name] = purchase_present_client_dataframe[
            material_number_client_name]

        material_description_default_name = purchase_columns_json_data['Material_Desc']['default_column_name']
        material_description_client_name = purchase_columns_json_data['Material_Desc']['client_column_name']
        purchase_present_new_dataframe[material_description_default_name] = purchase_present_client_dataframe[
            material_description_client_name]

        vendor_number_default_name = purchase_columns_json_data['Vendor_Number']['default_column_name']
        vendor_number_client_name = purchase_columns_json_data['Vendor_Number']['client_column_name']
        purchase_present_new_dataframe[vendor_number_default_name] = purchase_present_client_dataframe[
            vendor_number_client_name]

        vendor_name_default_name = purchase_columns_json_data['Vendor_Name']['default_column_name']
        vendor_name_client_name = purchase_columns_json_data['Vendor_Name']['client_column_name']
        purchase_present_new_dataframe[vendor_name_default_name] = purchase_present_client_dataframe[
            vendor_name_client_name]

        gr_quantity_default_name = purchase_columns_json_data['GR_Qty']['default_column_name']
        gr_quantity_client_name = purchase_columns_json_data['GR_Qty']['client_column_name']
        purchase_present_new_dataframe[gr_quantity_default_name] = purchase_present_client_dataframe[
            gr_quantity_client_name]

        gr_amount_in_loc_cur_default_name = purchase_columns_json_data['GR_Amt_in_loc_curr']['default_column_name']
        gr_amount_in_loc_cur_client_name = purchase_columns_json_data['GR_Amt_in_loc_curr']['client_column_name']
        purchase_present_new_dataframe[gr_amount_in_loc_cur_default_name] = purchase_present_client_dataframe[
            gr_amount_in_loc_cur_client_name]

        unit_price_default_name = purchase_columns_json_data['Unit_Price']['default_column_name']
        unit_price_client_name = purchase_columns_json_data['Unit_Price']['client_column_name']
        purchase_present_new_dataframe[unit_price_default_name] = purchase_present_client_dataframe[
            unit_price_client_name]

        currency_key_default_name = purchase_columns_json_data['Currency_Key']['default_column_name']
        currency_key_client_name = purchase_columns_json_data['Currency_Key']['client_column_name']
        purchase_present_new_dataframe[currency_key_default_name] = purchase_present_client_dataframe[
            currency_key_client_name]

    except Exception as purchase_json_exception:
        logging.error(
            "Exception occurred while getting column names from the JSON data in 'input file configuration' datatable")
        raise purchase_json_exception

    try:
        purchase_present_new_dataframe.columns = ['Plant', 'GR Document Number', 'GR Posting Date',
                                                  'Valuation Class', 'Valuation Class Text',
                                                  'Material No.',
                                                  'Material Desc', 'Vendor No.', 'Vendor Name', 'GR Qty',
                                                  'GR Amt.in loc.cur.', 'Unit Price', 'Currency Key']

        # Below 4 columns are int datatype and converting them from Object to int datatype.
        # and raise exception if contains any data rather than numbers.
        purchase_present_new_dataframe[["Plant", "GR Document Number", "Valuation Class", "GR Qty"]] = \
            purchase_present_new_dataframe[["Plant", "GR Document Number", "Valuation Class", "GR Qty"]].fillna(
                0).astype(int, errors='raise')

        # Below 4 datatypes are object when read from excel, converting back to String, raises exception if not suitable datatype is found
        purchase_present_new_dataframe[["Valuation Class Text", "Material No.", "Material Desc", "Currency Key"]] = \
            purchase_present_new_dataframe[
                ["Valuation Class Text", "Material No.", "Material Desc", "Currency Key"]].astype(str, errors='raise')

        # purchase_present_new_dataframe[["GR Posting Date"]] = purchase_present_new_dataframe[["GR Posting Date"]].apply(pd.to_datetime)

        # vendor number - int datatype and can have nan values replace them with ''
        purchase_present_new_dataframe[["Vendor No."]] = purchase_present_new_dataframe[["Vendor No."]].fillna(
            '').astype(int, errors='ignore')
        # vendor number - string datatype and can have nan values replace them with ''
        purchase_present_new_dataframe[["Vendor Name"]] = purchase_present_new_dataframe[["Vendor Name"]].fillna(
            '').astype(str, errors='ignore')

        # Gr amount in loc cur & Unit price - float datatype and can have nan values, replace them with 0
        purchase_present_new_dataframe[["GR Amt.in loc.cur.", "Unit Price"]] = purchase_present_new_dataframe[
            ["GR Amt.in loc.cur.", "Unit Price"]].fillna(0).astype(float, errors='ignore')

        # print(purchase_present_new_dataframe.dtypes.tolist())
        logging.info("purchase register previous quarter datatypes are changed successfully ")
        print("purchase register previous quarter datatypes are changed successfully")
    except Exception as datatype_conversion_exception:
        logging.error("Exception occurred while converting datatypes of present quarter purchase register input file")
        raise datatype_conversion_exception

    # create new Excel file in ID folder in Input folder
    try:
        with pd.ExcelWriter(filtered_purchase_present_file_saving_path, engine="openpyxl") as writer:
            purchase_present_new_dataframe.to_excel(writer, sheet_name=filtered_purchase_present_sheet_name,
                                                    index=False)
            return purchase_present_new_dataframe
    except Exception as filtered_purchase_present_error:
        logging.error("Exception occurred while creating filtered purchase register present quarter file")
        raise filtered_purchase_present_error


if __name__ == '__main__':
    pass
