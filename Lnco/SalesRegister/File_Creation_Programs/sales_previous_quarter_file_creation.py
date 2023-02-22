import logging
import pandas as pd


def sales_previous_quarter_file_creation(config_main, sales_previous_client_dataframe, json_data_list,
                                         filtered_sales_previous_file_saving_path,
                                         filtered_sales_previous_sheet_name):
    try:
        # read sales register json data from the list
        sales_columns_json_data = json_data_list[2]
        sales_previous_new_dataframe = pd.DataFrame()

        billing_date_default_name = sales_columns_json_data['Billing_Date']['default_column_name']
        billing_date_client_name = sales_columns_json_data['Billing_Date']['client_column_name']
        sales_previous_new_dataframe[billing_date_default_name] = sales_previous_client_dataframe[
            billing_date_client_name]
        config_main['billing_date_default_name'] = billing_date_default_name
        config_main[billing_date_default_name] = billing_date_client_name

        doc_type_text_default_name = sales_columns_json_data['Doc_Type_Text']['default_column_name']
        doc_type_text_client_name = sales_columns_json_data['Doc_Type_Text']['client_column_name']
        sales_previous_new_dataframe[doc_type_text_default_name] = sales_previous_client_dataframe[
            doc_type_text_client_name]
        config_main['doc_type_text_default_name'] = doc_type_text_default_name
        config_main[doc_type_text_default_name] = doc_type_text_client_name

        plant_default_name = sales_columns_json_data['Plant']['default_column_name']
        plant_client_name = sales_columns_json_data['Plant']['client_column_name']
        sales_previous_new_dataframe[plant_default_name] = sales_previous_client_dataframe[
            plant_client_name]
        config_main['plant_default_name'] = plant_default_name
        config_main[plant_default_name] = plant_client_name

        base_price_in_inr_default_name = sales_columns_json_data['Base_Price_in_INR']['default_column_name']
        base_price_in_inr_client_name = sales_columns_json_data['Base_Price_in_INR']['client_column_name']
        sales_previous_new_dataframe[base_price_in_inr_default_name] = sales_previous_client_dataframe[
            base_price_in_inr_client_name]
        config_main['base_price_in_inr_default_name'] = base_price_in_inr_default_name
        config_main[base_price_in_inr_default_name] = base_price_in_inr_client_name

        payer_name_default_name = sales_columns_json_data['Payer_Name']['default_column_name']
        payer_name_client_name = sales_columns_json_data['Payer_Name']['client_column_name']
        sales_previous_new_dataframe[payer_name_default_name] = sales_previous_client_dataframe[
            payer_name_client_name]
        config_main['payer_name_default_name'] = payer_name_default_name
        config_main[payer_name_default_name] = payer_name_client_name

        material_number_default_name = sales_columns_json_data['Material_Number']['default_column_name']
        material_number_client_name = sales_columns_json_data['Material_Number']['client_column_name']
        sales_previous_new_dataframe[material_number_default_name] = sales_previous_client_dataframe[
            material_number_client_name]
        config_main['material_number_default_name'] = material_number_default_name
        config_main[material_number_default_name] = material_number_client_name

        material_description_default_name = sales_columns_json_data['Material_Description']['default_column_name']
        material_description_client_name = sales_columns_json_data['Material_Description']['client_column_name']
        sales_previous_new_dataframe[material_description_default_name] = sales_previous_client_dataframe[
            material_description_client_name]
        config_main['material_description_default_name'] = material_description_default_name
        config_main[material_description_default_name] = material_description_client_name

        billing_qty_default_name = sales_columns_json_data['Billing_Qty']['default_column_name']
        billing_qty_client_name = sales_columns_json_data['Billing_Qty']['client_column_name']
        sales_previous_new_dataframe[billing_qty_default_name] = sales_previous_client_dataframe[
            billing_qty_client_name]
        config_main['billing_qty_default_name'] = billing_qty_default_name
        config_main[billing_qty_default_name] = billing_qty_client_name

        # product_type_descp_default_name = sales_columns_json_data['Product_Type_Descp']['default_column_name']
        # product_type_descp_client_name = sales_columns_json_data['Product_Type_Descp']['client_column_name']
        # sales_previous_new_dataframe[product_type_descp_default_name] = sales_previous_client_dataframe[
        #     product_type_descp_client_name]
        # config_main['product_type_descp_default_name'] = product_type_descp_default_name
        # config_main[product_type_descp_default_name] = product_type_descp_client_name

        material_type_descp_default_name = sales_columns_json_data['Material_Type_Description']['default_column_name']
        material_type_descp_client_name = sales_columns_json_data['Material_Type_Description']['client_column_name']
        sales_previous_new_dataframe[material_type_descp_default_name] = sales_previous_client_dataframe[
            material_type_descp_client_name]
        config_main['material_type_descp_default_name'] = material_type_descp_default_name
        config_main[material_type_descp_default_name] = material_type_descp_client_name

        payer_default_name = sales_columns_json_data['Payer']['default_column_name']
        payer_client_name = sales_columns_json_data['Payer']['client_column_name']
        sales_previous_new_dataframe[payer_default_name] = sales_previous_client_dataframe[
            payer_client_name]
        config_main['payer_default_name'] = payer_default_name
        config_main[payer_default_name] = payer_client_name

        ref_doc_no_default_name = sales_columns_json_data['Ref_Doc_No']['default_column_name']
        ref_doc_no_client_name = sales_columns_json_data['Ref_Doc_No']['client_column_name']
        sales_previous_new_dataframe[ref_doc_no_default_name] = sales_previous_client_dataframe[
            ref_doc_no_client_name]
        config_main['ref_doc_no_default_name'] = ref_doc_no_default_name
        config_main[ref_doc_no_default_name] = ref_doc_no_client_name

        cgst_value_default_name = sales_columns_json_data['CGST_Value']['default_column_name']
        cgst_value_client_name = sales_columns_json_data['CGST_Value']['client_column_name']
        sales_previous_new_dataframe[cgst_value_default_name] = sales_previous_client_dataframe[
            cgst_value_client_name]
        config_main['cgst_value_default_name'] = cgst_value_default_name
        config_main[cgst_value_default_name] = cgst_value_client_name

        sgst_value_default_name = sales_columns_json_data['SGST_Value']['default_column_name']
        sgst_value_client_name = sales_columns_json_data['SGST_Value']['client_column_name']
        sales_previous_new_dataframe[sgst_value_default_name] = sales_previous_client_dataframe[
            sgst_value_client_name]
        config_main['sgst_value_default_name'] = sgst_value_default_name
        config_main[sgst_value_default_name] = sgst_value_client_name

        igst_value_default_name = sales_columns_json_data['IGST_Value']['default_column_name']
        igst_value_client_name = sales_columns_json_data['IGST_Value']['client_column_name']
        sales_previous_new_dataframe[igst_value_default_name] = sales_previous_client_dataframe[
            igst_value_client_name]
        config_main['igst_value_default_name'] = igst_value_default_name
        config_main[igst_value_default_name] = igst_value_client_name

        jtcs_value_default_name = sales_columns_json_data['JTCS_Value']['default_column_name']
        jtcs_value_client_name = sales_columns_json_data['JTCS_Value']['client_column_name']
        sales_previous_new_dataframe[jtcs_value_default_name] = sales_previous_client_dataframe[
            jtcs_value_client_name]
        config_main['jtcs_value_default_name'] = jtcs_value_default_name
        config_main[jtcs_value_default_name] = jtcs_value_client_name

        grand_total_value_default_name = sales_columns_json_data['Grand_Total_Value']['default_column_name']
        grand_total_value_client_name = sales_columns_json_data['Grand_Total_Value']['client_column_name']
        sales_previous_new_dataframe[grand_total_value_default_name] = sales_previous_client_dataframe[
            grand_total_value_client_name]
        config_main['grand_total_value_default_name'] = grand_total_value_default_name
        config_main[grand_total_value_default_name] = grand_total_value_client_name

        hsn_code_default_name = sales_columns_json_data['HSN_Code']['default_column_name']
        hsn_code_client_name = sales_columns_json_data['HSN_Code']['client_column_name']
        sales_previous_new_dataframe[hsn_code_default_name] = sales_previous_client_dataframe[
            hsn_code_client_name]
        config_main['hsn_code_default_name'] = hsn_code_default_name
        config_main[hsn_code_default_name] = hsn_code_client_name

        sales_order_default_name = sales_columns_json_data['Sales_Order']['default_column_name']
        sales_order_client_name = sales_columns_json_data['Sales_Order']['client_column_name']
        sales_previous_new_dataframe[sales_order_default_name] = sales_previous_client_dataframe[
            sales_order_client_name]
        config_main['sales_order_default_name'] = sales_order_default_name
        config_main[sales_order_default_name] = sales_order_client_name

        delivery_number_default_name = sales_columns_json_data['Delivery_No']['default_column_name']
        delivery_number_client_name = sales_columns_json_data['Delivery_No']['client_column_name']
        sales_previous_new_dataframe[delivery_number_default_name] = sales_previous_client_dataframe[
            delivery_number_client_name]
        config_main['delivery_number_default_name'] = delivery_number_default_name
        config_main[delivery_number_default_name] = delivery_number_client_name

        billing_number_default_name = sales_columns_json_data['Billing_No']['default_column_name']
        billing_number_client_name = sales_columns_json_data['Billing_No']['client_column_name']
        sales_previous_new_dataframe[billing_number_default_name] = sales_previous_client_dataframe[
            billing_number_client_name]
        config_main['billing_number_default_name'] = billing_number_default_name
        config_main[billing_number_default_name] = billing_number_client_name

        po_number_default_name = sales_columns_json_data['PO_No']['default_column_name']
        po_number_client_name = sales_columns_json_data['PO_No']['client_column_name']
        sales_previous_new_dataframe[po_number_default_name] = sales_previous_client_dataframe[
            po_number_client_name]
        config_main['po_number_default_name'] = po_number_default_name
        config_main[po_number_default_name] = po_number_client_name

        po_date_default_name = sales_columns_json_data['PO_Date']['default_column_name']
        po_date_client_name = sales_columns_json_data['PO_Date']['client_column_name']
        sales_previous_new_dataframe[po_date_default_name] = sales_previous_client_dataframe[
            po_date_client_name]
        config_main['po_date_default_name'] = po_date_default_name
        config_main[po_date_default_name] = po_date_client_name

        so_unit_price_default_name = sales_columns_json_data['SO_Unit_Price']['default_column_name']
        so_unit_price_client_name = sales_columns_json_data['SO_Unit_Price']['client_column_name']
        sales_previous_new_dataframe[so_unit_price_default_name] = sales_previous_client_dataframe[
            so_unit_price_client_name]
        config_main['so_unit_price_default_name'] = so_unit_price_default_name
        config_main[po_date_default_name] = so_unit_price_client_name

    except Exception as sales_json_exception:
        logging.error(
            "Exception occurred while getting column names from the JSON data in 'input file configuration' datatable")
        raise sales_json_exception

    try:
        sales_previous_new_dataframe.columns = ['Billing Date', 'Doc. Type Text', 'Plant',
                                               'Base Price in INR', 'Payer Name',
                                               'Material No.',
                                               'Material Description', 'Billing Qty.', 'Material Type Descri', 'Payer',
                                               'Ref.Doc.No.', 'CGST Value', 'SGST Value', 'IGST Value', 'JTCS Value',
                                               'Grand Total Value(IN', 'HSN Code', 'Sales Order',
                                               'Delivery No.', 'Billing No.',
                                               'PO. No.', 'PO Date', 'So Unit Price'
                                                ]

        # change datatype of billing date column
        sales_previous_new_dataframe['Billing Date'] = pd.to_datetime(sales_previous_new_dataframe['Billing Date'],
                                                                      errors='coerce')
        sales_previous_new_dataframe['PO Date'] = pd.to_datetime(sales_previous_new_dataframe['PO Date'],
                                                                 errors='coerce')
        # create month column
        sales_previous_new_dataframe['Month'] = sales_previous_new_dataframe['Billing Date'].dt.month_name().str[:3]

        # create type of sale column with values
        sales_previous_new_dataframe['Type of sale'] = ''
        sales_previous_new_dataframe.loc[
            (sales_previous_new_dataframe['Doc. Type Text'].str.lower() == 'Export Order'.lower()) | (
                    sales_previous_new_dataframe[
                        'Doc. Type Text'].str.lower() == 'Export Ordr w/o Duty'.lower()), 'Type of sale'] = 'Export sales'
        sales_previous_new_dataframe.loc[sales_previous_new_dataframe[
                                            'Doc. Type Text'].str.lower() == 'Scrap Order'.lower(), 'Type of sale'] = 'Scrap sales'
        sales_previous_new_dataframe.loc[
            (sales_previous_new_dataframe['Doc. Type Text'].str.lower() == 'Service Order'.lower()) | (
                    sales_previous_new_dataframe['Doc. Type Text'].str.lower() == 'SEZ Sales order'.lower()) | (
                    sales_previous_new_dataframe['Doc. Type Text'].str.lower() == 'Standard Order'.lower()) | (
                    sales_previous_new_dataframe[
                        'Doc. Type Text'].str.lower() == 'Trade Order'.lower()), 'Type of sale'] = 'Domestic sales'
        sales_previous_new_dataframe.loc[sales_previous_new_dataframe[
                                            'Doc. Type Text'].str.lower() == 'Asset Sale Order'.lower(), 'Type of sale'] = 'Sale of asset'
        sales_previous_new_dataframe.loc[sales_previous_new_dataframe[
                                            'Doc. Type Text'].str.lower() == 'INTER PLANT SERVICES'.lower(), 'Type of sale'] = 'Job work services'
        sales_previous_new_dataframe.loc[
            (sales_previous_new_dataframe['Doc. Type Text'].str.lower() == 'Returns'.lower()) | (
                    sales_previous_new_dataframe[
                        'Doc. Type Text'].str.lower() == 'PLL credit memo req'.lower()), 'Type of sale'] = 'Sales return'
        sales_previous_new_dataframe.loc[sales_previous_new_dataframe[
                                            'Doc. Type Text'].str.lower() == 'Debit memo request'.lower(), 'Type of sale'] = 'Debit memo'

        print(list(sales_previous_new_dataframe.columns))
        # Below 4 columns are int datatype and converting them from Object to int datatype.
        # and raise exception if contains any data rather than numbers.
        sales_previous_new_dataframe[
            ["Plant", "Payer", "HSN Code", "Sales Order", "Delivery No.",
             "Billing No."]] = \
            sales_previous_new_dataframe[
                ["Plant", "Payer", "HSN Code", "Sales Order", "Delivery No.",
                 "Billing No."]].fillna(
                0).astype(int, errors='raise')

        # Below 4 datatypes are object when read from excel, converting back to String, raises exception if not suitable datatype is found
        sales_previous_new_dataframe[
            ["Doc. Type Text", "Payer Name", "Material No.", "Material Description", "Material Type Descri",
             "Ref.Doc.No.", "PO. No."]] = \
            sales_previous_new_dataframe[
                ["Doc. Type Text", "Payer Name", "Material No.", "Material Description", "Material Type Descri",
                 "Ref.Doc.No.", "PO. No."]].astype(str, errors='raise')

        # Gr amount in loc cur & Unit price - float datatype and can have nan values, replace them with 0
        sales_previous_new_dataframe[
            ["Base Price in INR", "Billing Qty.", "CGST Value", "SGST Value", "IGST Value", "JTCS Value",
             "Grand Total Value(IN", "So Unit Price"]] = sales_previous_new_dataframe[
            ["Base Price in INR", "Billing Qty.", "CGST Value", "SGST Value", "IGST Value", "JTCS Value",
             "Grand Total Value(IN", "So Unit Price"]].fillna(0).astype(float, errors='ignore')

        logging.info("sales register previous quarter datatypes are changed successfully ")
        print("Sales register previous quarter datatypes are changed successfully")
    except Exception as datatype_conversion_exception:
        logging.error("Exception occurred while converting datatypes of previous quarter sales register input file")
        raise datatype_conversion_exception

    # create new Excel file in ID folder in Input folder
    try:
        with pd.ExcelWriter(filtered_sales_previous_file_saving_path, engine="openpyxl") as writer:
            sales_previous_new_dataframe.to_excel(writer, sheet_name=filtered_sales_previous_sheet_name,
                                                  index=False)
            return [sales_previous_new_dataframe, config_main]
    except Exception as filtered_sales_previous_error:
        logging.error("Exception occurred while creating filtered sales register previous quarter file")
        raise filtered_sales_previous_error


if __name__ == '__main__':
    pass
