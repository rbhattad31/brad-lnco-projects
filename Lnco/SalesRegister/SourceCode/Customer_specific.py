from string import ascii_lowercase
from openpyxl.styles import Font, PatternFill, Side, Border
import openpyxl
import logging
import pandas as pd


class customer_specific_exception(Exception):
    pass


def customer_specific(sales_register_df, dict_config_main):
    try:
        sales_register_filtered_df = sales_register_df[["Material No.", "Material Description", "Payer Name"]]
        # print(sales_register_filtered_df)
    except Exception as filtered_column_name_error:
        str_exception_message = filtered_column_name_error
        logging.error("Exception occurred while specified filtered was not found in the input list")
        raise customer_specific_exception(str_exception_message)

    try:
        sr_duplicates_filtered_df = sales_register_filtered_df.drop_duplicates(keep='first')
        duplicate = sr_duplicates_filtered_df[
            sr_duplicates_filtered_df.duplicated(subset=["Material Description"], keep=False)]
        # print(duplicate)
    except KeyError:
        str_exception_message = "material Description column was not found"
        logging.error("Exception occurred while specified Material description was not found in the input list")
        raise customer_specific_exception(str_exception_message)
    try:
        output_file_path = dict_config_main["Output_File_Path"]
        output_sheet_name = dict_config_main["Output_Customer_Specific_sheetname"]

        # duplicate.to_excel(output_file_path, sheet_name=output_sheet_name, index=False)
        with pd.ExcelWriter(output_file_path, engine="openpyxl", mode="a",
                            if_sheet_exists="replace") as writer:
            duplicate.to_excel(writer, sheet_name=output_sheet_name, index=False)

        workbook = openpyxl.load_workbook(output_file_path)
        worksheet = workbook[output_sheet_name]

        format_font = Font(name="Calibri", size=11, color="000000", bold=True)
        for c in ascii_lowercase:
            worksheet[c + "1"].font = format_font
        #  Header Fill
        format_fill = PatternFill(patternType='solid', fgColor='ADD8E6')

        for c in ascii_lowercase:
            worksheet[c + "1"].fill = format_fill
            if c == 'c':
                break

        thin = Side(border_style="thin", color='000000')
        for row in worksheet.iter_rows(min_row=1, min_col=1, max_row=worksheet.max_row, max_col=3):
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

        for c in ascii_lowercase:
            worksheet.column_dimensions[c].width = 45
        print(workbook.sheetnames)
        workbook.save(output_file_path)
        print(output_file_path)
        return output_file_path

    except Exception as exception:
        print(exception)
        pass


if __name__ == '__main__':
    pass
