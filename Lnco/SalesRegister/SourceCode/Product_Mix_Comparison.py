import logging
import pandas as pd
import numpy
import openpyxl
from openpyxl.styles import Font, PatternFill
from string import ascii_uppercase
from openpyxl.utils import get_column_letter
import os
from ReusableTasks.send_mail_reusable_task import send_mail


class BusinessException(Exception):
    pass


def product_mix_comparison(main_config, in_config, present_quarter_pd):
    try:
        logging.info("Starting product mix comparison code execution")
        # Read Purchase Register Sheets
        read_present_quarter_pd = present_quarter_pd

        # Fetch To Address
        to_address = main_config["To_Mail_Address"]
        cc_address = main_config["CC_Mail_Address"]

        # Check Exception
        if read_present_quarter_pd.shape[0] == 0:
            subject = in_config["EmptyInput_Subject"]
            body = in_config["EmptyInput_Body"]
            send_mail(to=to_address, cc=cc_address, subject=subject, body=body)
            logging.error("Empty present quarter Sales Register found")
            raise BusinessException("Sheet is empty")

        # Check Column Present
        present_quarter_col = read_present_quarter_pd.columns.values.tolist()
        for col in ["Product Type Descp.", "Billing Qty.", "Base Price in INR"]:
            if col not in present_quarter_col:
                subject = in_config["ColumnMiss_Subject"]
                body = in_config["ColumnMiss_Body"]
                body = body.replace("ColumnName +", col)
                send_mail(to=to_address, cc=cc_address, subject=subject, body=body)
                logging.error("{} Column is missing".format(col))
                raise BusinessException(col + " Column is missing")

        # Filter rows
        product_type_descp = read_present_quarter_pd[read_present_quarter_pd['Product Type Descp.'].notna()]
        billing_qty = read_present_quarter_pd[read_present_quarter_pd['Billing Qty.'].notna()]
        price_inr = read_present_quarter_pd[read_present_quarter_pd['Base Price in INR'].notna()]

        if len(product_type_descp) == 0:
            subject = in_config["Product_type_descp_Subject"]
            body = in_config["Product_type_descp_Body"]
            send_mail(to=to_address, cc=cc_address, subject=subject, body=body)
            logging.error("Product type descp Column is empty")
            raise BusinessException("Product type descp Column is empty")
        elif len(billing_qty) == 0:
            subject = in_config["Billing_Qty_Subject"]
            body = in_config["Billing_Qty_Body"]
            send_mail(to=to_address, cc=cc_address, subject=subject, body=body)
            logging.error("Billing Qty Column is empty")
            raise BusinessException("Billing Qty Column is empty")
        elif len(price_inr) == 0:
            subject = in_config["Base_Price_INR_Subject"]
            body = in_config["Base_Price_INR_Body"]
            send_mail(to=to_address, cc=cc_address, subject=subject, body=body)
            logging.error("Base Price in INR Column is empty")
            raise BusinessException("Base Price in INR Column is empty")
        else:
            pass

        # Create Pivot Table Sales Register
        try:
            read_present_quarter_pd[["Product Type Descp."]] = read_present_quarter_pd[["Product Type Descp."]].fillna('')
            pivot_index = ["Product Type Descp."]
            pivot_values = ["Billing Qty.", "Base Price in INR"]
            product_mix_pivot_df = pd.pivot_table(read_present_quarter_pd, index=pivot_index, values=pivot_values,
                                                  aggfunc=numpy.sum,
                                                  margins=False,
                                                  margins_name="Grand Total")

            print(" Product mix Comparison Pivot table is Created")
            logging.info("Product mix Comparison Pivot table is Created")
        except Exception as create_pivot_table:
            subject = in_config["subject_pivot_table"]
            body = in_config["body_pivot_table"]
            send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
            print("Product mix Comparison Wise Process-", str(create_pivot_table))
            logging.info("Product mix Comparison pivot table is not created")
            raise create_pivot_table

        # Remove Index
        product_mix_pivot_df = product_mix_pivot_df.reset_index()

        # Assign Pivot Sheets
        product_mix_df = product_mix_pivot_df

        # Remove Empty Rows
        product_mix_df = product_mix_df.replace(numpy.nan, 0, regex=True)

        # Get Pivot Column Names
        col_name = product_mix_df.columns.values.tolist()

        # Change Column names
        product_mix_df = product_mix_df.rename(columns={col_name[1]: "Sum Of Billing Qty"})
        product_mix_df = product_mix_df.rename(columns={col_name[2]: "Sum of Base price in INR"})
        product_mix_df = product_mix_df.iloc[:, [0, 2, 1]]
        try:
            # Log Sheet
            with pd.ExcelWriter(main_config["Output_File_Path"], engine="openpyxl", mode="a",
                                if_sheet_exists="replace") as writer:
                product_mix_df.to_excel(writer, sheet_name=main_config["Output_Product_Mix_Comparison_sheetname"],
                                        index=False, startrow=2)
            print("Product mix Comparison sheet Out file is saved")
        except Exception as save_output_file:
            subject = in_config["subject_save_output_file"]
            body = in_config["body_save_output_file"]
            send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
            print("Product mix Comparison Wise Process-", str(save_output_file))
            logging.info("Product mix Comparison sheet Out file is not saved")
            return save_output_file

        # Check outfile creation
        if os.path.exists(main_config["Output_File_Path"]):
            print("Product mix Comparison Logged")
            logging.info("Product mix Comparison sheet is created")
        else:
            subject = in_config["OutputNotFound_Subject"]
            body = in_config["OutputNotFound_Body"]
            send_mail(to=to_address, cc=cc_address, subject=subject, body=body)
            logging.warning("Product mix Comparison sheet is not created")
            raise BusinessException("Output file not generated")

        # Load Sheet in openpyxl
        wb = openpyxl.load_workbook(main_config["Output_File_Path"])
        ws = wb[main_config["Output_Product_Mix_Comparison_sheetname"]]

        for cell in ws["C"]:
            cell.number_format = "#,##"
        for cell in ws["B"]:
            cell.number_format = "#,##"

        full_range = "A3:" + get_column_letter(ws.max_column) + str(ws.max_row)
        ws.auto_filter.ref = full_range
        font_style = Font(name="Cambria", size=11, bold=True, color="000000")
        for c in ascii_uppercase:
            ws[c + "3"].font = font_style
        # for c in ascii_uppercase:
        # ws[c + str(ws.max_row)].font = font_style
        fill_pattern = PatternFill(patternType="solid", fgColor="ADD8E6")
        for c in ascii_uppercase:
            ws[c + "3"].fill = fill_pattern
            if c == "C":
                break

        for c in ascii_uppercase:
            column_length = max(len(str(cell.value)) for cell in ws[c])
            ws.column_dimensions[c].width = column_length * 1.25
            if c == 'C':
                break
        ws['B2'] = '=SUBTOTAL(9,B4:B' + str(ws.max_row) + ')'
        ws['C2'] = '=SUBTOTAL(9,C4:C' + str(ws.max_row) + ')'
        # Save File
        wb.save(main_config["Output_File_Path"])
        logging.info("Completed Product mix Comparison code execution")
        return ws

    except PermissionError as file_error:
        subject = in_config["Permission_Error_Subject"]
        body = in_config["Permission_Error_body"]
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Product mix Comparison Process-", str(file_error))
        print("Please close the file")
        logging.exception(file_error)
        return file_error
    except FileNotFoundError as notfound_error:
        subject = in_config["FileNotFound_Subject"]
        body = in_config["FileNotFound_Body"]
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Product mix Comparison Process-", str(notfound_error))
        logging.exception(notfound_error)
        return notfound_error
    except BusinessException as business_error:
        print("Product mix Comparison Process-", str(business_error))
        logging.exception(business_error)
        return business_error
    except ValueError as value_error:
        subject = in_config["Value_Error"]
        body = in_config["Value_Error_body"]
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Product mix Comparison Process-", str(value_error))
        logging.exception(value_error)
        return value_error
    except TypeError as type_error:
        subject = in_config["Type_Error"]
        body = in_config["Type_Error_body"]
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Product mix Comparison Process-", str(type_error))
        logging.exception(type_error)
        return type_error
    except (OSError, ImportError, MemoryError, RuntimeError, Exception) as error:
        subject = in_config["SystemError_Subject"]
        body = in_config["SystemError_Body"]
        body = body.replace("SystemError +", str(error))
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Product mix Comparison Process-", str(error))
        logging.exception(error)
        return error
    except KeyError as key_error:
        subject = in_config["Name_Error"]
        body = in_config["Name_Error_body"]
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Product mix Comparison Process-", str(key_error))
        logging.exception(key_error)
        return key_error
    except NameError as nameError:
        subject = in_config["Key_Error"]
        body = in_config["Key_Error_body"]
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Product mix Comparison Process-", str(nameError))
        logging.exception(nameError)
        return nameError
    except AttributeError as attributeError:
        subject = in_config["Attribute_Error"]
        body = in_config["Attribute_Error_body"]
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Product mix Comparison Process-", str(attributeError))
        logging.exception(attributeError)
        return attributeError


# Read config details and parse to dictionary
config = {}
main_config = {}
present_quarter_pd = pd.DataFrame()

if __name__ == "__main__":
    print(product_mix_comparison(main_config, config, present_quarter_pd))
