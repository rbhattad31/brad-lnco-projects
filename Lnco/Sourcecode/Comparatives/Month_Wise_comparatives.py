import pandas as pd
import numpy
import openpyxl
from win32com import client
import pywintypes
from openpyxl.styles import Font, PatternFill, Side, Border
from string import ascii_lowercase


class BusinessException(Exception):
    pass


def send_mail(to, cc, subject, body):
    try:
        outlook = client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.cc = cc
        mail.Subject = subject
        mail.Body = body
        mail.Send()
    except pywintypes.com_error as message_error:
        print("Sendmail error - Please check outlook connection")
        return message_error
    except Exception as error:
        return error


def purchasemonth(main_config, in_config, present_quarter_pd, previous_quarter_pd):
    try:
        read_present_quarter_pd = present_quarter_pd

        read_previous_quarter_pd = previous_quarter_pd

        # Check Exception
        if read_present_quarter_pd.shape[0] == 0 or read_previous_quarter_pd.shape[0] == 0:
            send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=in_config["subject_mail"],
                      body=in_config["Body_mail"])
            raise BusinessException("Sheet is empty")

        PreviousQuarterSheet_col = read_previous_quarter_pd.columns.values.tolist()
        for col in ["Month", "GR Amt.in loc.cur."]:
            if col not in PreviousQuarterSheet_col:
                subject = in_config["ColumnMiss_Subject"]
                body = in_config["ColumnMiss_Body"]
                body = body.replace("ColumnName +", col)

                send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
                raise BusinessException(col + " Column is missing")

        PresentQuarterSheetColumns = read_present_quarter_pd.columns.values.tolist()
        for col in ["Month", "GR Amt.in loc.cur."]:
            if col not in PresentQuarterSheetColumns:
                subject = in_config["ColumnMiss_Subject"]
                body = in_config["ColumnMiss_Body"]
                body = body.replace("ColumnName +", col)

                send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
                raise BusinessException(col + " Column is missing")

        # Filter Rows
        Month_pd = read_present_quarter_pd[read_present_quarter_pd['Month'].notna()]
        Gr_Amt_pd = read_present_quarter_pd[read_present_quarter_pd['GR Amt.in loc.cur.'].notna()]

        if len(Month_pd) == 0:
            send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=in_config["Month_subject"],
                      body=in_config["Month_Body"])
            raise BusinessException("Month Column is empty")

        elif len(Gr_Amt_pd) == 0:
            send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=in_config["Gr Amt_Subject"],
                      body=in_config["Gr Amt_Body"])
            raise BusinessException("GR Amt Column is empty")
        else:
            pass

        Month_pd_2 = read_previous_quarter_pd[read_previous_quarter_pd['Month'].notna()]
        Gr_Amt_pd_2 = read_previous_quarter_pd[read_previous_quarter_pd['GR Amt.in loc.cur.'].notna()]

        if len(Month_pd_2) == 0:
            send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=in_config["Month_subject"],
                      body=in_config["Month_Body"])
            raise BusinessException("Month Column is empty")

        elif len(Gr_Amt_pd_2) == 0:
            send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=in_config["Gr Amt_Subject"],
                      body=in_config["Gr Amt_Body"])
            raise BusinessException("GR Amt Column is empty")
        else:
            pass

        # create Pivot Table Q4
        pivot_index = ["Month"]
        pivot_values = ["GR Amt.in loc.cur."]
        pivot_PresentQuarter = pd.pivot_table(read_present_quarter_pd, index=pivot_index, values=pivot_values, aggfunc=numpy.sum, margins=True,
                                              margins_name='Grand Total')

        # Get Pivot Column Names
        col_name = pivot_PresentQuarter.columns.values.tolist()

        # Rename Column
        pivot_PresentQuarter = pivot_PresentQuarter.rename(columns={col_name[0]: main_config["PresentQuarterColumnName"]})
        pivot_PresentQuarter = pivot_PresentQuarter.reset_index()

        # Sort based on month
        month_dict = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8, 'September': 9,
                      'October': 10, 'November': 11, 'December': 12, 'Grand Total': 13}

        pivot_PresentQuarter = pivot_PresentQuarter.sort_values('Month', key=lambda x: x.apply(lambda y: month_dict[y]))
        pivot_PresentQuarter.reset_index(inplace=True, drop=True)


        # Create Pivot Table Q3
        pivot_index = ["Month"]
        pivot_values = ["GR Amt.in loc.cur."]
        pivot_PreviousQuarter = pd.pivot_table(read_previous_quarter_pd, index=pivot_index, values=pivot_values, aggfunc=numpy.sum, margins=True,
                                               margins_name='Grand Total')

        # Get Pivot Column Names
        col_name = pivot_PreviousQuarter.columns.values.tolist()

        # Rename Column
        pivot_PreviousQuarter = pivot_PreviousQuarter.rename(columns={col_name[0]: main_config["PreviousQuarterColumnName"]})

        # Remove Index
        pivot_PreviousQuarter = pivot_PreviousQuarter.reset_index()

        # Sort based on month
        pivot_PreviousQuarter = pivot_PreviousQuarter.sort_values('Month', key=lambda x: x.apply(lambda a: month_dict[a]))
        pivot_PreviousQuarter.reset_index(inplace=True, drop=True)

        # Merge Pivot Sheets
        pivot_sheet = pd.concat([pivot_PresentQuarter, pivot_PreviousQuarter], axis=1, sort=False)
        # Remove Empty Rows
        pivot_sheet = pivot_sheet.replace(numpy.nan, '', regex=True)

        # Get Pivot Column Names
        col_name = pivot_sheet.columns.values.tolist()

        # Delete row of Q4 and Q3 columns values as zero
        pivot_sheet.drop(pivot_sheet.index[(pivot_sheet[col_name[1]] == 0) & (pivot_sheet[col_name[3]] == 0)])

        pd.options.mode.chained_assignment = None

        # Variance Formula
        variance_list = []
        for index in pivot_sheet.index:
            quarter_4 = (pivot_sheet[col_name[1]][index])
            quarter_3 = (pivot_sheet[col_name[3]][index])

            if quarter_3 == 0:
                variance = 1
            else:
                variance = (quarter_4 - quarter_3) / quarter_3

            variance_list.append(variance)

        # Create Variance Column
        pivot_sheet['Variance'] = variance_list


        # Log Sheet
        with pd.ExcelWriter(main_config["Output_File_Path"], engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            pivot_sheet.to_excel(writer, sheet_name=main_config["Output_Comparatives_Month_sheetname"], index=False, startrow=16)

        # Load Sheet in openpyxl
        wb = openpyxl.load_workbook(main_config["Output_File_Path"])
        ws = wb[main_config["Output_Comparatives_Month_sheetname"]]

        # Format Q4 & Q3
        for col in ['B', 'D']:
            for cell in ws[col]:
                cell.number_format = "#,###,##.##"

        # Format Variance
        for cell in ws['E']:
            cell.number_format = '0%'

        # Format Header
        format_font = Font(name="Calibri", size=11, color="000000", bold=True)
        font_style1 = Font(name='Cambria', size=12, color='002060', bold=False)
        font_style2 = Font(name='Cambria', size=12, color='002060', bold=True, underline='single')
        font_style3 = Font(name='Cambria', size=14, color='002060', bold=True)


        for c in ascii_lowercase:
            ws[c + "17"].font = format_font

        # Format Footer
        m_row = ws.max_row
        for c in ascii_lowercase:
            ws[c + str(m_row)].font = format_font

        # Header Fill
        format_fill = PatternFill(patternType='solid', fgColor='ADD8E6')
        for c in ascii_lowercase:
            ws[c + "17"].fill = format_fill
            if c == 'e':
                break

        # Set Width
        for c in ascii_lowercase:
            ws.column_dimensions[c].width = 20

        thin = Side(border_style="thin", color='b1c5e7')

        for row in ws.iter_rows(min_row=17, min_col=1, max_row=ws.max_row, max_col=5):
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

        # Cell merge for headers implementation
        ws.merge_cells('A1:E1')
        ws.merge_cells('A2:E2')
        ws.merge_cells('A3:E3')
        ws.merge_cells('A4:E4')
        ws.merge_cells('A5:E5')
        ws.merge_cells('A6:E6')
        ws.merge_cells('A7:E7')
        ws.merge_cells('A8:E8')
        ws.merge_cells('A9:E9')
        ws.merge_cells('A10:E10')
        ws.merge_cells('A11:E11')
        ws.merge_cells('A12:E12')
        ws.merge_cells('A13:E13')
        ws.merge_cells('A14:E14')

        # Headers implementation
        ws['A1'] = main_config['CompanyName']
        ws['A2'] = main_config['StatutoryAuditQuarter']
        ws['A3'] = main_config['FinancialYear']
        ws['A4'] = in_config['A4']
        ws['A5'] = in_config['A5']
        ws['A7'] = in_config['A7']
        ws['A8'] = in_config['A8']
        ws['A10'] = in_config['A10']
        ws['A11'] = in_config['A11']
        ws['A12'] = in_config['A12']

        # Headers formatting and styling
        for row in ws.iter_rows(min_row=1, min_col=1, max_row=5, max_col=1):
            for cell in row:
                cell.font = font_style3

        for row in ws.iter_rows(min_row=7, min_col=1, max_row=7, max_col=1):
            for cell in row:
                cell.font = font_style2

        for row in ws.iter_rows(min_row=10, min_col=1, max_row=10, max_col=1):
            for cell in row:
                cell.font = font_style2

        for row in ws.iter_rows(min_row=8, min_col=1, max_row=8, max_col=1):
            for cell in row:
                cell.font = font_style1

        for row in ws.iter_rows(min_row=11, min_col=1, max_row=12, max_col=1):
            for cell in row:
                cell.font = font_style1

        ws.sheet_view.showGridLines = False
        print(wb.sheetnames)
        # Save File
        wb.save(main_config["Output_File_Path"])
        return ws

    # Excepting Errors here
    except FileNotFoundError as notfound_error:
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=in_config["subject_file_not_found"],
                  body=in_config["body_file_not_found"])
        print("Month Type Wise Comparatives Process-", notfound_error)
        return notfound_error
    except ValueError as V_error:
        subject = in_config["SystemError_Subject"]
        body = in_config["SystemError_Body"]
        body = body.replace("SystemError +", str(V_error))
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Purchase Type Wise Comparatives Process-",V_error )
        return V_error
    except BusinessException as business_error:
        print("Month Type Wise Comparatives Process-", business_error)
        return business_error
    except TypeError as type_error:
        subject = in_config["SystemError_Subject"]
        body = in_config["SystemError_Body"]
        body = body.replace("SystemError +", str(type_error))
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Month Type Wise Comparatives Process-", type_error)
        return type_error
    except (OSError, ImportError, MemoryError, RuntimeError, Exception) as error:
        subject = in_config["SystemError_Subject"]
        body = in_config["SystemError_Body"]
        body = body.replace("SystemError +", str(error))
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Month Type Wise Comparatives Process-",error)
        return error
    except KeyError as key_error:
        subject = in_config["SystemError_Subject"]
        body = in_config["SystemError_Body"]
        body = body.replace("SystemError +", str(key_error))
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Month Type Wise Comparatives Process-", key_error)
        return key_error
    except PermissionError as file_error:
        subject = in_config["SystemError_Subject"]
        body = in_config["SystemError_Body"]
        body = body.replace("SystemError +", str(file_error))
        send_mail(to=main_config["To_Mail_Address"], cc=main_config["CC_Mail_Address"], subject=subject, body=body)
        print("Please close the file")
        return file_error


# Read config details and parse to dictionary
config = {}
main_config = {}
present_quarter_pd = pd.DataFrame()
previous_quarter_pd = pd.DataFrame()
if __name__ == "__main__":
    print(purchasemonth(main_config, config, present_quarter_pd, previous_quarter_pd))

