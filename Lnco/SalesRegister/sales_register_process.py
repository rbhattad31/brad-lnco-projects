import boto3

import json
import logging

from Lnco.ReusableTasks.downloadFilesFromS3Bucket import download_files_from_s3
from Lnco.ReusableTasks.UploadFilesToS3Bucket import upload_file
from Lnco.SalesRegister.sales_register_main import process_execution

import os
from datetime import datetime

from Lnco.ReusableTasks.send_mail_reusable_task import send_mail_with_attachment


def audit_process(aws_bucket_name, aws_access_key, aws_secret_key,
                  config_main, earliest_request_row, env_file, db_connection, company_name):
    row = earliest_request_row
    client_id = row[6]
    print("Client Id is", client_id)
    success_request_status_keyword = config_main['Success_Request_Status']
    fail_request_status_keyword = config_main['Fail_Request_Status']
    db_cursor = db_connection.cursor()
    request_id = row[0]
    print("Request Id is ", request_id)
    try:
        # read Json string from request
        try:

            inputs_string = row[1]
            inputs_json_object = json.loads(inputs_string)  # converts to dictionary
            logging.info("Json data is read from the request data")
            logging.debug(inputs_json_object)
        except Exception as json_string_exception:
            logging.critical("Exception occurred while reading the JSON data from the request")
            raise json_string_exception

        # extract values from json
        try:
            bucket_sub_folder_path = inputs_json_object['path']
            logging.debug("bucket sub folder path is {}".format(bucket_sub_folder_path))

            mb51_file_path = inputs_json_object['inputs']['Sales_Register_MB51']['input_url']
            print("MB 51 file path in aws s3 bucket is \n\t{}".format(mb51_file_path))
            logging.debug("MB 51 file path in aws s3 bucket is \n\t{}".format(mb51_file_path))

            mb51_sheet_name = inputs_json_object['inputs']['Sales_Register_MB51']['sheet_name']
            print(mb51_sheet_name)
            logging.debug("MB 51 sheet name is {}".format(mb51_sheet_name))

            hsn_codes_file_path = inputs_json_object['inputs']['HSN_Codes']['input_url']
            print(hsn_codes_file_path)
            logging.debug("HSN Codes file path in aws s3 bucket is {}".format(hsn_codes_file_path))

            hsn_codes_file_sheet_name = inputs_json_object['inputs']['HSN_Codes']['sheet_name']
            print(hsn_codes_file_sheet_name)
            logging.debug("HSN Codes file sheet name is {}".format(hsn_codes_file_sheet_name))

            sales_register_present_quarter_file_path = \
                inputs_json_object['inputs']['Sales_Register_Current_Quarter_File']['input_url']
            print(sales_register_present_quarter_file_path)
            logging.debug("Sales register present quarter file path in aws s3 bucket is {}".format(
                sales_register_present_quarter_file_path))

            sales_register_present_quarter_sheet_name = \
                inputs_json_object['inputs']['Sales_Register_Current_Quarter_File']['sheet_name']
            print(sales_register_present_quarter_sheet_name)
            logging.debug("purchase register present quarter file sheet name is {}".format(
                sales_register_present_quarter_sheet_name
            ))

            sales_register_present_quarter_name = \
                inputs_json_object['inputs']['Sales_Register_Current_Quarter_Name']['input_value']
            print(sales_register_present_quarter_name)
            logging.debug("Sales register present quarter name is {}".format(sales_register_present_quarter_name))

            sales_register_present_quarter_financial_year = \
                inputs_json_object['inputs']['Sales_Register_Current_Quarter_Year']['input_value']
            print(sales_register_present_quarter_financial_year)
            logging.debug("Sales register present quarter financial year is {}".format(
                sales_register_present_quarter_financial_year))

            sales_register_previous_quarter_file_path = \
                inputs_json_object['inputs']['Sales_Register_Previous_Quarter_File']['input_url']
            print(sales_register_previous_quarter_file_path)
            logging.debug("Sales register previous quarter file name is {}".format(
                sales_register_previous_quarter_file_path
            ))

            sales_register_previous_quarter_sheet_name = \
                inputs_json_object['inputs']['Sales_Register_Previous_Quarter_File']['sheet_name']
            print(sales_register_previous_quarter_sheet_name)
            logging.debug("Sales register previous quarter sheet name is {}".format(
                sales_register_previous_quarter_sheet_name))

            sales_register_previous_quarter_name = \
                inputs_json_object['inputs']['Sales_Register_Previous_Quarter_Name']['input_value']
            print(sales_register_previous_quarter_name)
            logging.debug(
                "Sales register previous quarter name is {}".format(sales_register_previous_quarter_name))

            sales_register_previous_quarter_financial_year = \
                inputs_json_object['inputs']['Sales_Register_Previous_Quarter_Year']['input_value']
            print(sales_register_previous_quarter_financial_year)
            logging.debug("Sales register previous quarter financial year is {}".format(
                sales_register_previous_quarter_financial_year))

            sales_ledger_file_path = inputs_json_object['inputs']['Sales_Ledger']['input_url']
            print(sales_ledger_file_path)
            logging.debug("Sales register previous quarter file name is {}".format(sales_ledger_file_path))

            sales_ledger_sheet_name = inputs_json_object['inputs']['Sales_Ledger']['sheet_name']
            print(sales_ledger_sheet_name)
            logging.debug("Sales register previous quarter sheet name is {}".format(sales_ledger_sheet_name))

            open_po_file_path = inputs_json_object['inputs']['Open_PO']['input_url']
            print(open_po_file_path)
            logging.debug("Sales register previous quarter file name is {}".format(open_po_file_path))

            open_po_sheet_name = inputs_json_object['inputs']['Open_PO']['sheet_name']
            print(open_po_sheet_name)
            logging.debug("Sales register previous quarter sheet name is {}".format(open_po_sheet_name))

            present_quarter_column_name = sales_register_present_quarter_name + " FY " + sales_register_present_quarter_financial_year
            print(present_quarter_column_name)
            logging.debug("present quarter is {}".format(present_quarter_column_name))

            previous_quarter_column_name = sales_register_previous_quarter_name + " FY " + sales_register_previous_quarter_financial_year
            print(previous_quarter_column_name)
            logging.debug("previous quarter is {}".format(previous_quarter_column_name))

            statutory_audit_quarter = "Statutory Audit " + sales_register_present_quarter_name
            print(statutory_audit_quarter)
            logging.debug("statutory audit quarter is {}".format(statutory_audit_quarter))

            financial_year = "FY " + sales_register_present_quarter_financial_year
            print(financial_year)
            logging.debug("financial year is {}".format(financial_year))

        except Exception as input_files_data_extraction_exception:
            print("Exception occurred during extracting data from Request")
            print("Request data is not created properly or incomplete...")
            print(input_files_data_extraction_exception)
            logging.critical("Exception occurred during extracting data from Request")
            logging.exception(input_files_data_extraction_exception)
            logging.critical("Request data is not created properly or incomplete...")
            raise input_files_data_extraction_exception

        try:
            if env_file('SR_VS_MB51') == 'YES':
                print("downloading MB51 file from AWS S3 Bucket")
                logging.info("downloading MB51 file from AWS S3 Bucket")
                mb51_file_saved_path = download_files_from_s3(bucket_name=aws_bucket_name, prefix_name=mb51_file_path,
                                                              aws_access_key_id=aws_access_key,
                                                              aws_secret_access_key=aws_secret_key,
                                                              request_id=request_id
                                                              )
                print("MB51 file is downloaded...")
                logging.info("MB51 file is downloaded...")
                print("mb51 file path: ", mb51_file_saved_path)
                logging.info("mb51 file path is {}".format(mb51_file_saved_path))
                config_main['mb51_file_saved_path'] = mb51_file_saved_path
                # update the path in config and save it in output id folder
            elif env_file('SR_VS_MB51') == 'NO':
                print("Skipped Downloading MB51 File as 'Sales Register Vs MB51' program execution is disabled")
                logging.warning("Skipped Downloading MB51 File as 'Sales Register Vs MB51' program execution is disabled")
                mb51_file_saved_path = None
                logging.info("mb51 file path is {}".format(mb51_file_saved_path))
                config_main['mb51_file_saved_path'] = mb51_file_saved_path
            else:
                mb51_file_saved_path = None
                logging.info("mb51 file path is {}".format(mb51_file_saved_path))
                config_main['mb51_file_saved_path'] = mb51_file_saved_path

        except Exception as mb51_file_download_exception:
            logging.critical("Exception occurred while downloading mb51 file from AWS s3 bucket")
            raise mb51_file_download_exception

        print("--------------------------------------------------------------------")

        try:
            print("downloading sales register present quarter file from AWS S3 Bucket")
            logging.info("downloading sales register present quarter file from AWS S3 Bucket")
            sales_register_present_quarter_saved_path = \
                download_files_from_s3(bucket_name=aws_bucket_name,
                                       prefix_name=sales_register_present_quarter_file_path,
                                       aws_access_key_id=aws_access_key,
                                       aws_secret_access_key=aws_secret_key,
                                       request_id=request_id)
            print("Sales register present quarter file is downloaded...")
            logging.info("Sales register present quarter file is downloaded...")
            print("Sales register present quarter file saved path: ",
                  sales_register_present_quarter_saved_path)
            logging.debug("Sales register present quarter file saved path: {}".format(
                sales_register_present_quarter_saved_path)
            )
            config_main['sales_register_present_quarter_saved_path'] = sales_register_present_quarter_saved_path

        except Exception as sales_register_present_quarter_save_exception:
            logging.critical(
                "Exception occurred while downloading sales register present quarter file from AWS s3 bucket")
            raise sales_register_present_quarter_save_exception

        print("--------------------------------------------------------------------")

        try:
            print("downloading sales register previous quarter file from AWS S3 Bucket")
            # log info
            sales_register_previous_quarter_saved_path = \
                download_files_from_s3(bucket_name=aws_bucket_name,
                                       prefix_name=sales_register_previous_quarter_file_path,
                                       aws_access_key_id=aws_access_key,
                                       aws_secret_access_key=aws_secret_key,
                                       request_id=request_id
                                       )
            print("Sales register previous quarter file is downloaded...")
            logging.info("Sales register previous quarter file is downloaded...")
            print("Sales register previous quarter file saved path: ",
                  sales_register_previous_quarter_saved_path)
            logging.info("Sales register previous quarter file saved path: {}".format
                         (sales_register_previous_quarter_saved_path))
            config_main['Sales_register_previous_quarter_saved_path'] = sales_register_previous_quarter_saved_path

        except Exception as sales_register_previous_quarter_save_exception:
            logging.critical(
                "Exception occurred while downloading sales register previous quarter file from AWS s3 bucket")
            raise sales_register_previous_quarter_save_exception

        print("--------------------------------------------------------------------")
        try:
            if env_file('GST_RATE_CHECK') == 'YES':
                print("downloading HSN Codes file from AWS S3 Bucket")
                logging.info("downloading HSN Codes file from AWS S3 Bucket")
                hsn_codes_file_saved_path = download_files_from_s3(bucket_name=aws_bucket_name,
                                                                   prefix_name=hsn_codes_file_path,
                                                                   aws_access_key_id=aws_access_key,
                                                                   aws_secret_access_key=aws_secret_key,
                                                                   request_id=request_id
                                                                   )
                print("HSN Codes file is downloaded...")
                logging.info("HSN Codes file is downloaded...")
                print("HSN Codes file path: ", hsn_codes_file_saved_path)
                logging.info("HSN Codes file path: {}".format(hsn_codes_file_saved_path))
                config_main['HSN_codes_file_saved_path'] = hsn_codes_file_saved_path
            elif env_file('GST_RATE_CHECK') == 'NO':
                print("Skipped Downloading HSN File as GST RATE CHECK program execution is disabled")
                logging.warning("Skipped Downloading HSN File as GST RATE CHECK program execution is disabled")
                hsn_codes_file_saved_path = None
                config_main['HSN_codes_file_saved_path'] = hsn_codes_file_saved_path
            else:
                hsn_codes_file_saved_path = None
                config_main['HSN_codes_file_saved_path'] = hsn_codes_file_saved_path

        except Exception as hsn_codes_file_save_exception:
            logging.critical("Exception occurred while downloading HSN Codes file from AWS s3 bucket")
            raise hsn_codes_file_save_exception

        print("--------------------------------------------------------------------")
        try:
            if env_file('SR_VS_SL') == 'YES':
                print("downloading Sales Ledger file from AWS S3 Bucket")
                logging.info("downloading Sales Ledger file from AWS S3 Bucket")
                sales_ledger_file_saved_path = download_files_from_s3(bucket_name=aws_bucket_name,
                                                                      prefix_name=sales_ledger_file_path,
                                                                      aws_access_key_id=aws_access_key,
                                                                      aws_secret_access_key=aws_secret_key,
                                                                      request_id=request_id
                                                                      )
                print("Sales Ledger file is downloaded...")
                logging.info("Sales Ledger file is downloaded...")
                print("Sales Ledger file path: ", sales_ledger_file_saved_path)
                logging.info("Sales Ledger file path: {}".format(sales_ledger_file_saved_path))
                config_main['sales_ledger_file_saved_path'] = sales_ledger_file_saved_path
            elif env_file('SR_VS_SL') == 'NO':
                print("Skipped Downloading Sales Ledger File as 'Sales Register Vs Sales Ledger' program execution is disabled")
                logging.warning(
                    "Skipped Downloading Sales Ledger File as 'Sales Register Vs Sales Ledger' program execution is disabled")
                sales_ledger_file_saved_path = None
                logging.info("Sales Ledger file path: {}".format(sales_ledger_file_saved_path))
                config_main['sales_ledger_file_saved_path'] = sales_ledger_file_saved_path
            else:
                sales_ledger_file_saved_path = None
                logging.info("Sales Ledger file path: {}".format(sales_ledger_file_saved_path))
                config_main['sales_ledger_file_saved_path'] = sales_ledger_file_saved_path

        except Exception as sales_ledger_file_save_exception:
            logging.critical("Exception occurred while downloading sales ledger file from AWS s3 bucket")
            raise sales_ledger_file_save_exception

        print("--------------------------------------------------------------------")

        try:
            if env_file('OPEN_PO') == 'YES':
                print("downloading Open PO file from AWS S3 Bucket")
                logging.info("downloading Open PO file from AWS S3 Bucket")
                open_po_file_saved_path = download_files_from_s3(bucket_name=aws_bucket_name,
                                                                 prefix_name=open_po_file_path,
                                                                 aws_access_key_id=aws_access_key,
                                                                 aws_secret_access_key=aws_secret_key,
                                                                 request_id=request_id
                                                                 )
                print("Open PO file is downloaded...")
                logging.info("Open PO file is downloaded...")
                print("Open PO file path: ", open_po_file_saved_path)
                logging.info("Open PO file path: {}".format(open_po_file_saved_path))
                config_main['open_po_file_saved_path'] = open_po_file_saved_path
            elif env_file('OPEN_PO') == 'NO':
                print("Skipped Downloading OPEN PO File as 'Open PO' program execution is disabled")
                logging.warning("Skipped Downloading OPEN PO File as 'Open PO' program execution is disabled")
                open_po_file_saved_path = None
                logging.info("Open PO file path: {}".format(open_po_file_saved_path))
                config_main['open_po_file_saved_path'] = open_po_file_saved_path
            else:
                open_po_file_saved_path = None
                logging.info("Open PO file path: {}".format(open_po_file_saved_path))
                config_main['open_po_file_saved_path'] = open_po_file_saved_path

        except Exception as open_po_file_save_exception:
            logging.critical("Exception occurred while downloading open po file from AWS s3 bucket")
            raise open_po_file_save_exception

        print("--------------------------------------------------------------------")

        # create new input files based on column names provided by client
        # get column names from input configuration table -
        # -------------------------------------------------------------------------------
        mb51_file_id_in_db = env_file('SR_MB51_FILE_ID_IN_DB')
        query_to_get_mb51_column_names = \
            "SELECT `column_names_json` FROM `input_file_configurations` WHERE `user_id` =" + str(
                client_id) + " AND `file_id` = " + str(mb51_file_id_in_db)

        logging.info(
            "query to get MB51 file column names is: \n\t {}".format(str(query_to_get_mb51_column_names)))
        db_cursor.execute(query_to_get_mb51_column_names)
        logging.info("query to get MB51 file column names is completed")
        mb51_column_names_json = db_cursor.fetchall()
        logging.info("")
        mb51_column_names_json_object = json.loads(mb51_column_names_json[0][0])
        logging.info("MB51 file column names data in json format : \n\t {}".format(mb51_column_names_json_object))
        # -------------------------------------------------------------------------------
        hsn_codes_file_id_in_db = env_file('HSN_CODES_FILE_ID_IN_DB')
        query_to_get_hsn_codes_file_column_names = \
            "SELECT `column_names_json` FROM `input_file_configurations` WHERE `user_id` =" + str(
                client_id) + " AND `file_id` = " + str(hsn_codes_file_id_in_db)
        logging.info(
            "query to get hsn codes file column names is: \n\t {}".format(str(query_to_get_hsn_codes_file_column_names)))
        db_cursor.execute(query_to_get_hsn_codes_file_column_names)
        hsn_codes_file_column_names_json = db_cursor.fetchall()
        hsn_codes_file_column_names_json_object = json.loads(hsn_codes_file_column_names_json[0][0])
        logging.info("HSN Codes file column data read from database: \n\t {}".format(hsn_codes_file_column_names_json_object))
        # -------------------------------------------------------------------------------
        sales_register_file_id_in_db = env_file('SALES_REGISTER_FILE_ID_IN_DB')
        query_to_get_sales_register_column_names = \
            "SELECT `column_names_json` FROM `input_file_configurations` WHERE `user_id` =" + str(
                client_id) + " AND `file_id` = " + str(sales_register_file_id_in_db)
        logging.info(
            "query to get sales register file column names is: \n\t {}".format(str(query_to_get_sales_register_column_names)))
        db_cursor.execute(query_to_get_sales_register_column_names)
        sales_register_column_names_json = db_cursor.fetchall()
        sales_register_column_names_json_object = json.loads(sales_register_column_names_json[0][0])
        logging.info(
            "Sales register file column data read from database: \n\t {}".format(sales_register_column_names_json_object))
        # -------------------------------------------------------------------------------
        sales_ledger_file_id_in_db = env_file('SALES_LEDGER_FILE_ID_IN_DB')
        query_to_get_sales_ledger_column_names = \
            "SELECT `column_names_json` FROM `input_file_configurations` WHERE `user_id` =" + str(
                client_id) + " AND `file_id` = " + str(sales_ledger_file_id_in_db)
        logging.info(
            "query to get sales ledger file column names is: \n\t {}".format(
                str(query_to_get_sales_ledger_column_names)))
        db_cursor.execute(query_to_get_sales_ledger_column_names)
        sales_ledger_column_names_json = db_cursor.fetchall()
        sales_ledger_column_names_json_object = json.loads(sales_ledger_column_names_json[0][0])
        logging.info(
            "Sales register file column data read from database: \n\t {}".format(sales_ledger_column_names_json_object))
        # -------------------------------------------------------------------------------
        open_po_file_id_in_db = env_file('OPEN_PO_FILE_ID_IN_DB')
        query_to_get_open_po_column_names = \
            "SELECT `column_names_json` FROM `input_file_configurations` WHERE `user_id` =" + str(
                client_id) + " AND `file_id` = " + str(open_po_file_id_in_db)
        logging.info(
            "query to get open po file column names is: \n\t {}".format(str(query_to_get_open_po_column_names)))
        db_cursor.execute(query_to_get_open_po_column_names)
        open_po_column_names_json = db_cursor.fetchall()
        open_po_column_names_json_object = json.loads(open_po_column_names_json[0][0])
        logging.info(
            "Open PO file column data read from database: \n\t {}".format(open_po_column_names_json_object))
        # -------------------------------------------------------------------------------
        # create files with required columns using default columns in input folder under ID Folder

        # Check and change datatypes of each column in the updated files

        input_files = [mb51_file_saved_path, hsn_codes_file_saved_path, sales_register_present_quarter_saved_path,
                       sales_register_previous_quarter_saved_path, sales_ledger_file_saved_path, open_po_file_saved_path]
        logging.info("mb51 file path is {}".format(input_files[0]))
        logging.info("HSN Codes file path is {}".format(input_files[1]))
        logging.info("purchase register present quarter file path is {}".format(input_files[2]))
        logging.info("purchase register previous quarter file path is {}".format(input_files[3]))
        logging.info("Sales Ledger file path is {}".format(input_files[4]))
        logging.info("Open PO report file path is {}".format(input_files[0]))

        json_data_list = [mb51_column_names_json_object, hsn_codes_file_column_names_json_object,
                          sales_register_column_names_json_object, sales_ledger_column_names_json_object,
                          open_po_column_names_json_object]

        try:
            output_file_path = \
                process_execution(input_files=input_files,
                                  present_quarter_sheet_name=sales_register_present_quarter_sheet_name,
                                  previous_quarter_sheet_name=sales_register_previous_quarter_sheet_name,
                                  hsn_codes_file_sheet_name=hsn_codes_file_sheet_name,
                                  mb51_sheet_name=mb51_sheet_name,
                                  sales_ledger_sheet_name=sales_ledger_sheet_name,
                                  open_po_sheet_name=open_po_sheet_name,
                                  present_quarter_column_name=present_quarter_column_name,
                                  previous_quarter_column_name=previous_quarter_column_name,
                                  company_name=company_name,
                                  statutory_audit_quarter=statutory_audit_quarter,
                                  financial_year=financial_year,
                                  config_main=config_main,
                                  request_id=request_id, json_data_list=json_data_list,
                                  env_file=env_file
                                  )

            print("Output file path is: ", output_file_path)
            output_file_name = os.path.basename(output_file_path)
            logging.info("Audit process program execution is completed")
            logging.info("Output file path is {}".format(output_file_path))
            config_main['output_file_path'] = output_file_path

        except Exception as output_program_exception:
            logging.critical("Exception occurred while executing audit process program")
            raise output_program_exception

        try:
            logging.info("Uploading output file to the AWS s3 bucket")
            aws_file_path = upload_file(output_file_path, aws_bucket_name, bucket_sub_folder_path,
                                        output_file_name, aws_access_key, aws_secret_key)
            logging.info("Uploading the output file to AWS s3 bucket has been complete")
            config_main['output_file_path_in_aws'] = aws_file_path
        except Exception as output_file_upload_exception:
            logging.critical("Exception occurred while uploading output file to AWS s3 Bucket")
            raise output_file_upload_exception

        try:
            logging.info("Connected to AWS s3 bucket for checking if output file is uploaded properly or not")
            client = boto3.client(
                's3',
                aws_access_key_id=aws_access_key,
                aws_secret_access_key=aws_secret_key
            )
            logging.info("Connected to AWS s3 bucket")
        except Exception as aws_client_connection_exception:
            logging.critical("Exception occurred while connecting to AWS S3 bucket")
            raise aws_client_connection_exception
        try:
            logging.info("Fetching objects in the bucket to find output file content")
            obj_list = client.list_objects_v2(
                Bucket=aws_bucket_name,
                Prefix=aws_file_path
            )
            logging.info("Fetching objects in the bucket to find output file content is complete")
        except Exception as aws_s3_list_objects_exception:
            logging.critical("Exception occurred while Fetching objects in the bucket to find output file content")
            raise aws_s3_list_objects_exception
        if 'Contents' in obj_list:
            logging.info("Output file is uploaded to s3 bucket correctly")
            try:
                logging.info("Updating Output file name in audit requests data table")
                db_cursor.execute("UPDATE audit_requests SET `output_file`=%(output_file)s where `id`=%(id)s",
                                  {'output_file': output_file_name, 'id': request_id})
                db_connection.commit()
                print("updated the table with output file name")
                logging.info("Output file name is updated in audit requests data table")
            except Exception as output_file_update_in_datatable_exception:
                logging.critical("Exception occurred while updating output file name in the audit request datatable")
                raise output_file_update_in_datatable_exception

            try:
                logging.info("Updating audit request status as {}".format(success_request_status_keyword))
                set_success_query = "UPDATE audit_requests SET `status`='" + success_request_status_keyword + "' where `id`='" + str(
                    request_id) + "'"
                db_cursor.execute(set_success_query)
                last_updated_timestamp_query = "UPDATE audit_requests SET `updated_at`='" + datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S") + "' where `id`='" + str(
                    request_id) + "'"
                db_cursor.execute(last_updated_timestamp_query)
                db_connection.commit()
                logging.info("Updated audit request status as {}".format(success_request_status_keyword))
                print("status changed to ", success_request_status_keyword)
                # send mail with output file as attachment , output_file_path
                end_to = config_main['To_Mail_Address']
                end_cc = config_main['CC_Mail_Address']
                end_subject = config_main['Success_Mail_Subject']
                end_body = config_main['Success_Mail_Body'].format("Sales Audit Report")
                send_mail_with_attachment(to=end_to, cc=end_cc, body=end_body, subject=end_subject,
                                          attachment_path=output_file_path)
                print("Process complete mail notification is sent")

                print("Bot successfully finished Processing of the sheets")
                print("====================================================================")
            except Exception as datatable_status_success_update_exception:
                logging.critical("Exception occurred while updating audit request status as {}".format(
                    success_request_status_keyword))
                raise datatable_status_success_update_exception

    except Exception as audit_request_exception:
        try:
            logging.warning("Updating audit request status as {} in datatable".format(fail_request_status_keyword))
            db_cursor.reset()
            set_fail_query = "UPDATE audit_requests SET `status`='" + fail_request_status_keyword + "' where `id`='" + str(
                request_id) + "'"
            db_cursor.execute(set_fail_query)

            last_updated_timestamp_query = "UPDATE audit_requests SET `updated_at`='" + datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S") + "' where `id`='" + str(request_id) + "'"
            db_cursor.execute(last_updated_timestamp_query)

            db_connection.commit()
            logging.warning("Updated audit request status as {} in datatable".format(fail_request_status_keyword))
            logging.info("Updating error message in audit request datatable")
            db_cursor.reset()
            db_cursor.execute("UPDATE audit_requests SET `error_message`=%(error_message)s where `id`=%(id)s",
                              {'error_message': str(audit_request_exception), 'id': request_id})
            db_connection.commit()
            logging.info("Updated error message in audit request datatable")
            raise audit_request_exception

        except Exception as audit_request_update_exception:
            logging.critical("Exception occurred while updating audit request failure status or error message")
            logging.debug(audit_request_update_exception)
            raise audit_request_update_exception


if __name__ == '__main__':
    pass
