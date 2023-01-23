import boto3

import json
import logging

from ReusableTasks.downloadFilesFromS3Bucket import download_files_from_s3
from ReusableTasks.UploadFilesToS3Bucket import upload_file
from PurchaseRegister.purchase_register_main import process_execution

import os
from datetime import datetime


def audit_process(aws_bucket_name, aws_access_key, aws_secret_key,
                  config_main, earliest_request_row, env_file, db_connection, company_name):
    row = earliest_request_row
    client_id = row[6]
    print(client_id)
    success_request_status_keyword = config_main['Success_Request_Status']
    fail_request_status_keyword = config_main['Fail_Request_Status']
    db_cursor = db_connection.cursor()
    request_id = row[0]
    print(request_id)
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
            mb51_file_path = inputs_json_object['inputs']['MB51']['input_url']
            print(mb51_file_path)
            logging.debug("MB 51 file path in aws s3 bucket is {}".format(mb51_file_path))
            mb51_sheet_name = inputs_json_object['inputs']['MB51']['sheet_name']
            print(mb51_sheet_name)
            logging.debug("MB 51 sheet name is {}".format(mb51_sheet_name))
            vendor_file_path = inputs_json_object['inputs']['Vendor_Master']['input_url']
            print(vendor_file_path)
            logging.debug("Vendor file path in aws s3 bucket is {}".format(vendor_file_path))
            vendor_file_sheet_name = inputs_json_object['inputs']['Vendor_Master']['sheet_name']
            print(vendor_file_sheet_name)
            logging.debug("Vendor file sheet name is {}".format(vendor_file_sheet_name))
            purchase_register_present_quarter_file_path = \
                inputs_json_object['inputs']['Purchase_Register_Present_Quarter_File']['input_url']
            print(purchase_register_present_quarter_file_path)
            logging.debug("purchase register present quarter file path in aws s3 bucket is {}".format(
                purchase_register_present_quarter_file_path))
            purchase_register_present_quarter_name = \
                inputs_json_object['inputs']['Purchase_Register_Present_Quarter_Name']['input_value']
            print(purchase_register_present_quarter_name)
            logging.debug("purchase register present quarter name is {}".format(purchase_register_present_quarter_name))
            purchase_register_present_quarter_financial_year = \
                inputs_json_object['inputs']['Purchase_Register_Present_Quarter_Financial_Year']['input_value']
            print(purchase_register_present_quarter_financial_year)
            logging.debug("purchase register present quarter financial year is {}".format(
                purchase_register_present_quarter_financial_year))
            purchase_register_present_quarter_sheet_name = \
                inputs_json_object['inputs']['Purchase_Register_Present_Quarter_File']['sheet_name']
            print(purchase_register_present_quarter_sheet_name)
            logging.debug("purchase register present quarter file sheet name is {}".format(
                purchase_register_present_quarter_sheet_name
            ))
            purchase_register_previous_quarter_file_path = \
                inputs_json_object['inputs']['Purchase_Register_Previous_Quarter_File']['input_url']
            print(purchase_register_previous_quarter_file_path)
            logging.debug("purchase register previous quarter file name is {}".format(
                purchase_register_previous_quarter_file_path
            ))
            purchase_register_previous_quarter_name = \
                inputs_json_object['inputs']['Purchase_Register_Previous_Quarter_Name']['input_value']
            print(purchase_register_previous_quarter_name)
            logging.debug(
                "purchase register previous quarter name is {}".format(purchase_register_previous_quarter_name))
            purchase_register_previous_quarter_financial_year = \
                inputs_json_object['inputs']['Purchase_Register_Previous_Quarter_Financial_Year']['input_value']
            print(purchase_register_previous_quarter_financial_year)
            logging.debug("purchase register previous quarter financial year is {}".format(
                purchase_register_previous_quarter_financial_year))
            purchase_register_previous_quarter_sheet_name = \
                inputs_json_object['inputs']['Purchase_Register_Previous_Quarter_File']['sheet_name']
            print(purchase_register_previous_quarter_sheet_name)
            logging.debug("purchase register previous quarter sheet name is {}".format(
                purchase_register_previous_quarter_sheet_name))
            present_quarter_column_name = purchase_register_present_quarter_name + " FY " + purchase_register_present_quarter_financial_year
            print(present_quarter_column_name)
            logging.debug("present quarter is {}".format(present_quarter_column_name))
            previous_quarter_column_name = purchase_register_previous_quarter_name + " FY " + purchase_register_previous_quarter_financial_year
            print(previous_quarter_column_name)
            logging.debug("previous quarter is {}".format(previous_quarter_column_name))
            statutory_audit_quarter = "Statutory Audit " + purchase_register_present_quarter_name
            print(statutory_audit_quarter)
            logging.debug("statutory audit quarter is {}".format(statutory_audit_quarter))
            financial_year = "FY " + purchase_register_present_quarter_financial_year
            print(financial_year)
            logging.debug("financial year is {}".format(financial_year))
        except Exception as input_files_data_extraction_exception:
            print("Exception occurred during extracting data from Request")
            print("Request data is not created properly or incomplete...")
            print(input_files_data_extraction_exception)
            logging.critical("Exception occurred during extracting data from Request")
            logging.critical("Request data is not created properly or incomplete...")
            raise input_files_data_extraction_exception

        try:
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

        except Exception as mb51_file_download_exception:
            logging.critical("Exception occurred while downloading mb51 file from AWS s3 bucket")
            raise mb51_file_download_exception

        print("--------------------------------------------------------------------")

        try:
            print("downloading purchase register present quarter file from AWS S3 Bucket")
            logging.info("downloading purchase register present quarter file from AWS S3 Bucket")
            purchase_register_present_quarter_saved_path = \
                download_files_from_s3(bucket_name=aws_bucket_name,
                                       prefix_name=purchase_register_present_quarter_file_path,
                                       aws_access_key_id=aws_access_key,
                                       aws_secret_access_key=aws_secret_key,
                                       request_id=request_id)
            print("Purchase register present quarter file is downloaded...")
            logging.info("Purchase register present quarter file is downloaded...")
            print("Purchase register present quarter file saved path: ",
                  purchase_register_present_quarter_saved_path)
            logging.debug("Purchase register present quarter file saved path: {}".format(
                purchase_register_present_quarter_saved_path)
            )
            config_main['purchase_register_present_quarter_saved_path'] = purchase_register_present_quarter_saved_path

        except Exception as purchase_register_present_quarter_save_exception:
            logging.critical(
                "Exception occurred while downloading purchase register present quarter file from AWS s3 bucket")
            raise purchase_register_present_quarter_save_exception

        print("--------------------------------------------------------------------")

        try:
            print("downloading purchase register previous quarter file from AWS S3 Bucket")
            # log info
            purchase_register_previous_quarter_saved_path = \
                download_files_from_s3(bucket_name=aws_bucket_name,
                                       prefix_name=purchase_register_previous_quarter_file_path,
                                       aws_access_key_id=aws_access_key,
                                       aws_secret_access_key=aws_secret_key,
                                       request_id=request_id
                                       )
            print("Purchase register previous quarter file is downloaded...")
            logging.info("Purchase register previous quarter file is downloaded...")
            print("Purchase register previous quarter file saved path: ",
                  purchase_register_previous_quarter_saved_path)
            logging.info("Purchase register previous quarter file saved path: {}".format
                         (purchase_register_previous_quarter_saved_path))
            config_main['purchase_register_previous_quarter_saved_path'] = purchase_register_previous_quarter_saved_path

        except Exception as purchase_register_previous_quarter_save_exception:
            logging.critical(
                "Exception occurred while downloading Purchase register previous quarter file from AWS s3 bucket")
            raise purchase_register_previous_quarter_save_exception

        print("--------------------------------------------------------------------")
        try:
            print("downloading Vendor file from AWS S3 Bucket")
            logging.info("downloading Vendor file from AWS S3 Bucket")
            vendor_file_saved_path = download_files_from_s3(bucket_name=aws_bucket_name,
                                                            prefix_name=vendor_file_path,
                                                            aws_access_key_id=aws_access_key,
                                                            aws_secret_access_key=aws_secret_key,
                                                            request_id=request_id
                                                            )
            print("Vendor file is downloaded...")
            logging.info("Vendor file is downloaded...")
            print("Vendor file path: ", vendor_file_saved_path)
            logging.info("Vendor file path: {}".format(vendor_file_saved_path))
            config_main['vendor_file_saved_path'] = vendor_file_saved_path

        except Exception as vendor_file_save_exception:
            logging.critical("Exception occurred while downloading Vendor file from AWS s3 bucket")
            raise vendor_file_save_exception

        print("--------------------------------------------------------------------")

        # create new input files based on column names provided by client
        # get column names from input configuration table -
        mb51_file_id_in_db = env_file('PR_MB51_FILE_ID_IN_DB')
        query_to_get_mb51_column_names = \
            "SELECT `column_names_json` FROM `input_file_configurations` WHERE `user_id` =" + str(
                client_id) + " AND `file_id` = " + str(mb51_file_id_in_db)
        logging.info(
            "query to get MB51 file column names is: \n\t {}".format(str(query_to_get_mb51_column_names)))
        db_cursor.execute(query_to_get_mb51_column_names)
        mb51_column_names_json = db_cursor.fetchall()
        mb51_column_names_json_object = json.loads(mb51_column_names_json[0][0])
        logging.info("MB51 file column names data in json format : \n\t {}".format(mb51_column_names_json_object))

        vendor_file_id_in_db = env_file('VENDOR_FILE_ID_IN_DB')
        query_to_get_vendor_file_column_names = \
            "SELECT `column_names_json` FROM `input_file_configurations` WHERE `user_id` =" + str(
                client_id) + " AND `file_id` = " + str(vendor_file_id_in_db)
        logging.info(
            "query to get vendor file column names is: \n\t {}".format(str(query_to_get_vendor_file_column_names)))
        db_cursor.execute(query_to_get_vendor_file_column_names)
        vendor_file_column_names_json = db_cursor.fetchall()
        vendor_file_column_names_json_object = json.loads(vendor_file_column_names_json[0][0])
        logging.info("vendor file column data read from database: \n\t {}".format(vendor_file_column_names_json_object))

        purchase_file_id_in_db = env_file('PURCHASE_FILE_ID_IN_DB')
        query_to_get_purchase_column_names = \
            "SELECT `column_names_json` FROM `input_file_configurations` WHERE `user_id` =" + str(
                client_id) + " AND `file_id` = " + str(purchase_file_id_in_db)
        logging.info(
            "query to get purchase register file column names is: \n\t {}".format(str(query_to_get_purchase_column_names)))
        db_cursor.execute(query_to_get_purchase_column_names)
        purchase_column_names_json = db_cursor.fetchall()
        purchase_column_names_json_object = json.loads(purchase_column_names_json[0][0])
        logging.info(
            "purchase register file column data read from database: \n\t {}".format(purchase_column_names_json_object))

        # create files with required columns using default columns in input folder under ID Folder

        # Check and change datatypes of each column in the updated files

        input_files = [mb51_file_saved_path, vendor_file_saved_path, purchase_register_present_quarter_saved_path,
                       purchase_register_previous_quarter_saved_path]
        logging.info("mb51 file path {}".format(input_files[0]))
        logging.info("Vendor file path {}".format(input_files[1]))
        logging.info("purchase register present quarter file path {}".format(input_files[2]))
        logging.info("purchase register previous quarter file path {}".format(input_files[3]))

        json_data_list = [mb51_column_names_json_object, vendor_file_column_names_json_object,
                          purchase_column_names_json_object]

        try:
            output_file_path = \
                process_execution(input_files=input_files,
                                  present_quarter_sheet_name=purchase_register_present_quarter_sheet_name,
                                  previous_quarter_sheet_name=purchase_register_previous_quarter_sheet_name,
                                  vendor_master_sheet_name=vendor_file_sheet_name,
                                  mb51_sheet_name=mb51_sheet_name,
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
