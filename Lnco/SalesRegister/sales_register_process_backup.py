import sys
import boto3
import mysql.connector
import json
import logging

import pandas as pd

from ReusableTasks.downloadFilesFromS3Bucket import download_files_from_s3
from ReusableTasks.UploadFilesToS3Bucket import upload_file
from PurchaseRegister.purchase_register_main import process_execution
from ReusableTasks.send_mail_reusable_task import send_mail
from decouple import Config, RepositoryEnv
from sys import platform
from psutil import process_iter
import os
from datetime import datetime


def audit_process(host, username, password, database, aws_bucket_name, aws_access_key, aws_secret_key,
                  file_name_to_be_saved_as_in_s3, config_main):
    present_working_directory = os.getcwd()
    env_file = os.path.join(present_working_directory, '../ENV/env.env')
    env_file = Config(RepositoryEnv(env_file))
    default_to_mail_address = env_file('DEFAULT_TO_EMAIL')
    default_cc_mail_address = env_file('DEFAULT_CC_EMAIL')
    logging.info("DB Host is {}".format(host))
    logging.info("DB Name is {}".format(database))
    logging.info("aws bucket name is {}".format(aws_bucket_name))
    logging.info("Output file name is {}".format(file_name_to_be_saved_as_in_s3))
    logging.info("To mail address in config before update from sql is {}".format(config_main['To_Mail_Address']))

    # open a connection to the database
    try:
        db_connection = mysql.connector.connect(
            host=host,
            username=username,
            password=password,
            database=database
        )
        print("db connection is established with database: ", database)
        logging.info("db connection is established with database: ".format(database))
        # create a cursor
        db_cursor = db_connection.cursor()
        logging.info("cursor is created to the db connection..")
    except Exception as db_connection_exception:
        print("Error Occurred while connecting to db")
        print("Error is: ", db_connection_exception)
        logging.critical("Error Occurred while connecting to db")
        raise db_connection_exception

    try:
        print("Audit requests statuses: ")
        logging.info("Audit requests statuses: ")
        new_status_keyword = config_main['New_Request_Status']
        print("\tNew status request Value in Env file is :{}".format(new_status_keyword))
        logging.info("New status request Value in Env file is :{}".format(new_status_keyword))
        in_progress_status_keyword = config_main['In_Progress_Request_Status']
        print("\tIn progress request value in Env file is :{}".format(in_progress_status_keyword))
        logging.info("In progress request value in Env file is :{}".format(in_progress_status_keyword))
        success_request_status_keyword = config_main['Success_Request_Status']
        print("\tSuccess request value in Env file is :{}".format(success_request_status_keyword))
        logging.info("Success request value in Env file is :{}".format(success_request_status_keyword))
        fail_request_status_keyword = config_main['Fail_Request_Status']
        print("\tFailed request value in Env file is :{}".format(fail_request_status_keyword))
        logging.info("Failed request value in Env file is :{}".format(fail_request_status_keyword))
    except Exception as config_values_not_found_exception:
        print(config_values_not_found_exception)
        logging.warning(
            "Status value keywords are not found or corrupt in env file, setting up default values as below: " + '\n\t' +
            "New, In progress, Success, Failed")
        new_status_keyword = 'New'
        in_progress_status_keyword = 'In Progress'
        success_request_status_keyword = 'Success'
        fail_request_status_keyword = 'Failed'

    # check if any in progress requests
    select_in_progress_requests_query = "select * from audit_requests where audit_requests.status='" + in_progress_status_keyword + "'"
    try:
        db_cursor.execute(select_in_progress_requests_query)
        logging.debug(
            "Query executed to find in progress requests on datatable :\n\t {}".format(
                select_in_progress_requests_query))
    except Exception as select_in_progress_requests_query_exception:
        print(select_in_progress_requests_query_exception)
        logging.critical("Exception occurred while reading in progress requests details from audit request database: ")
        logging.critical(select_in_progress_requests_query_exception)
        # send mail
        raise select_in_progress_requests_query_exception
    audit_requests_in_progress = db_cursor.fetchall()
    # log debug audit_request_in_progress

    # If in progress audit requests found
    if len(audit_requests_in_progress) > 0:
        print("number of requests that are already in progress are: ", len(audit_requests_in_progress))
        logging.info("number of requests that are already in progress are: {}".format(len(audit_requests_in_progress)))
        if platform != "win32":
            print(platform)
            logging.info("Operating platform is {}".format(platform))
        elif platform == "linux":
            print(platform)
            logging.info("Operating platform is {}".format(platform))
            # check if already in progress requests are really in progress or killed
            # get running processes in a list
            for proc in process_iter():
                name = proc.name()
                print(name)
            # get count
            # if count = 1 then
            # update the audit request database entries except present request id as failed with error message

        to_in_progress_request_found = default_to_mail_address
        print(to_in_progress_request_found)
        cc_in_progress_request_found = default_cc_mail_address
        print(cc_in_progress_request_found)
        subject_in_progress_request_found = config_main['subject_in_progress_request_found']
        print(subject_in_progress_request_found)
        body_in_progress_request_found = config_main['body_in_progress_request_found']
        print(body_in_progress_request_found)
        send_mail(to_in_progress_request_found, cc_in_progress_request_found,
                  subject_in_progress_request_found, body_in_progress_request_found)
        logging.info("Notification mail has been sent")
        print(
            "Already 'in progress' requests are found, sent notification mail to the user and stopping the execution..")
        logging.info(
            "Already 'in progress' requests are found, sent notification mail to the user and stopping the execution..")
        sys.exit("Already In progress requests found in the datatable, aborting the program execution")

    # read audit requests table using cursor
    elif len(audit_requests_in_progress) == 0:
        print("number of requests that are already in progress are: ", len(audit_requests_in_progress))
        logging.info("number of requests that are already in progress are: {}".format(len(audit_requests_in_progress)))
        logging.info("Continuing with the process with new requests")

    select_only_new_query = "select * from audit_requests where audit_requests.status='" + new_status_keyword + "'"
    logging.debug("query to get new requests from datatable" + '\n\t' + select_only_new_query)

    print(select_only_new_query)
    try:
        db_cursor.execute(select_only_new_query)
        logging.info("Executed query to get new requests from the datatable: audit_requests")
    except Exception as select_only_new_query_exception:
        print(select_only_new_query_exception)
        logging.critical(
            "Exception occurred while executing the query to get new requests from the audit requests datatable")
        logging.debug(select_only_new_query_exception)
        raise select_only_new_query_exception

    audit_requests_table = db_cursor.fetchall()

    if len(audit_requests_table) == 0:
        print("new audit requests not found")
        logging.info("new audit requests not found in the audit requests datatable, terminating the program")
        # send mail notification ??
        sys.exit("new audit requests not found, terminating the program...")

    print("new audit requests are found")
    logging.info("new audit requests are found")
    print("Number of new requests found: ", len(audit_requests_table))
    logging.info("Number of new requests found: {}".format(len(audit_requests_table)))
    list_of_request_numbers = []
    for row in audit_requests_table:
        try:
            request_id = row[0]
        except Exception as request_id_exception:
            print(request_id_exception)
            logging.critical("Exception occurred while reading request Id from the audit_request datatable")
            raise request_id_exception

        list_of_request_numbers.append(request_id)
    print("List: ", list_of_request_numbers)

    try:
        request_id = int(min(list_of_request_numbers))
        print("Min value in the list: ", request_id)
    except Exception as earliest_request_id_exception:
        print(earliest_request_id_exception)
        logging.critical("Exception occurred while fetching the earliest request ID from the audit_request datatable")
        raise earliest_request_id_exception

    # select only earliest request from the audit requests datatable
    select_only_early_request_query = "select * from audit_requests where `id`= {}".format(request_id)
    print(select_only_early_request_query)
    logging.debug(select_only_early_request_query)

    try:
        db_cursor.execute(select_only_early_request_query)
    except Exception as select_only_early_request_query_exception:
        print(select_only_early_request_query_exception)
        logging.critical("Exception occurred while reading the earliest request {}".format(request_id))
        raise select_only_early_request_query_exception

    print("executed the query to get earliest request of ID: ", request_id)
    logging.info("executed the query to get earliest request of ID: {}".format(request_id))

    audit_requests_table = db_cursor.fetchall()
    logging.info("Earliest new request is read from audit request datatable")
    logging.debug("Earliest audit requests entry in datatable is " + '\n' + str(audit_requests_table[0]))

    row = audit_requests_table
    try:
        print("Processing request number: ", request_id, " is started")
        logging.info("Processing request number: {} is started".format(request_id))
        set_in_progress_query = "UPDATE audit_requests SET `status`='" + in_progress_status_keyword + "' where `id`='" + str(
            request_id) + "'"
        print(set_in_progress_query)
        logging.debug("query to update the status of the request to In progress is :" + '\n\t' + set_in_progress_query)

        try:
            db_cursor.execute(set_in_progress_query)
            db_connection.commit()
            print("Changed the status of request:", request_id, " to ", in_progress_status_keyword)
            logging.info("Changed the status of request: {} to {}".format(request_id, in_progress_status_keyword))
        except Exception as set_in_progress_query_exception:
            logging.critical("Failed to change the status of request {} in audit request table".format(request_id))
            raise set_in_progress_query_exception

        try:
            project_home_directory = os.getcwd()
            print("Project home directory:\n\t", project_home_directory)

            config_download_file_path_in_input = os.path.join(project_home_directory, 'Data', 'Input', 'audit_requests',
                                                              str(request_id), 'config.xlsx')
            config_download_folder_path_in_input = os.path.join(project_home_directory, 'Data', 'Input',
                                                                'audit_requests', str(request_id))
            if not os.path.exists(config_download_folder_path_in_input):
                print("folder path {0} is not exist".format(config_download_folder_path_in_input))
                logging.warning("folder path {0} is not exist".format(config_download_folder_path_in_input))
                os.makedirs(config_download_folder_path_in_input)
                print("created the directory: {0}".format(config_download_folder_path_in_input))
                logging.warning("created the directory: {0}".format(config_download_folder_path_in_input))
                print("Created folder with request number in Input Folder")

            # ------------------------------------------------------------------------------------
            print("Creating new Config file in input folder of Request from Config Dictionary")
            logging.info("Creating new Config file in input folder of Request from Config Dictionary")
            try:
                new_config_df = pd.DataFrame.from_dict(config_main, orient='index', columns=['Value'])
                new_config_df.index.name = 'Key'
                new_config_df.reset_index(level=0, inplace=True)
                new_config_df.to_excel(config_download_file_path_in_input, index=False)
                print("Created new Config file in input folder of Request from Config Dictionary")
                logging.info("Created new Config file in input folder of Request from Config Dictionary")
            except Exception as config_file_save_exception:
                logging.warning("Exception occurred while saving config file in input folder")
                logging.exception(config_file_save_exception)
                exception_type, exception_object, exception_traceback = sys.exc_info()
                filename = exception_traceback.tb_frame.f_code.co_filename
                line_number = exception_traceback.tb_lineno
                logging.warning(str(exception_type))
                logging.warning("Exception occurred in file : {0} at line number: {1}".format(filename, line_number))
            # ------------------------------------------------------------------------------------

            print("Saved config data to an excel file in input folder of request")

        except Exception as config_file_save_exception:
            logging.info("Exception occurred while saving config file to the input folder of the request")
            raise config_file_save_exception

        try:
            client_id = row[0][6]
            print("Company ID in data table: ", client_id)
            logging.info("Company ID in data table: {}".format(client_id))
        except Exception as client_id_exception:
            logging.critical(
                "Exception occurred while reading client ID from the request details for the request {}".format(
                    request_id))
            raise client_id_exception

        # read client row from user data table
        try:
            db_cursor.execute("select * from `users` where `id`={}".format(int(client_id)))
            logging.info("Executed query to get client data from user datatable")
        except Exception as select_client_row_query_exception:
            print(select_client_row_query_exception)
            logging.critical("Exception occurred while reading client details from users datatable")
            raise select_client_row_query_exception

        # read client name and email from the client row read from user data table
        company_row = db_cursor.fetchall()
        company_name = ''
        try:
            if len(company_row) == 1:
                print("Client data from user datatable in database: ", company_row)
                logging.debug("Client data from user datatable in database: {}".format(company_row))
                company_name = company_row[0][1]
                print("Client name: ", company_name)
                logging.info("Client name: {}".format(company_name))
                company_email = company_row[0][4]
                logging.info("Client email: {}".format(company_email))
                print("Client email in SQL Datatable: ", company_email)
                config_main['To_Mail_Address'] = company_email
                print("To mail address in config after update from sql: ", config_main['To_Mail_Address'])
                logging.debug(
                    "To mail address in config after update from sql: {}".format(config_main['To_Mail_Address']))
            elif len(company_row) == 0:
                print("client rows found in user datatable are zero")
                logging.critical("Client details not found in users datatable")
                raise Exception("Client details not found in users datatable")
            elif len(company_row) > 1:
                print("More than one client row is found in user datatable")
                logging.critical("More than one client row is found in user datatable")
                raise Exception("More than one client row is found in user datatable")
            else:
                pass
        except Exception as client_details_Exception:
            print(client_details_Exception)
            raise client_details_Exception

        # read Json string from request
        try:
            inputs_string = row[0][1]
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
        mb51_file_id_in_db = env_file('MB51_FILE_ID_IN_DB')
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
            "query to get vendor file column names is: \n\t {}".format(str(query_to_get_vendor_file_column_names)))
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
                                  request_id=request_id, json_data_list=json_data_list
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
