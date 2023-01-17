import sys
import boto3
import mysql.connector
import json
import logging

import pandas as pd

from decouple import Config, RepositoryEnv
from sys import platform
from psutil import process_iter
import os
from datetime import datetime

from ReusableTasks.downloadFilesFromS3Bucket import download_files_from_s3
from ReusableTasks.UploadFilesToS3Bucket import upload_file

from PurchaseRegister import purchase_register_process
from SalesRegister import sales_register_process

from ReusableTasks.send_mail_reusable_task import send_mail


def audit_process(host, username, password, database, aws_bucket_name, aws_access_key, aws_secret_key,
                  config_main):
    present_working_directory = os.getcwd()
    env_file = os.path.join(os.path.dirname(present_working_directory), 'ENV', 'env.env')
    env_file = Config(RepositoryEnv(env_file))
    default_to_mail_address = env_file('DEFAULT_TO_EMAIL')
    default_cc_mail_address = env_file('DEFAULT_CC_EMAIL')
    logging.info("DB Host is {}".format(host))
    logging.info("DB Name is {}".format(database))
    logging.info("aws bucket name is {}".format(aws_bucket_name))

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
            pass
            # print(platform)
            # logging.info("Operating platform is {}".format(platform))
            # check if already in progress requests are really in progress or killed
            # get running processes in a list
            # for proc in process_iter():
            #     name = proc.name()
            #     print(name)
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

    earliest_request_row = audit_requests_table[0]

    try:
        print("Processing request number: ", request_id, " is started")
        logging.info("Processing request number: {} is started".format(request_id))
        set_in_progress_query = "UPDATE audit_requests SET `status`='" + in_progress_status_keyword + "' where `id`='" + str(
            request_id) + "'"
        print(set_in_progress_query)
        logging.debug("query to update the status of the request to In progress is :" + '\n\t' + set_in_progress_query)

        # update in progress status of the request
        try:
            db_cursor.execute(set_in_progress_query)
            db_connection.commit()
            print("Changed the status of request:", request_id, " to ", in_progress_status_keyword)
            logging.info("Changed the status of request: {} to {}".format(request_id, in_progress_status_keyword))
        except Exception as set_in_progress_query_exception:
            logging.critical("Failed to change the status of request {} in audit request table".format(request_id))
            raise set_in_progress_query_exception

        # create config file creation folder
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
            client_id = earliest_request_row[6]
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
                raise Exception("Unknown Exception occurred while reading client data from datatable")
        except Exception as client_details_Exception:
            print(client_details_Exception)
            raise client_details_Exception
        print(earliest_request_row)
        inputs_string = earliest_request_row[1]
        print(inputs_string)
        print(type(inputs_string))
        purchase_register_keyword = config_main['purchase_register_keyword_in_json']
        sales_register_keyword = config_main['sales_register_keyword_in_json']
        if purchase_register_keyword in inputs_string:
            print("The processing request is identified as Purchase register request")
            purchase_register_process.audit_process(aws_bucket_name, aws_access_key, aws_secret_key,
                                                    config_main, earliest_request_row, env_file,
                                                    db_connection, company_name
                                                    )
        elif sales_register_keyword in inputs_string:
            print("The processing request is identified as Sales register request")
            sales_register_process.audit_process()
    except Exception as request_process_exception:
        print(str(request_process_exception))


if __name__ == '__main__':
    pass
