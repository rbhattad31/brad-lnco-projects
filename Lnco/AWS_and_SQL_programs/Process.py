import sys
import boto3
import mysql.connector
import json
from AWS_and_SQL_programs.downloadFilesFromS3Bucket import download_files_from_s3
from AWS_and_SQL_programs.UploadFilesToS3Bucket import upload_file
from Main_Ubuntu import process_execution


def audit_process(host, username, password, database, aws_bucket_name, aws_access_key, aws_secret_key,
                  file_name_to_be_saved_as_in_s3):
    print("DB Host: ", host)
    print("DB :", database)

    print("aws bucket name: ", aws_bucket_name)

    print("Output file name is :", file_name_to_be_saved_as_in_s3)

    # open a connection to the database
    try:
        db_connection = mysql.connector.connect(
            host=host,
            username=username,
            password=password,
            database=database
        )
        print("db connection is established with database: ", database)
        # create a cursor
        db_cursor = db_connection.cursor()
    except Exception as e:
        print("Error Occurred while connecting to db")
        print("Error is: ", e)
        raise Exception("DB Connection establishment failed")
    # read table using cursor
    select_only_new_query = "select * from audit_requests where audit_requests.status='New'"
    db_cursor.execute(select_only_new_query)
    audit_requests_table = db_cursor.fetchall()
    if len(audit_requests_table) == 0:
        print("new audit requests not found")
        sys.exit("new audit requests not found, terminating the program...")
    print("new audit requests are found")
    print("Number of new requests found: ", len(audit_requests_table))
    for row in audit_requests_table:
        request_id = row[0]
        try:
            print("Processing request number: ", request_id)
            db_cursor.execute("UPDATE audit_requests SET status='In progress' where id={}".format(request_id))
            db_connection.commit()
            print("Changed the status of request:", request_id, " to 'In progress'")

            inputs_string = row[1]
            # print(inputs_string)
            inputs_json_object = json.loads(inputs_string)  # converts to dictionary
            # extract values from json
            try:
                bucket_sub_folder_path = inputs_json_object['path']

                mb51_file_path = inputs_json_object['inputs']['MB51']['input_url']
                print(mb51_file_path)
                mb51_sheet_name = inputs_json_object['inputs']['MB51']['sheet_name']
                print(mb51_sheet_name)
                vendor_file_path = inputs_json_object['inputs']['Vendor_Master']['input_url']
                print(vendor_file_path)
                vendor_file_sheet_name = inputs_json_object['inputs']['Vendor_Master']['sheet_name']
                print(vendor_file_sheet_name)
                purchase_register_present_quarter_file_path = \
                    inputs_json_object['inputs']['Purchase_Register_Present_Quarter_File']['input_url']
                print(purchase_register_present_quarter_file_path)
                purchase_register_present_quarter_name = \
                    inputs_json_object['inputs']['Purchase_Register_Present_Quarter_Name']['input_value']
                print(purchase_register_present_quarter_name)
                purchase_register_present_quarter_financial_year = \
                    inputs_json_object['inputs']['Purchase_Register_Present_Quarter_Financial_Year']['input_value']
                print(purchase_register_present_quarter_financial_year)
                purchase_register_present_quarter_sheet_name = \
                    inputs_json_object['inputs']['Purchase_Register_Present_Quarter_File']['sheet_name']
                print(purchase_register_present_quarter_sheet_name)
                purchase_register_previous_quarter_file_path = \
                    inputs_json_object['inputs']['Purchase_Register_Previous_Quarter_File']['input_url']
                print(purchase_register_previous_quarter_file_path)
                purchase_register_previous_quarter_name = \
                    inputs_json_object['inputs']['Purchase_Register_Previous_Quarter_Name']['input_value']
                print(purchase_register_previous_quarter_name)
                purchase_register_previous_quarter_financial_year = \
                    inputs_json_object['inputs']['Purchase_Register_Previous_Quarter_Financial_Year']['input_value']
                print(purchase_register_previous_quarter_financial_year)
                purchase_register_previous_quarter_sheet_name = \
                    inputs_json_object['inputs']['Purchase_Register_Previous_Quarter_File']['sheet_name']
                print(purchase_register_previous_quarter_sheet_name)
                present_quarter_column_name = purchase_register_present_quarter_name + " FY " + purchase_register_present_quarter_financial_year
                print(present_quarter_column_name)
                previous_quarter_column_name = purchase_register_previous_quarter_name + " FY " + purchase_register_previous_quarter_financial_year
                print(previous_quarter_column_name)

            except Exception:
                print("Exception occurred during extracting data from Request")
                print("Request data is not created properly or incomplete...")
                raise Exception("Erroneous request received")

            try:
                print("downloading MB51 file from AWS S3 Bucket")
                mb51_file_saved_path = download_files_from_s3(bucket_name=aws_bucket_name, prefix_name=mb51_file_path,
                                                              aws_access_key_id=aws_access_key,
                                                              aws_secret_access_key=aws_secret_key
                                                              )
                print("mb51 file path: ", mb51_file_saved_path)
                print("MB51 file is downloaded...")
            except Exception as e:
                print(e)
                raise Exception("Error while downloading MB51 file from AWS")

            print("--------------------------------------------------------------------")

            try:
                print("downloading purchase register present quarter file from AWS S3 Bucket")
                purchase_register_present_quarter_saved_path = \
                    download_files_from_s3(bucket_name=aws_bucket_name,
                                           prefix_name=purchase_register_present_quarter_file_path,
                                           aws_access_key_id=aws_access_key,
                                           aws_secret_access_key=aws_secret_key)

                print("Purchase register present quarter file saved path: ",
                      purchase_register_present_quarter_saved_path)
                print("Purchase register present quarter file is downloaded...")

            except Exception:
                raise Exception("Error while downloading purchase register present quarter file")
            print("--------------------------------------------------------------------")

            try:
                print("downloading purchase register previous quarter file from AWS S3 Bucket")
                purchase_register_previous_quarter_saved_path = \
                    download_files_from_s3(bucket_name=aws_bucket_name,
                                           prefix_name=purchase_register_previous_quarter_file_path,
                                           aws_access_key_id=aws_access_key,
                                           aws_secret_access_key=aws_secret_key)

                print("Purchase register previous quarter file saved path: ",
                      purchase_register_previous_quarter_saved_path)
                print("Purchase register previous quarter file is downloaded...")

            except Exception:
                raise Exception("Error while downloading purchase register previous quarter file")
            print("--------------------------------------------------------------------")

            try:
                print("downloading Vendor file from AWS S3 Bucket")
                vendor_file_saved_path = download_files_from_s3(bucket_name=aws_bucket_name,
                                                                prefix_name=vendor_file_path,
                                                                aws_access_key_id=aws_access_key,
                                                                aws_secret_access_key=aws_secret_key
                                                                )
                print("Vendor file path: ", vendor_file_saved_path)
                print("Vendor file is downloaded...")
            except Exception:
                raise Exception("Error while downloading Vendor master file")

            print("--------------------------------------------------------------------")

            input_files = [mb51_file_saved_path, vendor_file_saved_path, purchase_register_present_quarter_saved_path,
                           purchase_register_previous_quarter_saved_path]
            print(input_files[0])
            print(input_files[1])
            print(input_files[2])
            print(input_files[3])

            client_id = row[6]
            print("Company ID in data table: ", client_id)
            # read users data table to get company name
            db_cursor.execute("select `name` from `users` where `id`={}".format(int(client_id)))
            company_row = db_cursor.fetchall()
            print(company_row)
            print(type(company_row))

            company_name = company_row[0]
            print(type(company_name))
            company_name = company_name[0]
            print(company_name)
            statutory_audit_quarter = "Statutory Audit " + purchase_register_present_quarter_name
            print(statutory_audit_quarter)
            financial_year = "FY " + purchase_register_present_quarter_financial_year
            print(financial_year)
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
                                      )
            except Exception as e:
                raise e
            aws_file_path = upload_file(output_file_path, aws_bucket_name, bucket_sub_folder_path,
                                        file_name_to_be_saved_as_in_s3, aws_access_key, aws_secret_key)
            try:
                client = boto3.client(
                    's3',
                    aws_access_key_id=aws_access_key,
                    aws_secret_access_key=aws_secret_key
                )
            except Exception as e:
                raise e
            try:
                obj_list = client.list_objects_v2(
                    Bucket=aws_bucket_name,
                    Prefix=aws_file_path
                )
            except Exception as e:
                raise e

            if 'Contents' in obj_list:
                print("sub folder exist in bucket")
                try:
                    db_cursor.execute("UPDATE audit_requests SET `output_file`=%(output_file)s where `id`=%(id)s",
                                      {'output_file': file_name_to_be_saved_as_in_s3, 'id': request_id})
                    db_connection.commit()
                    print("updated the table with output file path")
                except Exception as e:
                    raise e
                try:
                    db_cursor.execute("UPDATE audit_requests SET `status`='Completed' where `id`={}".format(request_id))
                    db_connection.commit()
                    print("status changed to Completed")
                    print("====================================================================")
                except Exception as e:
                    raise e

        except Exception as e:
            try:

                db_cursor.reset()
                db_cursor.execute("UPDATE audit_requests SET `status`='Failed' where `id`={}".format(request_id))
                db_connection.commit()
                print("Exception occurred during the process of request number: ", request_id)
                print("Exception: ", e)
                print("Processing of request number ", request_id, "has been failed")
                print("Changed the status of request ", request_id, "to 'Failed' in the database")
                print("Updating the error code in datatable")
                db_cursor.reset()
                db_cursor.execute("UPDATE audit_requests SET `error_message`=%(error_message)s where `id`=%(id)s",
                                  {'error_message': str(e), 'id': request_id})
                db_connection.commit()
                print("Updated the error code in datatable")

            except Exception as e:
                print(e)
            continue


if __name__ == '__main__':
    pass
