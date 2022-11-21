import sys
import boto3
import mysql.connector
import json
from AWS_and_SQL_programs.downloadFilesFromS3Bucket import download_files_from_s3
from AWS_and_SQL_programs.UploadFilesToS3Bucket import upload_file
from Main_Ubuntu import process_execution
host = ""
username = ""
password = ""
database = ""
file_name_to_be_saved_as_in_s3 = 'Output.xlsx'
key_id = 'AKIATOQHLB3SCYWESOZL'
secret_key = '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'

AWS_ACCESS_KEY_ID = key_id
AWS_SECRET_ACCESS_KEY = secret_key
aws_bucket_name = 'ca-saas-audit-reports'


def read_sql_and_download_files(host, username, password, database, aws_bucket_name):
    # open a connection to the database
    db_connection = mysql.connector.connect(
        host=host,
        username=username,
        password=password,
        database=database
    )
    print("db connection is established with database: ", database)
    # create a cursor
    db_cursor = db_connection.cursor()

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
            print(inputs_string)
            inputs_json_object = json.loads(inputs_string)  # converts to dictionary
            bucket_sub_folder_path = inputs_json_object['path']
            mb51_file_path = inputs_json_object['inputs'][0]['input_url']
            print(mb51_file_path)
            mb51_file_saved_path = download_files_from_s3(bucket_name=aws_bucket_name, prefix_name=mb51_file_path)
            print("mb51 file path: ", mb51_file_saved_path)
            print("MB51 file is downloaded...")
            print("--------------------------------------------------------------------")
            purchase_register_file_path = inputs_json_object['inputs'][1]['input_url']
            print(purchase_register_file_path)
            purchase_register_saved_path = download_files_from_s3(bucket_name=aws_bucket_name, prefix_name=purchase_register_file_path)
            print("Vendor file saved path: ", purchase_register_saved_path)
            print("Vendor file is downloaded...")
            print("--------------------------------------------------------------------")
            vendor_file_path = inputs_json_object['inputs'][2]['input_url']
            print(vendor_file_path)
            vendor_file_saved_path = download_files_from_s3(bucket_name=aws_bucket_name, prefix_name=vendor_file_path)
            print("purchase register file path: ", vendor_file_saved_path)
            print("purchase register file is downloaded...")
            print("--------------------------------------------------------------------")
            input_files = [mb51_file_saved_path, purchase_register_saved_path, vendor_file_saved_path]
            print(input_files[0])
            print(input_files[1])
            print(input_files[2])
            output_file_path = process_execution(input_files=input_files)
            aws_file_path = upload_file(output_file_path, aws_bucket_name, bucket_sub_folder_path, file_name_to_be_saved_as_in_s3)
            client = boto3.client(
                's3',
                aws_access_key_id=AWS_ACCESS_KEY_ID,
                aws_secret_access_key=AWS_SECRET_ACCESS_KEY
            )
            obj_list = client.list_objects_v2(
                Bucket=aws_bucket_name,
                Prefix=aws_file_path
            )
            if 'Contents' in obj_list:
                print("sub folder exist in bucket")
                db_cursor.execute("UPDATE audit_requests SET `output_file`=%(output_file)s where `id`=%(id)s", {'output_file': file_name_to_be_saved_as_in_s3, 'id': request_id})
                db_connection.commit()
                print("updated the table with output file path")
                db_cursor.execute("UPDATE audit_requests SET `status`='Completed' where `id`={}".format(request_id))
                db_connection.commit()
                print("status changed to Completed")

        except Exception as e:
            print("Processing of request number: ", request_id, "has been failed")
            db_cursor.reset()
            db_cursor.execute("UPDATE audit_requests SET `status`='Failed' where `id`={}".format(request_id))
            db_connection.commit()
            print("Changed the status of request to 'Failed'")
            print(e)
            continue


if __name__ == '__main__':
    read_sql_and_download_files(host, username, password, database, aws_bucket_name)
