import os
import boto3
from botocore.client import ClientError
from pathlib import Path

key_id = 'AKIATOQHLB3SCYWESOZL'
secret_key = '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'

AWS_ACCESS_KEY_ID = key_id
AWS_SECRET_ACCESS_KEY = secret_key
aws_bucket_name = 'ca-saas-audit-reports'
bucket_subFolder = 'local/LMG-2/audit-requests/10/'

data_folder_name = 'Data'
input_file_folder = 'Input2'
file_name_to_be_uploaded = 'Output.xlsx'
project_home_directory = os.getcwd()
# Output file of the
upload_file_path = os.path.join(project_home_directory, data_folder_name, input_file_folder, file_name_to_be_uploaded)
file_name_to_be_saved_as_in_s3 = 'Output.xlsx'


def upload_file(file_path, bucket_name, bucket_sub_folder_path, file_name):
    print(file_path)
    print(bucket_name)
    print(bucket_sub_folder_path)
    print(file_name)
    client = boto3.client(
        's3',
        aws_access_key_id=AWS_ACCESS_KEY_ID,
        aws_secret_access_key=AWS_SECRET_ACCESS_KEY
    )

    file_path_check = Path(file_path)
    if file_path_check.exists():
        try:
            print("Input file is exist in the folder")
            print("Uploading file....")
            for i in range(0, 3):
                try:
                    client.upload_file(
                        file_path,
                        bucket_name,
                        os.path.join(bucket_sub_folder_path, file_name)
                    )
                    print("file uploaded successfully")
                    return os.path.join(bucket_sub_folder_path, file_name)
                except Exception as e:
                    print("Exception occurred while uploading file due to ", e)
                    if i + 1 == 3:
                        print("file upload failed")
                    else:
                        print("file upload failed for try", i + 1, ". Trying again...")
                    continue

        except ClientError as e:
            print('Credential is incorrect', e)
        except Exception as e:
            print(e)
    else:
        print("Input file is not valid or not exist")


if __name__ == '__main__':
    aws_file_path = upload_file(file_path=upload_file_path,
                                bucket_name=aws_bucket_name,
                                bucket_sub_folder_path=bucket_subFolder,
                                file_name=file_name_to_be_saved_as_in_s3)
