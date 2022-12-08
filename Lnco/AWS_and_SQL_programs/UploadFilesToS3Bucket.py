import os
import boto3
from botocore.client import ClientError
from pathlib import Path


def upload_file(file_path, bucket_name, bucket_sub_folder_path, file_name, aws_access_key_id, aws_secret_access_key):
    print(file_path)
    print(bucket_name)
    print(bucket_sub_folder_path)
    print(file_name)
    client = boto3.client(
        's3',
        aws_access_key_id=aws_access_key_id,
        aws_secret_access_key=aws_secret_access_key
    )

    file_path_check = Path(file_path)
    if file_path_check.exists():
        try:
            print("Output file is exist in the folder")
            print("Uploading file....")
            for i in range(0, 3):
                try:
                    client.upload_file(
                        file_path,
                        bucket_name,
                        os.path.join(bucket_sub_folder_path, file_name)
                    )
                    print("Output file uploaded successfully")
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
        print("Output file is not valid or not exist")


if __name__ == '__main__':
    pass
