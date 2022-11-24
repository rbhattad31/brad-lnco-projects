import os
import boto3
from botocore.client import ClientError
import pathlib


def save_file(local_dir_path, file_name, content):
    try:
        # create and store full file path in a variable
        file_path = os.path.join(local_dir_path, file_name)
        file_path = pathlib.Path(file_path)  # to use only backslash for path
        print("File full path is: ", file_path)
        # files saving directory path in a variable
        dir_path = os.path.dirname(file_path)
        print("Directory path is: ", dir_path)
        # If local directory doesn't exist, create it
        if not os.path.exists(dir_path):
            print("File saving directory doesn't exist")
            os.makedirs(dir_path)
            print("Directory is created")
        else:
            print("Files saving directory is exist")

        # write s3 file content to local file
        # wb means write binary
        with open(file_path, 'wb') as f:
            try:
                for chunk in content.iter_chunks(chunk_size=4096):
                    f.write(chunk)
                print("File saved successfully")
                return file_path
            except Exception as e:
                print("Exception occurred while saving the file: ", e)
                return "Failed"
    except Exception as e:
        print("Exception occurred while saving the file: ", e)
        return "Failed"


def download_files_from_s3(bucket_name, prefix_name, aws_access_key_id, aws_secret_access_key):
    print("bucket name: ", bucket_name)
    print("prefix: ", prefix_name)
    project_home_directory = os.getcwd()
    download_file_path = os.path.join(project_home_directory, 'Data', 'Input')
    print("file download path: ", download_file_path)
    # Get resource to check bucket exists
    try:
        resource = boto3.resource(
            's3',
            aws_access_key_id=aws_access_key_id,
            aws_secret_access_key=aws_secret_access_key
        )
        print("Get resource logic is executed")
    except Exception as e:
        print("Exception occurred at initializing resource: ", e)
        raise e

    # check bucket exists from resource
    try:
        resource.meta.client.head_bucket(Bucket=bucket_name)
        print("Bucket ", bucket_name, "exists in s3")
    except ClientError as error:
        error_code = int(error.response['Error']['Code'])
        print("Error occurred, error code is: ", error_code)
        if error_code == 403:
            print("Private Bucket. Forbidden Access! to bucket: ", bucket_name)
            raise error
        elif error_code == 404:
            print("Bucket", bucket_name, " Does Not Exist! ")
            raise error
    except Exception as e:
        print("Exception occurred: ", e)

    # Get s3 client to download files
    try:
        client = boto3.client(
            's3',
            aws_access_key_id=aws_access_key_id,
            aws_secret_access_key=aws_secret_access_key
        )
    except Exception as e:
        print("Exception is occurred while initiating s3 client")
        print("Exception is: ", e)
        raise e

    # check bucket sub folder exist and download files
    # get a list of all the objects that exist in bucket including folders and files
    try:
        obj_list = client.list_objects_v2(
            Bucket=bucket_name,
            Prefix=prefix_name
        )
    except Exception as e:
        print("Exception occurred while getting object lists from bucket ")
        print("Exception is :", e)
        raise e
    try:
        # if Contents in Objects list, sub folder is valid, else invalid
        if 'Contents' in obj_list:
            print("sub folder exist in bucket")
            for obj in obj_list['Contents']:
                try:
                    response = client.get_object(
                        Bucket=bucket_name,
                        Key=obj['Key']
                    )
                    print("Response from s3 bucket is received")
                except Exception as e:
                    print("Exception while receiving response from s3 bucket")
                    raise e
                # check if object is file and download it
                if 'application/x-directory' not in response['ContentType']:
                    print("file key in s3 bucket is ", obj['Key'])
                    for i in range(0, 3):
                        try:
                            save_file_return = save_file(local_dir_path=download_file_path, file_name=obj['Key'],
                                                         content=response['Body'])
                            print("downloaded file from s3 bucket is successful")
                            return save_file_return
                        except Exception as e:
                            print("Exception occurred while downloading file from s3 bucket""")
                            if i == 2:
                                raise e
                            else:
                                print("Retrying to save the file.")
        else:
            raise Exception("Sub Folder doesn't exist in s3 bucket")
    except Exception as e:
        print(e)
        raise e


AWS_ACCESS_KEY_ID_Temp = ''
AWS_SECRET_ACCESS_KEY_Temp = ''
aws_bucket_name_Temp = ''
bucket_subFolder_Temp = ''

if __name__ == '__main__':
    file_saving_path = download_files_from_s3(bucket_name=aws_bucket_name_Temp, prefix_name=bucket_subFolder_Temp,
                                              aws_access_key_id=AWS_ACCESS_KEY_ID_Temp,
                                              aws_secret_access_key=AWS_SECRET_ACCESS_KEY_Temp)
