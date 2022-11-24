from AWS_and_SQL_programs.Process import audit_process
from decouple import Config, RepositoryEnv
QUALITY_DOTENV_FILE = 'quality.env'
LOCAL_DOTENV_FILE = 'local.env'
LOCAL_NEW_DOTENV_FILE = 'local_new.env'
ENV_FILE = LOCAL_NEW_DOTENV_FILE
config = Config(RepositoryEnv(ENV_FILE))


db_host = config('DB_HOST')
db_username = config('DB_USERNAME')
db_password = config('DB_PASSWORD')
db_name = config('DB_NAME')

aws_bucket_name = config('AWS_BUCKET_NAME')
aws_access_key = config('AWS_ACCESS_KEY')
aws_secret_key = config('AWS_SECRET_KEY')

file_name_to_be_saved_as_in_s3 = config('OUTPUT_FILE_NAME')

# connect to SQL and download files from s3
try:
    audit_process(db_host,
                  db_username,
                  db_password,
                  db_name,
                  aws_bucket_name, aws_access_key, aws_secret_key,
                  file_name_to_be_saved_as_in_s3)
except Exception as e:
    print("Exception occurred in the process: ", e)