from AWS_and_SQL_programs.ReadSQLTable import read_sql_and_download_files

aws_bucket_name = 'ca-saas-audit-reports'
host = "ca-saas-audit-db.cwhebe3sgd67.ap-south-1.rds.amazonaws.com"
username = "ca_saas_admin"
password = "BRAD123!"
database = "ca-saas-db"

# connect to SQL and download files from s3
read_sql_and_download_files(host, username, password, database, aws_bucket_name)