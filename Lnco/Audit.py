import sys
from AWS_and_SQL_programs.Process import audit_process
from decouple import Config, RepositoryEnv
from Main_Purchase_Register import reading_sheets_names_from_config_main_sheet
import os
import logging
import datetime
from purchase_send_mail_reusable_task import send_mail

# read env file name from config file

ENV_FILE = 'envfiles/local.env'

try:
    env_file = Config(RepositoryEnv(ENV_FILE))
    print("Env file is successfully read")
except Exception as e:
    print("Exception ", e, "occurred during the reading of env file: ", ENV_FILE)
    print("stopping the process execution...")
    sys.exit("Failed to read env file.")

# get default To and CC address from Env file
default_sender_mail_address = env_file('DEFAULT_SENDER_EMAIL')
default_to_mail_address = env_file('DEFAULT_TO_EMAIL')
default_cc_mail_address = env_file('DEFAULT_CC_EMAIL')
attachment = None
subject = "Bot execution is stopped"
config_main = {}

# create log file name with current datetime
log_file_name = env_file('LOG_FILE_NAME_SUFFIX') + '_' + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".log"
print(log_file_name)

# create log file absolute path
log_file_path = os.path.join(os.getcwd(), 'Logs', log_file_name)
print(log_file_path)

# create log file
try:
    with open(log_file_path, "w") as fp:
        fp.close()
except Exception as file_exception:
    log_file_name = "default_log.log"
    print("Error occurred while creating log file with current date time extension")
    print("error is ", file_exception)
    print("Using default log file ", log_file_name, " for logging purpose,")
    body = "Exception occurred while creating log file with current date time extension. assigning the default log file : default_log.log"
    send_mail(default_to_mail_address, default_cc_mail_address, subject,
              body
              )

# read logging level from Env File
try:
    logging_level = int(env_file('Logging_Level'))
    print("logging value is: ", logging_level)
    if logging_level not in [10, 20, 30, 40, 50]:
        logging_level = 20

except Exception as logging_level_exception:
    print("Exception occurred while reading logging level from env file ", str(logging_level_exception))
    logging_level = 20
    print("logging level has been set to ", logging_level, " and continuing with the program execution")

# Set logging configuration - file name, level and format
try:
    logging.basicConfig(filename=log_file_path, level=logging_level,
                        format='%(asctime)s::%(levelname)s::%(message)s')
    print("Logging basic configuration has been set")
except Exception as log_file_config_exception:
    print("Exception ", log_file_config_exception, " occurred while setting up log file configuration")
    print("Setting default configuration for logging with default config file")
    logging.basicConfig(filename='default_log.log', level=logging_level,
                        format='%(asctime)s::%(levelname)s::%(message)s')

logging.info("Program Execution is started")
logging.info("ENV_FILE used for the program: {}".format(ENV_FILE))

# read config file
path = "Input/Config.xlsx"
config_sheet_name = "Main"

try:
    print("Reading config sheet")
    config_main = reading_sheets_names_from_config_main_sheet(path, config_sheet_name)
    config_main['To_Mail_Address'] = default_to_mail_address
    config_main['CC_Mail_Address'] = default_cc_mail_address
    config_main['Sender_Mail_Address'] = default_sender_mail_address
    print("Reading config sheet is complete")
    logging.info("Input file main sheet reading has been complete")
except Exception as config_exception:
    print("Input file not found or couldn't read")
    logging.critical("Exception occurred while reading config file.")

    subject = 'Bot execution is stopped'
    body = '''
Hello,\n
Below exception occurred while reading config file. Stopping the bot Execution. \n
{}
Thanks,\n
LnCo
    '''.format(str(config_exception))
    send_mail(default_to_mail_address, default_cc_mail_address, subject, body)

    print("Mail notification has been sent - Input load error")

    logging.critical("Failed to read Input file. Hence, stopping the bot")

    logging.warning("Mail notification has been sent - Input load error")
    logging.exception(config_exception)

    exception_type, exception_object, exception_traceback = sys.exc_info()
    filename = exception_traceback.tb_frame.f_code.co_filename
    line_number = exception_traceback.tb_lineno

    logging.critical(str(exception_type))
    logging.critical("Exception occurred in file : {} at line number: {}".format(filename, line_number))

    sys.exit(str(config_exception))

try:
    # get values from Env file to start the audit process
    db_host = env_file('DB_HOST')
    db_name = env_file('DB_NAME')

    db_username = env_file('DB_USERNAME')
    db_password = env_file('DB_PASSWORD')

    aws_bucket_name = env_file('AWS_BUCKET_NAME')
    aws_access_key = env_file('AWS_ACCESS_KEY')
    aws_secret_key = env_file('AWS_SECRET_KEY')

    file_name_to_be_saved_as_in_s3 = env_file('OUTPUT_FILE_NAME')
except Exception as env_file_variables_read_exception:
    logging.critical("Exception occurred while reading env variables hence stopping the bot")
    body = '''
Hello,\n
Below Exception occurred while reading Values from env file. Hence stopping the bot Execution. \n
{} \n
Thanks,\n
LnCo
        '''.format(str(env_file_variables_read_exception))
    print("Sending mail notification")
    logging.warning("Sending mail notification")
    send_mail(default_to_mail_address, default_cc_mail_address, subject, body)
    print("Mail notification has been sent - Env variable load error")
    logging.warning("Mail notification has been sent - Env variable load error")
    logging.exception(env_file_variables_read_exception)

    exception_type, exception_object, exception_traceback = sys.exc_info()
    filename = exception_traceback.tb_frame.f_code.co_filename
    line_number = exception_traceback.tb_lineno

    logging.critical(str(exception_type))
    logging.critical("Exception occurred in file : {} at line number: {}".format(filename, line_number))
    sys.exit(str(env_file_variables_read_exception))

try:
    logging.info("Starting Audit process program execution")
    audit_process(db_host,
                  db_username,
                  db_password,
                  db_name,
                  aws_bucket_name, aws_access_key, aws_secret_key,
                  file_name_to_be_saved_as_in_s3, config_main)

except Exception as audit_process_exception:
    logging.exception(audit_process_exception)
    exception_type, exception_object, exception_traceback = sys.exc_info()
    filename = exception_traceback.tb_frame.f_code.co_filename
    line_number = exception_traceback.tb_lineno

    logging.critical(str(exception_type))
    logging.critical("Exception occurred in file : {0} at line number: {1}".format(filename, line_number))
    # send mail notification
    body = \
        '''
Hello, \n
Below exception occurred while processing audit request.Hence, stopping the bot execution\n
{}\n
Regards,
LnCO
    '''.format(str(audit_process_exception))
    send_mail(default_to_mail_address, default_cc_mail_address, subject, body)
    sys.exit("Audit process execution failed because {}".format(audit_process_exception))
