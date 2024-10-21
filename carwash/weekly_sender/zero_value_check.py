import openpyxl
import logging
logger =logging.getLogger(__file__)
import os
from dotenv import load_dotenv

def check_zero_values(filename, sheet_name):
    # Load the existing workbook using openpyxl
    ret = False
    try:
        workbook = openpyxl.load_workbook(filename)
        worksheet = workbook[sheet_name]
        # Iterate over all rows and columns
        for row in worksheet.iter_rows():
            for cell in row:
                value = cell.value
                # print(value==0)
                if value==0:
                    print(cell.coordinate)
                    ret=True
                    break
                # if value is isinstance(value,(int,float)):
                #     if value==0:
                        # print("Error in report")
                
    except KeyError:
        logger.error(f"Sheet not found: {sheet_name}")
        ret=True
    except FileNotFoundError:
        logger.error(f"File not found: {filename}")
        ret=True
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        ret=True
        
    return ret

if __name__ == "__main__":
    from custom_mailer import send_email,get_excel_files,send_email_on_error
    
    load_dotenv()

    class emailConfig:
        env_vars    = os.environ
        FROM_EMAIL   = env_vars.get("FROM_EMAIL")
        FROM_NAME      = env_vars.get("FROM_NAME")
        SMTP_SERVER   = env_vars.get("SMTP_SERVER")
        SMTP_PORT = env_vars.get("SMTP_PORT")
        SMTP_USER=env_vars.get("SMTP_USER")
        SMTP_PASSWORD=env_vars.get("SMTP_PASSWORD")
        TO_EMAIL=env_vars.get("TO_EMAIL")
    
    # Configuration
    subject = 'Weekly reports'
    body = 'This is the body of the email.'
    to_email = emailConfig.TO_EMAIL
    from_email = emailConfig.FROM_EMAIL
    from_name = emailConfig.FROM_NAME
    smtp_server = emailConfig.SMTP_SERVER
    smtp_port = emailConfig.SMTP_PORT
    smtp_user = emailConfig.SMTP_USER
    smtp_password = emailConfig.SMTP_PASSWORD
    # Set logging level to INFO
    logging.basicConfig(level=logging.INFO)
    
    print()
    zero_val_check = check_zero_values(filename="/home/ubuntu/CAR_WASH_2/carwash_weekly/carwash/weekly_sender/data/2024/2024.xlsx", sheet_name="2024-10-13")
    sunday_date_str ='2024-10-13'
    subject="weekly report"
    #cc_email=[]
    if zero_val_check:
        body = f'Error in report  Ending {sunday_date_str}'
        relu_emails= ["vijaykumarmanthena@reluconsultancy.in"]
        to_email=relu_emails[0]
        cc_emails = relu_emails
        send_email_on_error(subject, body, to_email, from_email, from_name, smtp_server, smtp_port, smtp_user, smtp_password,cc_emails)
        logger.info("error in weekly report ")