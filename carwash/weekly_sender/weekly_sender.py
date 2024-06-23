
import sys
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta

# Add the path to the parent directory of "washify" to sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'washify')))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'sitewash')))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'hamilton')))

# print(sys.path)

from washify_weekly import get_week_dates as washify_week_dates
from washify_weekly import generate_weekly_report  as washify_week_report

from sitewatch_weekly import get_week_dates as sitewatch_week_dates
from sitewatch_weekly import generate_weekly_report as sitewatch_week_report

from hamilton_weekly  import get_week_dates as hamilton_week_dates
from hamilton_weekly  import generate_weekly_report as hamilton_week_report

from custom_mailer import send_email,get_excel_files

current_file_path = os.path.dirname(os.path.abspath(__file__))
# print(current_file_path)

data_path = os.path.join(current_file_path,"data")


# print(data_path)

# Your existing code

def get_week_dates_for_storage():
    # Get the current date
    today = datetime.today()
    
    # Find the current week's Monday date
    current_week_monday = today - timedelta(days=today.weekday())
    
    # Find the current week's Sunday date
    current_week_sunday = current_week_monday + timedelta(days=6)
    
    # Format the dates in dd/mm/yyyy format
    monday_date_str = current_week_monday.strftime("%m/%d/%Y")
    sunday_date_str = current_week_sunday.strftime("%m/%d/%Y")
    
    return f"{monday_date_str}-{sunday_date_str}".replace('/','_')



def create_storage_directory(path):
    created_path=None
    try:
        path = os.path.join(data_path,path)
        os.makedirs(path, exist_ok=True)
        print(f"Directory '{path}' created successfully")
        created_path =path
    except OSError as error:
        print(f"Directory '{path}' cannot be created: {error}")
    
    return created_path



if __name__=="__main__":

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
    
    path = get_week_dates_for_storage()
    storage_path = create_storage_directory(path)
    
    wahsify_week_days = washify_week_dates()
    # washify_file_name = 
    # washify_file_path_full = os.path.join(storage_path,washify_file_name)
    washify_week_report(storage_path,wahsify_week_days[0],wahsify_week_days[1])
    
    # hamilton_week_days = hamilton_week_dates()
    # hamilton_file_name = f"hamilton_{hamilton_week_days[0]}-{hamilton_week_days[-1]}.csv".replace('/','_')
    # hamilton_full_path= os.path.join(storage_path,hamilton_file_name)
    # hamilton_week_report(hamilton_full_path,hamilton_week_days[0],hamilton_week_days[-1])
    
    sitewatch_week_days = sitewatch_week_dates()
    site_watch_file_path = storage_path
    sitewatch_week_report(site_watch_file_path,sitewatch_week_days[0],sitewatch_week_days[-1])
    
    # Directory containing Excel files
    directory_path = storage_path
    attachments = get_excel_files(directory_path)
    
    #Sending email to email address
    send_email(subject, body, to_email, from_email, from_name, smtp_server, smtp_port, smtp_user, smtp_password, attachments)
    


