from datetime import datetime, timedelta
import os
import json

import pandas as pd 

from sitewatch4 import sitewatchClient
import sys
# Add the carwash directory to the sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from tg_sender.telegram import telegramBot


tg_message=[]

current_folder_path = os.path.dirname(os.path.abspath(__file__))

cookies_path = os.path.join(current_folder_path,"cookies")

xlfile_path = os.path.join(current_folder_path,"sitewash_data.xlsx")

site_watch_latest_json = os.path.join(current_folder_path,"sitewatch_data_latest.json")

differenec_json = os.path.join(current_folder_path,"difference.json")

sites_df = pd.read_excel(xlfile_path)

orginations_lst = ['SPKLUS-001', 'SPKLUS-002', 
                   'SPKLUS-003', 'SPKLUS-004', 
                   'SPKLUS-005', 'SPKLUS-006', 
                   'SPKLUS-007', 'SPKLUS-008', 
                   'SPKLUS-009', 'SPKLUS-012', 
                   'SPKLUS-013', 'SPKLUS-014', 
                   'SPKLUS-015', 'SUDZWL-002']

client_names = ['Belair', 'Evans', 'North Augusta', 
                'Greenwood', 'Grovetown 1', 'Windsor Springs',
                'Furys Ferry (Martinez)', 'Peach Orchard Rd.', 
                'Grovetown 2', 'Cicero', 'Matteson', 'Sparkle Express ', 
                "Fuller's Calumet City "]

def get_week_dates():
    # Get the current date
    today = datetime.today()
    
    # Find the current week's Monday date
    current_week_monday = today - timedelta(days=today.weekday())
    
    # Find the current week's Sunday date
    current_week_sunday = current_week_monday + timedelta(days=6)
    
    # Format the dates in dd/mm/yyyy format
    monday_date_str = current_week_monday.strftime("%Y-%m-%d")
    sunday_date_str = current_week_sunday.strftime("%Y-%m-%d")
    
    return monday_date_str, sunday_date_str


def generate_weekly_report(file_name,monday_date_str, sunday_date_str):
    
    for index,site in sites_df.iterrows():
        site_dict = site.to_dict()
        cookiesfile_name = f"{(site_dict.get('Organization')).strip().replace('-','_')}.pkl"
        # print(cookiesfile_name)
        print(site_dict)

        cookies_file = os.path.join(cookies_path,cookiesfile_name)
    

        client = sitewatchClient(cookies_file=cookies_file)
        employCode = site_dict.get("employee")
        password = site_dict.get("password")
        locationCode = site_dict.get("Organization")
        client_name = site_dict.get("client_name")
        remember = 1

        
        session_chek = client.check_session_auth(timeout=30)

        if not session_chek:
            token = client.login(employeeCode=employCode,password=password,locationCode=locationCode,remember=1)
            print(token)
        
        session_chek = client.check_session_auth(timeout=30)
        if session_chek:
            reportOn=site_dict.get("site")
            id=site_dict.get("id")
            idname=site_dict.get("id_name")
            request_id = client.get_general_sales_report_request_id(reportOn,id,idname,monday_date_str, sunday_date_str)

            if request_id:
                report_data = client.get_report(reportOn,request_id)
                
                with open(f"{locationCode}.json","w") as f:
                    json.dump(report_data,f,indent=4)



if __name__=="__main__":
    import pandas as pd
    # monday_date_str, sunday_date_str = get_week_dates()
    # generate_weekly_report("sitewatch.xlsx",monday_date_str, sunday_date_str)
    
    # with open("SPKLUS-002.json",'r') as f:
    #     data = json.load(f)
        
    df = pd.read_json("SPKLUS-002.json")
    
    df.to_excel("sitewatch.xlsx")