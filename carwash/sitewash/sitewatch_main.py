import os
import json

import pandas as pd 

from sitewatch4 import sitewatchClient


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

final_data ={}

difference_dictionary = {} #This will store the absolute difference between th intervells 


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
        request_id = client.get_general_sales_report_request_id(reportOn=reportOn,id=id,name=idname)

        if request_id:
            report_data = client.get_report(reportOn,request_id)
            try:
                gsviews = report_data.get("gsviews")
                gsviews_0 = gsviews[0]
                section_1 = gsviews_0.get("sections")[1]
                reports_0 = section_1.get("reports")[0]

                saleItemCount = reports_0.get("saleItemCount")
                # print(saleItemCount)
                print(f"{locationCode} :=> saleItemCount : {saleItemCount}")
                final_data[locationCode] = saleItemCount
            except Exception as e:
                try:

                    print(f"Exception as scraping record data {e}")
                    gsViews_0 = report_data.get("gsviews")[0]
                    section_4 = gsViews_0.get("sections")[4]
                    reports_0 = section_4.get("reports")[0]
                    quantity = reports_0.get("displayField2")
                    print(f"{locationCode} :=> saleItemCount : {quantity}")
                    final_data[locationCode] = int(quantity)
                except Exception as e:
                    print(f"Error in secound data extraction logic {e}")
                    with open(f"erro2_{cookiesfile_name.replace('pkl','json')}","w") as f:
                        json.dump(report_data,f,indent=4)

                       
with open(site_watch_latest_json,'r') as f:
    old_data =json.load(f)

for org_name in orginations_lst:
    if (org_name in old_data) and (org_name in final_data):
        absolute_difference = abs(old_data[org_name]-final_data[org_name])
        difference_dictionary[org_name] = absolute_difference

with open(differenec_json,'w') as f:
    json.dump(difference_dictionary,f,indent=4)

#update latest data to the file 
with open(site_watch_latest_json,'w') as f:
    json.dump(final_data,f,indent=4)

# current_file_path = os.path.dirname(os.path.abspath(__file__))
# # print(current_file_path)

#

# cookies_file = os.path.join(cookies_path,"sitewatch_cookies.pkl")

# client = sitewatchClient(cookies_file=cookies_file)
# employCode = "20"
# password = 'Cameron1'
# locationCode = 'SPKLUS-001'
# remember = 1

# token = client.login(employeeCode=employCode,password=password,locationCode=locationCode,remember=1)
# print(token)
# session_chek = client.check_session_auth(timeout=15)
# print(session_chek)
# reportOn="site-4"
# id=2121400001
# idname="Site Financial Detail & Chem Report-2021"
# request_id = client.get_general_sales_report_request_id(reportOn=reportOn,id=id,name=idname)

# report_data = client.get_report(reportOn,request_id)
# # print(report_data)

# try:
#     gsviews = report_data.get("gsviews")
#     gsviews_0 = gsviews[0]
#     section_1 = gsviews_0.get("sections")[1]
#     reports_0 = section_1.get("reports")[0]

#     saleItemCount = reports_0.get("saleItemCount")
#     print(saleItemCount)
# except Exception as e:
#     print(f"Exception as e {e}")