from hamilton import hamiltonClient
import json
import os

current_file_path = os.path.dirname(os.path.abspath(__file__))

hamilton_latest_file = os.path.join(current_file_path,"hamilton_latest.json")

proxy_url="http://relu;country=US:7d35d7-123852-7e371e-8a2bf4-8e8ad8@private.residential.proxyrack.net:10003"

proxy = {
    "http":proxy_url,
    "https":proxy_url
}

tg_message=[]

proxy=None

client = hamiltonClient()
login_data = {
    'UserName': 'CR@Sparklecw.com',
    'Password': 'CameronRay1'
    }
# login= client.login(login_data,proxy)
# print(f"login:{login}")
daily_report = client.get_daily_report(proxy)

print("daily report:",daily_report)
if not daily_report:
    print("doing relogin")
    login= client.login(login_data,proxy)
    daily_report = client.get_daily_report(proxy)

if daily_report:
    
    with open(hamilton_latest_file,"r") as f:
        old_data=json.load(f)
        old_diff = old_data.get("East Peoria",0)
        new_value = daily_report.get("East Peoria",0)
        diff = abs(old_diff-new_value)
    print(f"difference : {diff}")
    tg_message.append({"location":"East Peoria","new_value":new_value,"diff":diff})
else:
    tg_message.append(tg_message.append({"location":"East Peoria","msg":"This location is offline message"}))
    
with open(hamilton_latest_file,"w") as f:
    json.dump(daily_report,f)


if tg_message:
    print(tg_message)
