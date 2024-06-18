from washify import washifyClient
import os 
import json

tg_messages=[]
final_data = {}
diff_dictionary = {}
proxy_url="http://relu;country=US:7d35d7-123852-7e371e-8a2bf4-8e8ad8@private.residential.proxyrack.net:10003"

proxy = {
    "http":proxy_url,
    "https":proxy_url
}

proxy=None

current_file_path = os.path.dirname(os.path.abspath(__file__))

username = 'Cameron'
password = 'Password1'
companyName = 'cleangetawayexpress'
userType = 'CWA'
client  = washifyClient()
# login = client.login(username=username,password=password,
#                     companyName=companyName,userType=userType,proxy=proxy)
# print(f"login check : {login}")
is_logged_in = client.check_login(proxy=proxy)

if not is_logged_in:
    login = client.login(username=username,password=password,
                    companyName=companyName,userType=userType,proxy=proxy)
client_locations = client.get_user_locations()
print(f"is logged in :{is_logged_in}")
print(f"user locations : {client_locations}")
if client_locations:
    for location_name,location_id in client_locations.items():
        curent_car_cnt = client.get_car_count_report([location_id,])
        print(location_name,":",curent_car_cnt)
        final_data[location_name]= curent_car_cnt

latest_json = os.path.join(current_file_path,'washfu_latest.json')
diff_json = os.path.join(current_file_path,'washfu_diff.json')
with open(latest_json,"r") as f:
    old_data = json.load(f)

for location,car_cnt in final_data.items():
    old_cnt = old_data.get(location,0)
    new_cnt = car_cnt if car_cnt else 0
    diff = abs(old_cnt-new_cnt)
    diff_dictionary[location] = diff
    tg_messages.append({"location":"East Peoria","new_value":new_cnt,"diff":diff})

with open(diff_json,"w") as f:
    json.dump(diff_dictionary,f,indent=4)

with open(latest_json,"w") as f:
    json.dump(final_data,f,indent=4)

if tg_messages:
    print(tg_messages)

