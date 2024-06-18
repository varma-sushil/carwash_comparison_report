import requests
import pickle
from datetime import date
import os
import json

current_file_path = os.path.dirname(os.path.abspath(__file__))
cookies_path = os.path.join(current_file_path,"cookies")

cookie_file_path = os.path.join(cookies_path,"cookie.json")

data2=os.path.join(current_file_path,"data2.json")

class hamiltonClient():
    def __init__(self) -> None:
        pass
    
    def login(self,login_data:dict,proxy)->bool:
        session  = requests.Session()
        try:
            response = session.post('https://hamiltonservices.com/web/', data=login_data, proxies={'http': proxy})
            if response.status_code==200:
                with open(cookie_file_path, 'wb') as f:
                    pickle.dump(session.cookies, f)
                return True

        except Exception as e:
            print(f"Exception as {e}")
        return False
    
    def get_ccokies(self):
        with open(cookie_file_path, 'rb') as f:
            cookies = pickle.load(f)
        return cookies
    
    def get_daily_report(self,proxy)->int:
        data =None

        cookies = self.get_ccokies()

        headers = {
            'accept': '*/*',
            'accept-language': 'en-US,en;q=0.9',
            'content-type': 'application/json; charset=utf-8',
            # 'cookie': 'ASP.NET_SessionId=h1czucp2l0tedpohyozxp0kn; HamiltonHostedSolutions=9BC65455E711B87BA2E3529488F7E05868084F816AA9395BE06405F19B3D936258D84D23D5D398896E0A8CD32EB682612964CAAA5FB294B72D7B8FACCCF141F39E9BE8E70A72EE30C98D90912F42059B296D6ABEE49C543AE2654761A70D8ECB5C42EC15757808521CCE410970EB5A55E29A0512791475CBFD0B9744CD262595B1D6B4CE0C613AB62ECD69A106E6B1580A7E8354B57477D6AED5A59AC25214F6',
            'dnt': '1',
            'origin': 'https://hamiltonservices.com',
            'priority': 'u=1, i',
            'referer': 'https://hamiltonservices.com/web/Reporting/DailyRevenueTable',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        }

        current_date = date.today()

        formatted_date = current_date.strftime('%Y-%m-%d')

        json_data = {
            'startDate': formatted_date,
            'endDate': formatted_date,
        }

        try:
            response = requests.post(
            'https://hamiltonservices.com/web/Reporting/GetDailyReport',
            cookies=cookies,
            headers=headers,
            json=json_data,proxies=proxy
        )
            print("daily report status code :",response)
            
            if response.status_code==200:
                json_data = response.json()
                data = json_data.get("Data").get("Data")
                items = data.get("Items")
                vi_premier_wash_purchases = 0
                delux_wash_wash_purchases =0
                ultimate_wash_purchases = 0
                dash_wash_purchases =0

                with open(data2,'w') as f:
                    json.dump(json_data,f,indent=4)
                total_wash_purchases = 0
                for item in items :
                    
                    item_id = item.get("ItemId")
                    flag = item.get("Flag")
                    if item_id==1 and not flag:
                        vi_premier_wash_purchases+=1
                    
                    if item_id==2 and not flag:
                        ultimate_wash_purchases+=1
                    if item_id ==3 and not flag:
                        delux_wash_wash_purchases+=1

                    if item_id==4 and not flag:
                        dash_wash_purchases+=1

                total_wash_purchases=sum([vi_premier_wash_purchases,delux_wash_wash_purchases,ultimate_wash_purchases,dash_wash_purchases])

                return {"East Peoria":total_wash_purchases}
        except Exception as e:
            print(f"Error in get_daily_report() {e}")
        
        return data



        

if __name__=="__main__":
    proxy_url=None

    proxy = {
        "http":proxy_url,
        "https":proxy_url
    }
    client = hamiltonClient()
    login_data = {
        'UserName': 'CR@Sparklecw.com',
        'Password': 'CameronRay1'
        }
    # login= client.login(login_data,proxy)
    # print(f"login:{login}")
    daily_report = client.get_daily_report(proxy)
    print(f"daily :{daily_report}")