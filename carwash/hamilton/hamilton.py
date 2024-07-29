import requests
import pickle
from datetime import date
import os
import json
import datetime as dt
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
from hamilton_weekly import generate_past_4_weeks_days,generate_past_4_week_days_full
from dotenv import load_dotenv


import csv
current_file_path = os.path.dirname(os.path.abspath(__file__))
cookies_path = os.path.join(current_file_path, "cookies")

cookie_file_path = os.path.join(cookies_path, "cookie.json")

data2 = os.path.join(current_file_path, "data2.json")

#proxy support

dotenv_path="/home/ubuntu/CAR_WASH_2/carwash_weekly/.env"
load_dotenv()

class Config:
    # get environment variables as a dictionary
    env_vars           = os.environ
    PROXY_USER_NAME    = env_vars.get('PROXY_USER_NAME')
    PROXY_PASSWORD     = env_vars.get('PROXY_PASSWORD')
    PROXY_HOST         = env_vars.get('PROXY_HOST')
    PROXY_PORT         = env_vars.get('PROXY_PORT')
    PROXY_ZONE         = env_vars.get("PROXY_ZONE")
    IS_PROXY           =env_vars.get("IS_PROXY")



# Bright Data proxy credentials
username = Config.PROXY_USER_NAME
password = Config.PROXY_PASSWORD
zone = Config.PROXY_ZONE

# # Proxy configuration
# # proxy_host = 'zproxy.lum-superproxy.io'
proxy_host = Config.PROXY_HOST
proxy_port = Config.PROXY_PORT

# # Proxy URL format for datacenter proxy
proxy_url = f'http://{username}-zone-{zone}:{password}@{proxy_host}:{proxy_port}'

username = Config.PROXY_USER_NAME
password = Config.PROXY_PASSWORD
zone = Config.PROXY_ZONE

# # Proxy configuration
# # proxy_host = 'zproxy.lum-superproxy.io'
proxy_host = Config.PROXY_HOST
proxy_port = Config.PROXY_PORT
IS_PROXY = Config.IS_PROXY
# # Proxy URL format for datacenter proxy
proxies =None
# print(IS_PROXY)
if IS_PROXY:
    proxy_url = f'http://{username}-zone-{zone}:{password}@{proxy_host}:{proxy_port}'
    proxies={"http":proxy_url,"https":proxy_url}

class hamiltonClient:
    def __init__(self) -> None:
        self.proxies = proxies

    def login(self, login_data: dict, proxy) -> bool:
        session = requests.Session()
        try:
            response = session.post(
                "https://hamiltonservices.com/web/",
                data=login_data,
                proxies=self.proxies,
            )
            if response.status_code == 200:
                with open(cookie_file_path, "wb") as f:
                    pickle.dump(session.cookies, f)
                return True

        except Exception as e:
            print(f"Exception as {e}")
        return False

    def get_ccokies(self):
        with open(cookie_file_path, "rb") as f:
            cookies = pickle.load(f)
        return cookies

    def get_revenue(self, start_date, end_date):
        headers = {
            "accept": "application/json, text/javascript, */*; q=0.01",
            "accept-language": "en-GB,en-US;q=0.9,en;q=0.8",
            "content-type": "application/json; charset=UTF-8",
            "dxcss": "https://fonts.googleapis.com/css?family=Roboto:300,300i,400,400i,500,500i,700,700i,/web/Content/favicon.svg?rnd=rnd=1.2.7,/web/Content/bootstrap.min.css,/web/Content/ui.dynatree.css,/web/Content/Site.css?rnd=rnd=1.2.8,/web/Content/dx.light.compact.css,https://use.fontawesome.com/releases/v5.7.2/css/all.css,https://cdn.jsdelivr.net/npm/simplebar@latest/dist/simplebar.css,1_69,1_71,0_857,0_718,1_250,0_722,0_853,1_251,0_757,7_19,0_761,0_728,0_733,1_109,0_694,0_880,0_744,4_124,4_115,4_116,0_748,5_3,0_786,4_125,1_77,1_75,/web/Content/dx.common.css",
            "dxscript": "1_16,1_66,1_17,1_18,1_19,1_20,1_21,1_25,1_52,1_51,1_22,1_14,17_6,17_13,1_28,1_225,1_226,1_26,1_27,1_231,1_228,1_234,17_5,1_44,17_27,10_0,10_1,10_2,10_3,10_4,17_28,1_224,17_29,1_24,1_254,1_265,1_266,1_253,1_259,1_257,1_260,1_261,1_258,1_262,1_255,1_263,1_264,1_252,1_268,1_276,1_278,1_279,1_267,1_271,1_272,1_273,1_256,1_269,1_270,1_274,1_275,1_277,1_280,17_0,17_8,1_29,1_36,1_37,1_42,17_17,17_14,1_227,1_46,17_3,1_230,17_21,1_233,17_24,1_235,17_23,1_64,1_236,26_14,26_15,26_12,26_13,26_17,26_19,17_31,26_16,7_16,7_14,7_15,7_13,17_22,8_26,8_8,8_9,8_7,8_17",
            "origin": "https://hamiltonservices.com",
            "priority": "u=1, i",
            "referer": "https://hamiltonservices.com/web/Reporting/NewDxRevenue",
            "sec-ch-ua": '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "sec-fetch-dest": "empty",
            "sec-fetch-mode": "cors",
            "sec-fetch-site": "same-origin",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
            "x-requested-with": "XMLHttpRequest",
        }
        data = {
            "Refresh": False,
            "IsCustomQuery": True,
            "CustomStart": start_date,
            "CustomEnd": end_date,
        }
        response = requests.post(
            "https://hamiltonservices.com/web/Reporting/NewRevenueFilterUpdate",
            cookies=self.get_ccokies(),
            headers=headers,
            json=data,
            proxies=self.proxies
        )

        data = {
            "DXCallbackName": "RevenueReport",
            "__DXCallbackArgument": "c0:page=",
            "RevenueReport": "{&quot;drillDown&quot;:{},&quot;parameters&quot;:{},&quot;cacheKey&quot;:&quot;&quot;,&quot;currentPageIndex&quot;:0}",
            "ClientTime": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3]
            + "Z",
        }

        headers = {
            "accept": "text/html, */*; q=0.01",
            "accept-language": "en-GB,en-US;q=0.9,en;q=0.8",
            "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
            "dxcss": "https://fonts.googleapis.com/css?family=Roboto:300,300i,400,400i,500,500i,700,700i,/web/Content/favicon.svg?rnd=rnd=1.2.7,/web/Content/bootstrap.min.css,/web/Content/ui.dynatree.css,/web/Content/Site.css?rnd=rnd=1.2.8,/web/Content/dx.light.compact.css,https://use.fontawesome.com/releases/v5.7.2/css/all.css,https://cdn.jsdelivr.net/npm/simplebar@latest/dist/simplebar.css,1_69,1_71,0_857,0_718,1_250,0_722,0_853,1_251,0_757,7_19,0_761,0_728,0_733,1_109,0_694,0_880,0_744,4_124,4_115,4_116,0_748,5_3,0_786,4_125,1_77,1_75,/web/Content/dx.common.css",
            "dxscript": "1_16,1_66,1_17,1_18,1_19,1_20,1_21,1_25,1_52,1_51,1_22,1_14,17_6,17_13,1_28,1_225,1_226,1_26,1_27,1_231,1_228,1_234,17_5,1_44,17_27,10_0,10_1,10_2,10_3,10_4,17_28,1_224,17_29,1_24,1_254,1_265,1_266,1_253,1_259,1_257,1_260,1_261,1_258,1_262,1_255,1_263,1_264,1_252,1_268,1_276,1_278,1_279,1_267,1_271,1_272,1_273,1_256,1_269,1_270,1_274,1_275,1_277,1_280,17_0,17_8,1_29,1_36,1_37,1_42,17_17,17_14,1_227,1_46,17_3,1_230,17_21,1_233,17_24,1_235,17_23,1_64,1_236,26_14,26_15,26_12,26_13,26_17,26_19,17_31,26_16,7_16,7_14,7_15,7_13,17_22,8_26,8_8,8_9,8_7,8_17",
            "origin": "https://hamiltonservices.com",
            "priority": "u=1, i",
            "referer": "https://hamiltonservices.com/web/Reporting/NewDxRevenue",
            "sec-ch-ua": '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "sec-fetch-dest": "empty",
            "sec-fetch-mode": "cors",
            "sec-fetch-site": "same-origin",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
            "x-requested-with": "XMLHttpRequest",
        }
        response = requests.post(
            "https://hamiltonservices.com/web/Reporting/NewRevenuePartial",
            cookies=self.get_ccokies(),
            headers=headers,
            data=data,
            proxies=self.proxies
        )

        _return = self.extract(response.text)


        data = {
            "DXCallbackName": "RevenueReport",
            "__DXCallbackArgument": "c0:page=",
            "RevenueReport": "{&quot;drillDown&quot;:{},&quot;parameters&quot;:{},&quot;cacheKey&quot;:&quot;&quot;,&quot;currentPageIndex&quot;:1}",
            "ClientTime": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3]
            + "Z",
        }

        response = requests.post(
            "https://hamiltonservices.com/web/Reporting/NewRevenuePartial",
            cookies=self.get_ccokies(),
            headers=headers,
            data=data,
            proxies=self.proxies
        )

        _return.update(self.extract2(response.text))

        return _return
    
    def extract(self, response):
        # headers = ("","AppWashClubBilling","AppWashClubSignUp","PrepaidPassBilling","Wash","WashClubChange","WashClubReactivation","WashClubSignUp","Cash","CreditCard","AppWashClub","BundleCode","CodeCoupon","RewashCode")
        response = response.replace(r"\-", "-")
        headers = {
            "AppWashClubBilling": "Gross Sales by Item",
            "AppWashClubSignUp": "Gross Sales by Item",
            "PrepaidPassBilling": "Gross Sales by Item",
            "Wash": "Gross Sales by Item",
            "WashClubChange": "Gross Sales by Item",
            "WashClubReactivation": "Gross Sales by Item",
            "WashClubSignUp": "Gross Sales by Item",
            "Cash": "Income Detail",
            "CreditCard": "Income Detail",
            "AppWashClub": "Redeemed Detail",
            "BundleCode": "Redeemed Detail",
            "CodeCoupon": "Redeemed Detail",
            "RewashCode": "Redeemed Detail",
        }
        soup = BeautifulSoup(response, features="html.parser")
        base_table = soup.find("table")

        _return = {
            "Gross Sales by Item": [],
            "Income Detail": [],
            "Refund Detail": [],
            "Redeemed Detail": [],
        }

        print(len(base_table.find_all(class_="csC3423257")))

        for header in base_table.find_all(class_="csC3423257"):
            if header.text not in headers and header.text != "":
                continue

            amount, ratio = header.parent.find_all(class_="cs412993DB")
            if header.text in headers:
                _return[headers[header.text]].append(
                    {"header": header.text, "amount": amount.text, "ratio": ratio.text}
                )

            elif header.text == "" and not _return["Refund Detail"]:
                _return["Refund Detail"].append(
                    {"header": header.text, "amount": amount.text, "ratio": ratio.text}
                )

            elif header.text == "" and not _return["Redeemed Detail"]:
                _return["Redeemed Detail"].append(
                    {"header": header.text, "amount": amount.text, "ratio": ratio.text}
                )

        return _return

    def extract2(self, response):
        response = response.replace(r"\-", "-")

        soup = BeautifulSoup(response, features="html.parser")
        base_table = soup.find("table")

        _return = {"Tax Detail": []}

        header = base_table.find(class_="csC3423257")

        amount, ratio = header.parent.find_all(class_="cs412993DB")

        _return["Tax Detail"].append(
            {"header": header.text, "amount": amount.text, "ratio": ratio.text}
        )

        return _return
    
    def to_csv(self, data, filename):
        _return = []
        
        for key, value in data.items():
            if key == "Tax Detail":
                _return.append([key, 'Rate', 'Current Amount'])
            else:
                _return.append([key, 'Amount', 'Ratio'])

            x = 0

            for item in value:
                _return.append([item['header'],item['amount'], item['ratio']])
                if key == "Tax Detail":
                    x += float(item['ratio'].replace(",",'').replace("$","").replace("%",""))
                else:
                    x += float(item['amount'].replace(",",'').replace("$","").replace("%",""))
            
            if key == "Tax Detail":
                _return.append(['Total',"", f"${x}"])

            else:
                _return.append(['Total',f"${x}",""])

            _return.append(['','',''])

        
        with open(f'{filename}','w', newline="") as f:
            csv_writer = csv.writer(f)
            csv_writer.writerows(_return)

    def to_excel(self,data, filename):
        _return = []
        
        for key, value in data.items():
            if key == "Tax Detail":
                _return.append([key, 'Rate', 'Current Amount'])
            else:
                _return.append([key, 'Amount', 'Ratio'])

            x = 0

            for item in value:
                _return.append([item['header'], item['amount'], item['ratio']])
                if key == "Tax Detail":
                    x += float(item['ratio'].replace(",", '').replace("$", "").replace("%", ""))
                else:
                    x += float(item['amount'].replace(",", '').replace("$", "").replace("%", ""))
            
            if key == "Tax Detail":
                _return.append(['Total', "", f"${x}"])
            else:
                _return.append(['Total', f"${x}", ""])

            _return.append(['', '', ''])

        # Convert the list of lists to a DataFrame
        df = pd.DataFrame(_return)

        # Write the DataFrame to an Excel file
        df.to_excel(filename, index=False, header=False)

    def get_car_count(self,items):
        "This function will return car count"

        wash_purchases_total_cnt3 = 0
        reedeemd_total_cnt3 = 0
        retail_revenue3=0.0
        total_revenue3 = 0.0
        arm_plans_sold3 = 0
        
        for item in items :
            itemtyp = item.get("ItemType")
            discount = item.get("Discount")
            flag= item.get("Flag")
            price = item.get("Price")
            
            if flag:
                reedeemd_total_cnt3+=1
            
            elif itemtyp=="Wash" and not (flag or discount): #wash purchase
                wash_purchases_total_cnt3+=1
                retail_revenue3+=price
            
            if not (flag or   discount): 
                total_revenue3+=price
                
            if itemtyp in ["WashClubReactivation","WashClubSignUp","AppWashClubSignUp"] :# WashClubSignUp,  # arm plans sold 
                arm_plans_sold3+=1
                
        return sum([wash_purchases_total_cnt3,reedeemd_total_cnt3])

    def get_daily_report(self, proxy) -> int:
        data = None

        cookies = self.get_ccokies()

        headers = {
            "accept": "*/*",
            "accept-language": "en-US,en;q=0.9",
            "content-type": "application/json; charset=utf-8",
            # 'cookie': 'ASP.NET_SessionId=h1czucp2l0tedpohyozxp0kn; HamiltonHostedSolutions=9BC65455E711B87BA2E3529488F7E05868084F816AA9395BE06405F19B3D936258D84D23D5D398896E0A8CD32EB682612964CAAA5FB294B72D7B8FACCCF141F39E9BE8E70A72EE30C98D90912F42059B296D6ABEE49C543AE2654761A70D8ECB5C42EC15757808521CCE410970EB5A55E29A0512791475CBFD0B9744CD262595B1D6B4CE0C613AB62ECD69A106E6B1580A7E8354B57477D6AED5A59AC25214F6',
            "dnt": "1",
            "origin": "https://hamiltonservices.com",
            "priority": "u=1, i",
            "referer": "https://hamiltonservices.com/web/Reporting/DailyRevenueTable",
            "sec-ch-ua": '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "sec-fetch-dest": "empty",
            "sec-fetch-mode": "cors",
            "sec-fetch-site": "same-origin",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        }

        current_date = date.today()

        formatted_date = current_date.strftime("%Y-%m-%d")

        json_data = {
            "startDate": formatted_date,
            "endDate": formatted_date,
        }

        try:
            response = requests.post(
                "https://hamiltonservices.com/web/Reporting/GetDailyReport",
                cookies=cookies,
                headers=headers,
                json=json_data,
                proxies=proxy,
            )
            print("daily report status code :", response)

            if response.status_code == 200:
                json_data = response.json()
                data = json_data.get("Data").get("Data")
                items = data.get("Items")
                vi_premier_wash_purchases = 0
                delux_wash_wash_purchases = 0
                ultimate_wash_purchases = 0
                dash_wash_purchases = 0

                # with open(data2,'w') as f:
                #     json.dump(json_data,f,indent=4)
                total_wash_purchases = 0
                for item in items:

                    item_id = item.get("ItemId")
                    flag = item.get("Flag")
                    if item_id == 1 and not flag:
                        vi_premier_wash_purchases += 1

                    if item_id == 2 and not flag:
                        ultimate_wash_purchases += 1
                    if item_id == 3 and not flag:
                        delux_wash_wash_purchases += 1

                    if item_id == 4 and not flag:
                        dash_wash_purchases += 1

                total_wash_purchases = sum(
                    [
                        vi_premier_wash_purchases,
                        delux_wash_wash_purchases,
                        ultimate_wash_purchases,
                        dash_wash_purchases,
                    ]
                )

                return {"East Peoria": total_wash_purchases}
        except Exception as e:
            print(f"Error in get_daily_report() {e}")

        return data

    def get_dail_report_v2(self,startDate,endDate,proxy=None):
        

        headers = {
            'accept': '*/*',
            'accept-language': 'en-US,en;q=0.9',
            'content-type': 'application/json; charset=utf-8',
            # 'cookie': 'ASP.NET_SessionId=yfzp50ajywiy1p2ycdrvrdbc; HamiltonHostedSolutions=8B0E35FEB571F88C228642C6C49325020CF1BF4DEC130EF15487BD9F2B3FDE2D059596E7B0623ADF26EF23824FF452D9BDB4212DB9852AC88FBD29149BEA544CFADF1FD03A8448F8638836F6170CA2837CDED44D2B405B0E2FBC338299A217AFE59077F1827FCD8C2964D35F91FA945D2AA97C9FADD25975EB4251FA8D247D9D46D3D0166CF7A1F74B43162A247DA79F35B6C6B25A3F73FFFD73C193DF721E40',
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

        json_data = {
            'startDate':startDate ,# "2024-06-29"
            'endDate': endDate,
        }
        cookies = self.get_ccokies()

        try:
            response = requests.post(
            'https://hamiltonservices.com/web/Reporting/GetDailyReport',
            cookies=cookies,
            headers=headers,
            json=json_data,
            proxies=self.proxies
        )
            if response.status_code==200:
                data= response.json().get("Data")
                data2 = data.get("Data")
                items = data2.get("Items")
                
                return items
        except Exception as e:
            print(f"Exception in get_dail_report_v2() {e}")
        
    def get_days_for_Total_membership(self,monday):


        # Get the old date 
        current_date = datetime.strptime(monday, "%Y-%m-%d")#.now()#
        
        # Get the date 30 days before the current date
        date_30_days_ago = current_date - timedelta(days=30)
        
        # Format both dates to the desired format
        current_date_str = current_date.strftime('%Y-%m-%d')
        date_30_days_ago_str = date_30_days_ago.strftime('%Y-%m-%d')
        
        print(f"Current Date: {current_date_str}")
        print(f"Date 30 Days Ago: {date_30_days_ago_str}")
        
        return current_date_str,date_30_days_ago_str


    def get_total_plan_members(self,monday):
        "will return total plan members by considering last 30days"
        total_plan_members = 0
        
        cookies = self.get_ccokies()

        headers = {
            'accept': '*/*',
            'accept-language': 'en-US,en;q=0.9',
            'content-type': 'application/json; charset=utf-8',
            # 'cookie': 'ASP.NET_SessionId=pmqmo14o2uq3smgwvzpvuqvk; HamiltonHostedSolutions=4C37D484F9CA28F54EB1FE2D62D3FF7AB1EF23DC5672B266016BC960225599A3D40CD32A0779F501D33DE06AE9624DEFD925BCF95774F6C1FCEDF0B8D9393846BD08BEC0C4A8A6DC6D36D3E4125C4EA4660313B2D2AF082910B9639C71B3757DD2F7F1D5E230707124504E61C1C862FCAD2BD52CC868057DED2243429F885DB55DE87FEA6A060B852B720F59C35541CEF789DB4337C12FD39F43D05521DE0BC4',
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

        endDate,startDate = self.get_days_for_Total_membership(monday)
        json_data = {
            'startDate': startDate,
            'endDate': endDate,
        }

        try:
            response = requests.post(
                'https://hamiltonservices.com/web/Reporting/GetDailyReport',
                cookies=cookies,
                headers=headers,
                json=json_data,
                proxies=self.proxies
            )
            if response.status_code==200:
                
                data = response.json().get("Data").get("Data")
                items = data.get("Items")
                for item in items:
    
                    itemtyp = item.get("ItemType")
                    
                    if itemtyp in ["AppWashClubBilling","PrepaidPassBilling"]:
                        total_plan_members+=1
            
        except Exception as e:
            print(f"Exception in get_total_plan_members()  in {e}")
            
        return total_plan_members
    

   
    


def conversion_rate_hamilton(arm_plans_sold_cnt,wash_purchases_total_cnt,wash_purchases_total_cnt2):
    rate = 0
    try:
        rate = arm_plans_sold_cnt/ sum([wash_purchases_total_cnt,wash_purchases_total_cnt2])
        rate =rate*100
        rate = round(rate,2)
    except Exception as e:
        print(f"Exception in conversion_rate_hamilton() {e}")
    return rate

def find_retail_revenue_and_total_revenue(items):
    wash_purchases_total_cnt = 0
    reedeemd_total_cnt = 0
    retail_revenue=0.0
    total_revenue = 0.0
    arm_plans_sold = 0
    
    for item in items :
        itemtyp = item.get("ItemType")
        discount = item.get("Discount")
        flag= item.get("Flag")
        price = item.get("Price")
        
        if flag:            #Reedemed
            reedeemd_total_cnt+=1
        
        elif itemtyp=="Wash" and not (flag or discount): #wash purchase
            wash_purchases_total_cnt+=1
            retail_revenue+=price
        
        if not (flag or   discount): #total revenue monday- friday
            total_revenue+=price
            
        if itemtyp in ["WashClubReactivation","WashClubSignUp","AppWashClubSignUp"] :# WashClubSignUp,  # arm plans sold 
            arm_plans_sold+=1

    return {
        'reedeemd_total_cnt':reedeemd_total_cnt,
        'wash_purchases_total_cnt':wash_purchases_total_cnt,
        'retail_revenue':retail_revenue,
        'total_revenue':total_revenue,
        'arm_plans_sold':arm_plans_sold
    }

def generate_report(monday_date_str, friday_date_str, saturday_date_str, sunday_date_str):
    final_data = {}
    proxy_url = None

    proxy = {"http": proxy_url, "https": proxy_url}
    client = hamiltonClient()
    login_data = {"UserName": "CR@Sparklecw.com", "Password": "CameronRay1"}
    login = client.login(login_data, proxy)
    
    items = client.get_dail_report_v2(monday_date_str,friday_date_str)
    
    data_v1 = find_retail_revenue_and_total_revenue(items)
    wash_purchases_total_cnt = data_v1['wash_purchases_total_cnt']
    retail_revenue = data_v1['retail_revenue']
    total_revenue = data_v1['total_revenue']
    reedeemd_total_cnt = data_v1['reedeemd_total_cnt']
    arm_plans_sold1 = data_v1['arm_plans_sold']
    
    # wash_purchases_total_cnt = 0
    # reedeemd_total_cnt = 0
    # retail_revenue=0.0
    # total_revenue = 0.0
    # arm_plans_sold1 = 0
    
    # for item in items :
    #     itemtyp = item.get("ItemType")
    #     discount = item.get("Discount")
    #     flag= item.get("Flag")
    #     price = item.get("Price")
        
    #     if flag:                    #Reedemed
    #         reedeemd_total_cnt+=1
        
    #     elif itemtyp=="Wash" and not (flag or discount): #wash purchase
    #         wash_purchases_total_cnt+=1
    #         retail_revenue+=price
        
    #     if not (flag or   discount): #total revenue monday- friday
    #         total_revenue+=price
            
    #     if itemtyp in ["WashClubReactivation","WashClubSignUp","AppWashClubSignUp"] :# WashClubSignUp,  # arm plans sold 
    #         arm_plans_sold1+=1
        
    
    final_data["car_count_monday_to_friday"] = sum([wash_purchases_total_cnt,reedeemd_total_cnt])
    final_data["arm_plans_reedemed_monday_to_friday_cnt"]  = "" #update
    final_data["retail_car_count_monday_to_friday"] = wash_purchases_total_cnt
    final_data["retail_revenue_monday_to_friday"] = retail_revenue
    final_data["total_revenue_monday_to_friday"] =  total_revenue
    final_data["labour_hours_monday_to_friday"]  = ""
    final_data["cars_per_labour_hour_monday_to_friday"] = ""
    
    #for saturday to sunday
    items = client.get_dail_report_v2(saturday_date_str, sunday_date_str)

    data_v2 = find_retail_revenue_and_total_revenue(items)
    wash_purchases_total_cnt2 = data_v2['wash_purchases_total_cnt']
    retail_revenue2 = data_v2['retail_revenue']
    total_revenue2 = data_v2['total_revenue']
    reedeemd_total_cnt2 = data_v2['reedeemd_total_cnt']
    arm_plans_sold2 = data_v2['arm_plans_sold']
    
    # wash_purchases_total_cnt2 = 0
    # reedeemd_total_cnt2 = 0
    # retail_revenue2=0.0
    # total_revenue2 = 0.0
    # arm_plans_sold2 = 0
    
    # for item in items :
    #     itemtyp = item.get("ItemType")
    #     discount = item.get("Discount")
    #     flag= item.get("Flag")
    #     price = item.get("Price")
        
    #     if flag:
    #         reedeemd_total_cnt2+=1
        
    #     elif itemtyp=="Wash" and not (flag or discount): #wash purchase
    #         wash_purchases_total_cnt2+=1
    #         retail_revenue2+=price
        
    #     if not (flag or   discount): 
    #         total_revenue2+=price
            
    #     if itemtyp in ["WashClubReactivation","WashClubSignUp","AppWashClubSignUp"] :# WashClubSignUp,  # arm plans sold 
    #         arm_plans_sold2+=1
            
    final_data["car_count_saturday_sunday"] = sum([wash_purchases_total_cnt2,reedeemd_total_cnt2])
    final_data["arm_plans_reedemed_saturday_sunday"] = "" #update
    final_data["retail_car_count_saturday_sunday"]   = wash_purchases_total_cnt2
    final_data["retail_revenue_saturday_sunday"]    = retail_revenue2
    final_data["total_revenue_saturday_sunday"]    = total_revenue2
    final_data["labour_hours_saturday_sunday"]     = ""
    final_data["cars_per_labour_hour_saturday_sunday"] = ""
    
    arm_plans_sold_cnt = sum([arm_plans_sold1,arm_plans_sold2])
    
    total_arm_planmembers_cnt = client.get_total_plan_members(sunday_date_str)
    
    
    #past 4 weeks data 
    past_4_week_day1,past_4_week_day2 = generate_past_4_weeks_days(monday_date_str)
    items = client.get_dail_report_v2(past_4_week_day1,past_4_week_day2)

    data_v3 = find_retail_revenue_and_total_revenue(items)
    
    wash_purchases_total_cnt3 = data_v3['wash_purchases_total_cnt']
    reedeemd_total_cnt3 = data_v3['reedeemd_total_cnt']
    retail_revenue3= data_v3['retail_revenue']
    total_revenue3 = data_v3['total_revenue']
    arm_plans_sold3 = data_v3['arm_plans_sold']
    
    # for item in items :
    #     itemtyp = item.get("ItemType")
    #     discount = item.get("Discount")
    #     flag= item.get("Flag")
    #     price = item.get("Price")
        
    #     if flag:
    #         reedeemd_total_cnt3+=1
        
    #     elif itemtyp=="Wash" and not (flag or discount): #wash purchase
    #         wash_purchases_total_cnt3+=1
    #         retail_revenue3+=price
        
    #     if not (flag or   discount): 
    #         total_revenue3+=price
            
    #     if itemtyp in ["WashClubReactivation","WashClubSignUp","AppWashClubSignUp"] :# WashClubSignUp,  # arm plans sold 
    #         arm_plans_sold3+=1
    
     
    past_4_week_cnt = sum([wash_purchases_total_cnt3,reedeemd_total_cnt3])
    final_data["past_4_week_cnt"] = past_4_week_cnt
    final_data["past_4_week_conversion_rate"] = conversion_rate_hamilton(arm_plans_sold3,wash_purchases_total_cnt3,0)
    final_data["past_4_weeks_total_revenue"] =total_revenue3
    final_data["past_4_weeks_arm_plans_sold_cnt"] = arm_plans_sold3
    final_data["past_4_weeks_retail_car_count"]  = wash_purchases_total_cnt3
    
    final_data["past_4_week_car_cnt_mon_fri"]=0
    # final_data["past_4_week_labour_hours_mon_fri"]=0
    
    final_data["past_4_week_car_cnt_sat_sun"]=0
    # final_data["past_4_week_labour_hours_sat_sun"]=0

    final_data["past_4_week_retail_car_count_mon_fri"]=0
    final_data["past_4_week_retail_car_count_sat_sun"]=0

    final_data['past_4_week_retail_revenue_mon_fri'] = 0
    final_data['past_4_week_retail_revenue_sat_sun'] = 0

    final_data['past_4_week_total_revenue_mon_fri'] = 0
    final_data['past_4_week_total_revenue_sat_sun'] = 0
    
    full_weeks_lst = generate_past_4_week_days_full(monday_date_str)
    cnt=0
    for single_week in full_weeks_lst:
        mon = single_week[0]
        fri = single_week[1]
        sat =single_week[2]
        sun = single_week[3]

        items = client.get_dail_report_v2(mon,fri)
        past_week_car_count_mon_fri = client.get_car_count(items)

        data_mon_fri = find_retail_revenue_and_total_revenue(items)
        wash_purchases_total_cnt4 = data_mon_fri['wash_purchases_total_cnt']
        reedeemd_total_cnt4 = data_mon_fri['reedeemd_total_cnt']
        retail_revenue4= data_mon_fri['retail_revenue']
        total_revenue4 = data_mon_fri['total_revenue']
        arm_plans_sold4 = data_mon_fri['arm_plans_sold']

        final_data["past_4_week_car_cnt_mon_fri"] = final_data.get("past_4_week_car_cnt_mon_fri",0) +  past_week_car_count_mon_fri
        final_data["past_4_week_retail_car_count_mon_fri"] += wash_purchases_total_cnt4
        final_data['past_4_week_retail_revenue_mon_fri'] += retail_revenue4
        final_data['past_4_week_total_revenue_mon_fri'] += total_revenue4
        
        final_data[f'past_4_week_retail_revenue_mon_fri_week_{cnt+1}'] = retail_revenue4
        final_data[f"past_4_week_retail_car_count_mon_fri_week_{cnt+1}"] = wash_purchases_total_cnt4
        print(f"past_4_week_retail_revenue_mon_fri:{retail_revenue4}")
        

        items2 = client.get_dail_report_v2(sat,sun)
        past_week_car_count_sat_sun = client.get_car_count(items2)

        data_sat_sun = find_retail_revenue_and_total_revenue(items2)
        wash_purchases_total_cnt5 = data_sat_sun['wash_purchases_total_cnt']
        reedeemd_total_cnt5 = data_sat_sun['reedeemd_total_cnt']
        retail_revenue5= data_sat_sun['retail_revenue']
        total_revenue5 = data_sat_sun['total_revenue']
        arm_plans_sold5 = data_sat_sun['arm_plans_sold']

        final_data["past_4_week_car_cnt_sat_sun"] = final_data.get("past_4_week_car_cnt_sat_sun",0) + past_week_car_count_sat_sun
        final_data["past_4_week_retail_car_count_sat_sun"] += wash_purchases_total_cnt5
        final_data['past_4_week_retail_revenue_sat_sun'] += retail_revenue5
        final_data['past_4_week_total_revenue_sat_sun'] += total_revenue5
        print(f"past_4_week_retail_revenue_sat_sun : {retail_revenue5}")
        
        final_data[f'past_4_week_retail_revenue_sat_sun_week_{cnt+1}'] = retail_revenue5
        final_data[f"past_4_week_retail_car_count_sat_sun_week_{cnt+1}"] = wash_purchases_total_cnt5
        # print(f"final check1 : {final_data[f'past_4_week_retail_revenue_sat_sun_week_{cnt+1}']}")
        cnt+=1
    
    print(f"past week cnt : {past_4_week_cnt}")
    
    final_data["total_revenue"] = sum([total_revenue,total_revenue2])
    final_data["arm_plans_sold_cnt"] = arm_plans_sold_cnt
    final_data["total_arm_planmembers_cnt"] = total_arm_planmembers_cnt
    final_data["conversion_rate"] = conversion_rate_hamilton(arm_plans_sold_cnt,wash_purchases_total_cnt,wash_purchases_total_cnt2)
    # final_data[""]
    place_format = {}
    place_format["Splash-Peoria"] = final_data
    
    print(place_format)
    
    return  place_format
  


                
        
            
            
    



if __name__ == "__main__":
    proxy_url = None

    proxy = {"http": proxy_url, "https": proxy_url}
    client = hamiltonClient()
    login_data = {"UserName": "CR@Sparklecw.com", "Password": "CameronRay1"}
    login = client.login(login_data, proxy)
    # print(f"login:{login}")
    #daily_report = client.get_daily_report(proxy)
    #print(f"daily :{daily_report}")
    #rev = client.get_revenue("05/22/2024", "06/23/2024")
    # client.to_csv(rev, "revenue.csv")
    # client.to_excel(rev, "revenue.xlsx")
    #print(rev)
    monday_date_str = "2024-06-03"
    friday_date_str = "2024-06-07"
    saturday_date_str = "2024-06-08"
    sunday_date_str  = "2024-06-09"
    
    monday_date_str = "2024-07-22"
    friday_date_str = "2024-07-26"
    saturday_date_str = "2024-07-27"
    sunday_date_str  = "2024-07-28"
    # dail_report_v2 = client.get_dail_report_v2(monday_date_str,friday_date_str)
    
    # print(f"Daily report v2 : {dail_report_v2}")
    
    # with open("hamiltin_data.json","w") as f:
    #     json.dump(dail_report_v2,f,indent=4)
    
    hamilton_report = generate_report(monday_date_str, friday_date_str, saturday_date_str, sunday_date_str)
    
    with open("hamiltin_data.json","w") as f:
        json.dump(hamilton_report,f,indent=4)
    
    print(client.get_total_plan_members("2024-07-07"))

