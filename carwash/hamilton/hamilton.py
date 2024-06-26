import requests
import pickle
from datetime import date
import os
import json
import datetime as dt
from bs4 import BeautifulSoup
import pandas as pd

import csv
current_file_path = os.path.dirname(os.path.abspath(__file__))
cookies_path = os.path.join(current_file_path, "cookies")

cookie_file_path = os.path.join(cookies_path, "cookie.json")

data2 = os.path.join(current_file_path, "data2.json")


class hamiltonClient:
    def __init__(self) -> None:
        pass

    def login(self, login_data: dict, proxy) -> bool:
        session = requests.Session()
        try:
            response = session.post(
                "https://hamiltonservices.com/web/",
                data=login_data,
                proxies={"http": proxy},
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


if __name__ == "__main__":
    proxy_url = None

    proxy = {"http": proxy_url, "https": proxy_url}
    client = hamiltonClient()
    login_data = {"UserName": "CR@Sparklecw.com", "Password": "CameronRay1"}
    login = client.login(login_data, proxy)
    # print(f"login:{login}")
    # daily_report = client.get_daily_report(proxy)
    # print(f"daily :{daily_report}")
    rev = client.get_revenue("05/22/2024", "06/23/2024")
    # client.to_csv(rev, "revenue.csv")
    client.to_excel(rev, "revenue.xlsx")
    print(rev)

