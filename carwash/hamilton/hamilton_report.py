import requests
import pickle
import os
import json
from datetime import datetime, timedelta
from dotenv import load_dotenv
import logging
import time


current_file_path = os.path.dirname(os.path.abspath(__file__))
cookies_path = os.path.join(current_file_path, "cookies")
cookie_file_path = os.path.join(cookies_path, "cookie.json")

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
    IS_PROXY           = env_vars.get("IS_PROXY")

username = Config.PROXY_USER_NAME
password = Config.PROXY_PASSWORD
zone = Config.PROXY_ZONE

# Proxy configuration
proxy_host = Config.PROXY_HOST
proxy_port = Config.PROXY_PORT
IS_PROXY = Config.IS_PROXY

# Proxy URL format for datacenter proxy
proxies = None

# print(IS_PROXY)
if IS_PROXY:
    proxy_url = f'http://{username}-zone-{zone}:{password}@{proxy_host}:{proxy_port}'
    proxies={"http":proxy_url,"https":proxy_url}


class hamiltonClient:

    def __init__(self) -> None:
        self.proxies = proxies


    def login(self, login_data: dict, proxy) -> bool:
        while True:
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
                logging.info(f"Exception in hamilton login {e} ")
                logging.info("sleeping 5 secounds before next retry")
                time.sleep(5)

        return False


    def get_ccokies(self):
        with open(cookie_file_path, "rb") as f:
            cookies = pickle.load(f)
        return cookies


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

        while True:
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
                logging.info("retrying after 5 secounds ")
                time.sleep(5)


    def get_days_for_Total_membership(self,end_date):

        # Get the old date
        current_date = datetime.strptime(end_date, "%Y-%m-%d") #.now()#

        # Get the date 30 days before the current date
        date_30_days_ago = current_date - timedelta(days=30)

        # Format both dates to the desired format
        current_date_str = current_date.strftime('%Y-%m-%d')
        date_30_days_ago_str = date_30_days_ago.strftime('%Y-%m-%d')

        print(f"Current Date: {current_date_str}")
        print(f"Date 30 Days Ago: {date_30_days_ago_str}")

        return current_date_str,date_30_days_ago_str


    def get_total_plan_members(self,end_date):
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

        endDate,startDate = self.get_days_for_Total_membership(end_date)
        json_data = {
            'startDate': startDate,
            'endDate': endDate,
        }

        while True:
            total_plan_members=0
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
                    break

            except Exception as e:
                print(f"Exception in get_total_plan_members()  in {e}")
                logging.info(f"Exception in get_total_plan_members()  in {e}")
            logging.info(f"sleeping 5 secounds before retry ")
            time.sleep(5)

        return total_plan_members


def conversion_rate_hamilton(arm_plans_sold_cnt,wash_purchases_total_cnt):
    rate = 0
    try:
        rate = arm_plans_sold_cnt/ wash_purchases_total_cnt
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

        if not (flag or   discount): #total revenue
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


def generate_report(start_date_current_year,end_date_current_year,start_date_last_year, end_date_last_year):

    logger = logging.getLogger(__name__)
    logger.info("started main script")
    final_data = {}
    proxy_url = None

    proxy = {"http": proxy_url, "https": proxy_url}
    client = hamiltonClient()
    login_data = {"UserName": "CR@Sparklecw.com", "Password": "CameronRay1"}
    login = client.login(login_data, proxy)

    items = client.get_dail_report_v2(start_date_current_year,end_date_current_year)

    data_v1 = find_retail_revenue_and_total_revenue(items)
    wash_purchases_total_cnt = data_v1['wash_purchases_total_cnt']
    retail_revenue = data_v1['retail_revenue']
    total_revenue = data_v1['total_revenue']
    reedeemd_total_cnt = data_v1['reedeemd_total_cnt']
    arm_plans_sold1 = data_v1['arm_plans_sold']

    final_data["car_count_current_year"] = sum([wash_purchases_total_cnt,reedeemd_total_cnt])
    final_data["arm_plans_reedemed_current_year_cnt"]  = "" #update
    final_data["retail_car_count_current_year"] = wash_purchases_total_cnt
    final_data["retail_revenue_current_year"] = retail_revenue
    final_data["total_revenue_current_year"] =  total_revenue
    final_data["labour_hours_current_year"]  = ""
    final_data["cars_per_labour_hour_current_year"] = ""
    final_data["total_revenue_current_year"] = total_revenue
    final_data["arm_plans_sold_cnt_current_year"] = arm_plans_sold1

    #for last year
    items = client.get_dail_report_v2(start_date_last_year, end_date_last_year)

    data_v2 = find_retail_revenue_and_total_revenue(items)
    wash_purchases_total_cnt2 = data_v2['wash_purchases_total_cnt']
    retail_revenue2 = data_v2['retail_revenue']
    total_revenue2 = data_v2['total_revenue']
    reedeemd_total_cnt2 = data_v2['reedeemd_total_cnt']
    arm_plans_sold2 = data_v2['arm_plans_sold']

    final_data["car_count_last_year"] = sum([wash_purchases_total_cnt2,reedeemd_total_cnt2])
    final_data["arm_plans_reedemed_last_year"] = "" #update
    final_data["retail_car_count_last_year"]   = wash_purchases_total_cnt2
    final_data["retail_revenue_last_year"]    = retail_revenue2
    final_data["total_revenue_last_year"]    = total_revenue2
    final_data["labour_hours_last_year"]     = ""
    final_data["cars_per_labour_hour_last_year"] = ""
    final_data["total_revenue_last_year"] = total_revenue2
    final_data["arm_plans_sold_cnt_last_year"] = arm_plans_sold2

    total_arm_planmembers_cnt_current_year = client.get_total_plan_members(end_date_current_year)
    total_arm_planmembers_cnt_last_year = client.get_total_plan_members(end_date_last_year)

    final_data["total_arm_planmembers_cnt_current_year"] = total_arm_planmembers_cnt_current_year
    final_data["total_arm_planmembers_cnt_last_year"] = total_arm_planmembers_cnt_last_year
    final_data["conversion_rate_current_year"] = conversion_rate_hamilton(arm_plans_sold1,wash_purchases_total_cnt)
    final_data["conversion_rate_last_year"] = conversion_rate_hamilton(arm_plans_sold2,wash_purchases_total_cnt2)

    # final_data[""]
    place_format = {}
    place_format["Splash-Peoria"] = final_data

    print(place_format)
    logger.info("final data")
    logger.info(f"{ place_format}")

    return  place_format


if __name__ == "__main__":

    start_date_current_year = "2024-11-01"
    end_date_current_year = "2024-11-14"

    start_date_last_year = "2023-11-01"
    end_date_last_year = "2023-11-14"

    hamilton_report = generate_report(start_date_current_year,end_date_current_year,start_date_last_year, end_date_last_year)

    with open("test_hamiltin_data.json","w") as f:
        json.dump(hamilton_report,f,indent=4)
