import random
import time
import requests
import pickle
import os
from datetime import datetime, timedelta,timezone
import json
from dotenv import load_dotenv

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

print(f"proxies in sitewatch:{proxies}")
current_file_path = os.path.dirname(os.path.abspath(__file__))
# print(current_file_path)

cookies_path = os.path.join(current_file_path,"cookies")

cookies_file = os.path.join(cookies_path,"sitewatch_cookies.pkl")



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

def generate_past_4_weeks_days(date_str):
    # Convert the string date to a datetime object
    date_format = "%Y-%m-%d"
    monday = datetime.strptime(date_str, date_format)
    
    # Subtract one day
    one_day_before = monday - timedelta(days=1)
    four_weeks_before = monday - timedelta(days=(7*4) + 1)

    # Format the dates in "dd/mm/yyyy" format
    formatted_date = one_day_before.strftime("%Y-%m-%d")
    four_weeks_before_fmt = four_weeks_before.strftime("%Y-%m-%d")

    print("One day before the current date:", formatted_date)
    print("4 weeks before day :", four_weeks_before_fmt)

    return four_weeks_before_fmt, formatted_date


def generate_heartbeatID():
    return round(random.random() * 1e7)

def generate_cb_value():
    return int(time.time() * 1000)





class sitewatchClient():
    def __init__(self,cookies_file) -> None:
        self.token = None
        self.heartbeatID =generate_heartbeatID()
        self.cb_value  = generate_cb_value()
        self.cookies_file = cookies_file
        self.proxies= proxies

    def login(self,employeeCode,password,locationCode,remember=0,timeout=60):
        """used for loginto the site and return jwt token for further processing

        Args:
            employeeCode (str): employee code 
            password (str): password
            locationCode (str):code for the place 
            remember (int, optional): remember login. Defaults to 0. not used
        """
        session = requests.Session()
        token = None
        headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'en-US,en;q=0.9',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded',
        # 'Cookie': '_ga=GA1.2.2058823687.1718534132; _gid=GA1.2.1004888831.1718534132; _gat=1',
        'DNT': '1',
        'Origin': 'https://sitewatch.cloud',
        'Pragma': 'no-cache',
        'Referer': 'https://sitewatch.cloud/remote/',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        }

        data = {
            'employeeCode': employeeCode,
            'password': password,
            'locationCode': locationCode,
            'remember': str(remember),
        }

        try:
            response = session.post('https://sitewatch.cloud/api/auth/authenticate', headers=headers, data=data,timeout=timeout,proxies=self.proxies)
            print("Login response:",response)
            if response.status_code==200:
                token = response.json().get("token")
                self.token=token
                # print(session.cookies)
                print(response.json())
                with open(self.cookies_file,'wb') as f:
                        pickle.dump(session.cookies,f)
            else:
                print(response.status_code,response.json())
        except Exception as e:
            print(f"Exception in login : {e}")
        return token

    def check_session_auth(self,timeout=60)->bool:
        """chekcs whether session is authenticated

        Returns:
            bool: authenticated (True or False)
        """ 
        # cookies = {
        #     '_ga': 'GA1.2.2058823687.1718534132',
        #     '_gid': 'GA1.2.1004888831.1718534132',
        #     'token': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiYjNkOGI3ZDYtOTQ5NC00NzY5LWJkNzMtZjY1YmI4MjIwNGZjIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDEiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcxODYyNjAzNn0.j6kjJM9g01tUwT-Uc43ehdF0AqH-72g8CwXswRtzAvY',
        #     '_gat': '1',
        # }

        headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'en-US,en;q=0.9',
            #'Authorization': f'Bearer {self.token}',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            #'Cookie': '_ga=GA1.2.2058823687.1718534132; _gid=GA1.2.1004888831.1718534132; token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiYjNkOGI3ZDYtOTQ5NC00NzY5LWJkNzMtZjY1YmI4MjIwNGZjIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDEiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcxODYyNjAzNn0.j6kjJM9g01tUwT-Uc43ehdF0AqH-72g8CwXswRtzAvY; _gat=1',
            'DNT': '1',
            'Pragma': 'no-cache',
            'Referer': 'https://sitewatch.cloud/',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }
        session  = requests.Session()
        
        with open(self.cookies_file,'rb') as f:
            cookies = pickle.load(f)

        session.cookies = cookies
        session.headers = headers
        
        params = {
            'cb': str(self.cb_value),
            'heartbeatID': str(self.heartbeatID),
        }
        authenticated=False
        try:
            response = session.get('https://sitewatch.cloud/api/auth/session',params=params,timeout=timeout,proxies=self.proxies)
            # print(response.json())
            if response.status_code==200:
                authenticated = response.json().get("authenticated")
                authenticated=True
                print("authentication success")
        except Exception as e:
            print(f"Exception in check_session_auth : {e} ")
            
        return authenticated


    def get_general_sales_report_request_id(self,reportOn,id,name,monday_date_str, sunday_date_str,timeout=60):
        data = None

        """This function will get general sales report

        Args:
            reportOn (_type_): _description_
            requestID (_type_): _description_
        """


        # cookies = {
        #     '_ga': 'GA1.2.2058823687.1718534132',
        #     '_gid': 'GA1.2.1004888831.1718534132',
        #     'token': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiNmI4YjY1MTItNWU4ZS00YzU1LTk1N2UtMzJkMTY0NjAwNWUwIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDEiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcxODYzNDk0MX0.PW-vPlDIAcJgDvofcv-wCfgKbnc54N6uL_w_S3S9xnE',
        # }

        headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'en-US,en;q=0.9',
            #'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiNmI4YjY1MTItNWU4ZS00YzU1LTk1N2UtMzJkMTY0NjAwNWUwIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDEiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcxODYzNDk0MX0.PW-vPlDIAcJgDvofcv-wCfgKbnc54N6uL_w_S3S9xnE',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            'Content-Type': 'application/json;charset=UTF-8',
            # 'Cookie': '_ga=GA1.2.2058823687.1718534132; _gid=GA1.2.1004888831.1718534132; token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiNmI4YjY1MTItNWU4ZS00YzU1LTk1N2UtMzJkMTY0NjAwNWUwIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDEiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcxODYzNDk0MX0.PW-vPlDIAcJgDvofcv-wCfgKbnc54N6uL_w_S3S9xnE',
            'DNT': '1',
            'Origin': 'https://sitewatch.cloud',
            'Pragma': 'no-cache',
            'Referer': 'https://sitewatch.cloud/',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        session  = requests.Session()
        
        with open(self.cookies_file,'rb') as f:
            cookies = pickle.load(f)

        session.cookies = cookies
        session.headers = headers
        params = {
            'allowCallback': '1',
            'heartbeatID': str(self.heartbeatID),
            'reportOn': reportOn,
        }

       

        # Format both times
        one_hour_before_formatted = f"{monday_date_str}T00:00:00"
        current_time_formatted = f"{sunday_date_str}T23:59:59"
        #one_hour_before_formatted = "2024-06-14T00:00:00"
        #current_time_formatted ="2024-06-14T23:59:59"
        
        json_data = {
            'startDate':  one_hour_before_formatted,
            'endDate': current_time_formatted,
            'shifts': [],
            'salesRoles': [],
            'terminals': [],
            'format': {
                'id': id,
                'name': name,
                'title': '1',
                'indentOffset': 2,
                'indentSpaces': 3,
            },
            'employees': [],
            'showEachShift': False,
            'showEachTerminal': False,
            'showEachSite': False,
            'paperSize': 'letter',
            'paginationOffset': None,
        }

        try:
            response = session.post(
            'https://sitewatch.cloud/api/gsreport/gsreport',
                params=params,
                json=json_data,
                timeout=timeout,proxies=self.proxies
            )
            if response.status_code==200:
                data = response.json().get("requestID")
            # print(f"sales id response :{response} , {response.json()}")
        except Exception as e:
            print(f"Exception in get_general_sales_report: {e} ")
        
        return data


    def get_report(self,reportOn,requestID,timeout=60):
        data = None

        # cookies = {
        #     '_ga': 'GA1.2.2058823687.1718534132',
        #     '_gid': 'GA1.2.1004888831.1718534132',
        #     'token': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiY2YxOTlhMjUtYWYzMC00ZDBlLWFhNzUtY2QyY2IzMWRlZTcxIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDEiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcyMTE0MTI3NX0.Zn74SEnlCHddFE5ARZ1YX-hJkzvsBqYowAkVta_1dd4',
        #     '_gat': '1',
        # }

        headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'en-US,en;q=0.9',
            #'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiY2YxOTlhMjUtYWYzMC00ZDBlLWFhNzUtY2QyY2IzMWRlZTcxIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDEiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcyMTE0MTI3NX0.Zn74SEnlCHddFE5ARZ1YX-hJkzvsBqYowAkVta_1dd4',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            # 'Cookie': '_ga=GA1.2.2058823687.1718534132; _gid=GA1.2.1004888831.1718534132; token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiY2YxOTlhMjUtYWYzMC00ZDBlLWFhNzUtY2QyY2IzMWRlZTcxIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDEiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcyMTE0MTI3NX0.Zn74SEnlCHddFE5ARZ1YX-hJkzvsBqYowAkVta_1dd4; _gat=1',
            'DNT': '1',
            'Pragma': 'no-cache',
            'Referer': 'https://sitewatch.cloud/',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        params = {
            'cb': str(self.cb_value),
            'heartbeatID': str(self.heartbeatID),
            'reportOn': reportOn,
            'requestID': requestID,
        }
        session  = requests.Session()
        
        with open(self.cookies_file,'rb') as f:
            cookies = pickle.load(f)
        session.cookies =cookies
        session.headers = headers
        try:
            response = session.get('https://sitewatch.cloud/api/request/results', params=params,timeout=timeout,proxies=self.proxies)
            if response.status_code==200:
                data = response.json()
            print("response:",response)
            # print("headers:",session.headers)
            # print(response.json())
            # with open(f"{requestID}.json","w") as f:
            #     #json.dump(response.json(),f,indent=4)
            # print(f"get report {response} , {response.json()}")
        except Exception as e:
            print(f"Exception in get_report: {e} ")
        
        return data


    def get_activity_by_date_proft_request_id(self,reportOn,startDate,endDate,timeout=60):
        requestID =None

        headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'en-US,en;q=0.9',
            #'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiZDA5NWNlMzUtZDBlMi00YWJlLWI3NzQtYjlhM2U3YjIzYTliIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDkiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcyMTg5NDQzN30.tn0aqhRbU6qkSbwIet7C8tgep_TW7XqHl54_fJrTKsc',
            'Connection': 'keep-alive',
            # 'Cookie': '_ga=GA1.2.938265008.1718553024; _gid=GA1.2.1473468462.1719290971; token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiZGJhZTAzOGEtYzhmOS00NDUzLWI0OTYtNTg1NmUyYTY0NDkwIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDIiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcxOTU2NzA1N30.TcoGK3UZ0v9enWREhpizjrGFvBw6IAHmjTpvyekXGFA; _gat=1',
            'DNT': '1',
            'Referer': 'https://sitewatch.cloud/',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }
        session  = requests.Session()
        
        with open(self.cookies_file,'rb') as f:
            cookies = pickle.load(f)

        session.cookies = cookies
        session.headers = headers    
        params = {
            'cb': generate_cb_value(),
            'activeView': 'custom',
            'allowCallback': '1',
            'endDate': endDate, # '2024-06-09'
            'heartbeatID': generate_heartbeatID,
            'paperSize': 'letter',
            'reportOn': reportOn,
            'startDate': startDate, # 2024-06-03
        }

        try:
            response = session.get(
            'https://sitewatch.cloud/api/activity-by-date-profit-center',
            params=params,
            headers=headers,timeout=timeout,proxies=self.proxies
        )
            if response.status_code==200:
                requestID = response.json().get("requestID")
            # print(f"response: in activity report {response}, {response.json()}")
        except Exception as e:
            print(f"Exception in get_activity_by_date_proft_request_id() {e}")

        return requestID
    
    def get_labour_hours(self,reportOn,requestID,):
        
        laborHours=None

        headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'en-US,en;q=0.9',
            #'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiZDA5NWNlMzUtZDBlMi00YWJlLWI3NzQtYjlhM2U3YjIzYTliIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDkiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcyMTg5NDQzN30.tn0aqhRbU6qkSbwIet7C8tgep_TW7XqHl54_fJrTKsc',
            'Connection': 'keep-alive',
            # 'Cookie': '_ga=GA1.2.938265008.1718553024; _gid=GA1.2.1473468462.1719290971; token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiZGJhZTAzOGEtYzhmOS00NDUzLWI0OTYtNTg1NmUyYTY0NDkwIiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDIiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcxOTU2NzA1N30.TcoGK3UZ0v9enWREhpizjrGFvBw6IAHmjTpvyekXGFA; _gat=1',
            'DNT': '1',
            'Referer': 'https://sitewatch.cloud/',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }

        session  = requests.Session()
                
        with open(self.cookies_file,'rb') as f:
            cookies = pickle.load(f)

            session.cookies = cookies
            session.headers = headers 
        params = {
            'cb': self.cb_value,
            'heartbeatID': self.heartbeatID,
            'reportOn': reportOn,
            'requestID': requestID,
        }

        try:
            response = session.get('https://sitewatch.cloud/api/request/results', params=params, headers=headers,proxies=self.proxies)
            if response.status_code==200:
                data = response.json()
                profitCenterData = data.get("profitCenterData")[0]
                laborHours= round(profitCenterData.get("laborHours",0.0),2)
        except Exception as e :
            print(f"exception in get_labour_hours() {e}")

    
        return laborHours

    def get_plan_analysis_request_id(self,date,reportOn,timeout=60):
        request_id = None
        try:
            headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'en-US,en;q=0.9',
            #'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjIuMjAwIiwic3R6IjoiLTA1OjAwOjAwIiwianRpIjoiOTMyNWZmNmQtZWQwNi00NzViLTg1YzctYTUzY2U3NWM5ZjJmIiwiZW1wIjoyMiwiZWlkIjozMDk1MDAwMDEsImxvYyI6IlNEWldJTC0wMDIiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcyMDg1NjYyNX0.SgZbMMmFxr95LP5FqQgOfQ4ktbq6xfQFNEQveK8pAEE',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            # 'Cookie': '_ga=GA1.2.938265008.1718553024; _gid=GA1.2.170212385.1720769994; token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjIuMjAwIiwic3R6IjoiLTA1OjAwOjAwIiwianRpIjoiOTMyNWZmNmQtZWQwNi00NzViLTg1YzctYTUzY2U3NWM5ZjJmIiwiZW1wIjoyMiwiZWlkIjozMDk1MDAwMDEsImxvYyI6IlNEWldJTC0wMDIiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcyMDg1NjYyNX0.SgZbMMmFxr95LP5FqQgOfQ4ktbq6xfQFNEQveK8pAEE',
            'DNT': '1',
            'Pragma': 'no-cache',
            'Referer': 'https://sitewatch.cloud/',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            }
            session  = requests.Session()
            
            with open(self.cookies_file,'rb') as f:
                cookies = pickle.load(f)

            session.cookies = cookies
            session.headers = headers  
            now = datetime.now(timezone.utc)
            formatted_time = now.strftime("T%H:%M:%S.%f")[:-3] + 'Z'
            date_time = f"{date}T22:00:00.000Z"  #2024-07-07T07:53:04.769Z
            url = f'https://sitewatch.cloud/api/pass-report/analysis?cb={self.cb_value}&allowCallback=1&date={date_time}&heartbeatID={self.heartbeatID}&level=summary&paperSize=letter&period=month&reportOn={reportOn}'
            response = session.get(url,timeout=timeout,proxies=self.proxies)
            print(f"response : {response}",response.json())
            if response.status_code==200:
                request_id = response.json().get("requestID")
        
        except Exception as e:
            print(f"exception in get_analysis_request_id() {e}")
            
        return request_id
        
    def get_total_plan_members(self,requestID,reportOn,timeout=60):
        total_members = None
        
        headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'en-US,en;q=0.9',
            #'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjIuMjAwIiwic3R6IjoiLTA1OjAwOjAwIiwianRpIjoiOTMyNWZmNmQtZWQwNi00NzViLTg1YzctYTUzY2U3NWM5ZjJmIiwiZW1wIjoyMiwiZWlkIjozMDk1MDAwMDEsImxvYyI6IlNEWldJTC0wMDIiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcyMDg1NjYyNX0.SgZbMMmFxr95LP5FqQgOfQ4ktbq6xfQFNEQveK8pAEE',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            # 'Cookie': '_ga=GA1.2.938265008.1718553024; _gid=GA1.2.170212385.1720769994; token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjIuMjAwIiwic3R6IjoiLTA1OjAwOjAwIiwianRpIjoiOTMyNWZmNmQtZWQwNi00NzViLTg1YzctYTUzY2U3NWM5ZjJmIiwiZW1wIjoyMiwiZWlkIjozMDk1MDAwMDEsImxvYyI6IlNEWldJTC0wMDIiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcyMDg1NjYyNX0.SgZbMMmFxr95LP5FqQgOfQ4ktbq6xfQFNEQveK8pAEE',
            'DNT': '1',
            'Pragma': 'no-cache',
            'Referer': 'https://sitewatch.cloud/',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }  
        session  = requests.Session()
            
        with open(self.cookies_file,'rb') as f:
            cookies = pickle.load(f)

        session.cookies = cookies
        session.headers = headers  
        
        try:
            params = {
            'cb': self.cb_value,
            'heartbeatID': self.heartbeatID,
            'reportOn': reportOn,
            'requestID': requestID,
            }
            
            response = session.get('https://sitewatch.cloud/api/request/results', params=params,proxies=self.proxies)
            print("resp  in total memebers :",response.status_code)
            if response.status_code==200:
               
                statistics = response.json().get("statistics")
                month  = statistics.get("month")
                total_members = month.get("endingMembers")
                
        except Exception as e :
            print(f"Exception in get_total_plan_members() {e}")
        
        return total_members
        

if __name__=="__main__":
    # print("HeartBeatID :",generate_heartbeatID())
    # print("cb_value :",generate_cb_value())

    client = sitewatchClient(cookies_file)
    employCode = "20"
    password = 'Cameron1'
    locationCode = 'SPKLUS-002'
    remember = 1
    client.login(employeeCode=employCode,password=password,locationCode=locationCode,remember=remember)
    # print(client.token)
    # print(client.cb_value)
    # print(client.heartbeatID)
    # print(client.check_session_auth())
    # client.cb_value=1718542975922
    # client.heartbeatID=5862273
    #client.token="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ2ZXIiOiIyNy4zLjQuMjM0Iiwic3R6IjoiLTA0OjAwOjAwIiwianRpIjoiNGQ1YzQ0MDUtYmFiMC00OGNlLThkNTEtMzI0NTk1NDE1MmI0IiwiZW1wIjoyMCwiZWlkIjoyMDAyMDAwMDAsImxvYyI6IlNQS0xVUy0wMDEiLCJhc3Npc3RlZCI6ZmFsc2UsImV4cCI6MTcyMTEzNzI4MX0.7OOJ8B_SGyuVmX2aQjQtbljpeX_JiVtlSCini5MeQOQ"
    # print(client.token)
    session_check = client.check_session_auth()
    print(session_check)
    print(type(session_check))
    print(client.token)
    reportOn  = "site-2"
    # print(client.get_requestid(reportOn=reportOn))
    # req_id = client.get_general_sales_report_request_id(reportOn,2121400001,'Site Financial Detail & Chem Report-2021','2024-06-03','2024-06-09')

    
    # report_data = client.get_report(reportOn,req_id)
    
    # print(report_data)
    # print('Testing data ')
    # req_id2 = client.get_activity_by_date_proft_request_id(reportOn,'2024-06-03','2024-06-09')
    # print(req_id2)
    # print(client.get_labour_hours(reportOn,req_id2))
    
    date= "2024-07-07"
    req_id = client.get_plan_analysis_request_id(date,reportOn)
    print("reqid:",req_id)
    print(client.get_total_plan_members(req_id,reportOn))
    

        

#Application working flow

# first login then get jwt token

# nest check session is authenticated or not suing another api

#for doing request for reports 
# for getting request id use https://sitewatch.cloud/api/self-info?cb=1718539644149&allowCallback=1&heartbeatID=6641209

# change log 
# - added headers in each fucntion


## Improvement docs
# removing unnecesary comments
# adding auto time for grepting payload
