import random
import time
import requests
import pickle
import os
from datetime import datetime, timedelta
import json

current_file_path = os.path.dirname(os.path.abspath(__file__))
# print(current_file_path)

cookies_path = os.path.join(current_file_path,"cookies")

cookies_file = os.path.join(cookies_path,"sitewatch_cookies.pkl")

def generate_heartbeatID():
    return round(random.random() * 1e7)

def generate_cb_value():
    return int(time.time() * 1000)





class sitewatchClient():
    def __init__(self,cookies_file) -> None:
        self.token = None
        self.heartbeatID =generate_heartbeatID()
        self.cb_value  = generate_cb_value
        self.cookies_file = cookies_file

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
            response = session.post('https://sitewatch.cloud/api/auth/authenticate', headers=headers, data=data,timeout=timeout)
            # print("Login response:",response)
            if response.status_code==200:
                token = response.json().get("token")
                self.token=token
                # print(session.cookies)
                # print(response.json())
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
            response = session.get('https://sitewatch.cloud/api/auth/session',params=params,timeout=timeout)
            # print(response.json())
            if response.status_code==200:
                authenticated = response.json().get("authenticated")
                authenticated=True
        except Exception as e:
            print(f"Exception in check_session_auth : {e} ")
            
        return authenticated


    def get_general_sales_report_request_id(self,reportOn,id,name,timeout=60):
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

        # Get the current time
        current_time = datetime.now()

        # Get the time one hour before the current time
        one_hour_before = current_time - timedelta(hours=1)

        # Define the desired format
        format_string = '%Y-%m-%d'

        # Format both times
        one_hour_before_formatted = f"{one_hour_before.strftime(format_string)}T00:00:00"
        current_time_formatted = f"{current_time.strftime(format_string)}T23:59:59"
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
                timeout=timeout
            )
            if response.status_code==200:
                data = response.json().get("requestID")
            # print(data)
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
            response = session.get('https://sitewatch.cloud/api/request/results', params=params,timeout=timeout)
            if response.status_code==200:
                data = response.json()
            print("response:",response)
            # print("headers:",session.headers)
            # print(response.json())
            # with open(f"{requestID}.json","w") as f:
            #     #json.dump(response.json(),f,indent=4)
        except Exception as e:
            print(f"Exception in get_report: {e} ")
        
        return data

# Note: json_data will not be serialized by requests
# exactly as it was in the original request.
#data = '{"startDate":"2024-06-14T00:00:00","endDate":"2024-06-14T23:59:59","shifts":[],"salesRoles":[],"terminals":[],"format":{"id":2121400001,"name":"Site Financial Detail & Chem Report-2021","title":"1","indentOffset":2,"indentSpaces":3},"employees":[],"showEachShift":false,"showEachTerminal":false,"showEachSite":false,"paperSize":"letter","paginationOffset":null}'
#response = requests.post('https://sitewatch.cloud/api/gsreport/gsreport', params=params, cookies=cookies, headers=headers, data=data)
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
    req_id = client.get_general_sales_report_request_id(reportOn=reportOn)
    
    report_data = client.get_report(reportOn,req_id)
    # print(report_data)

    try:
        gsviews = report_data.get("gsviews")
        gsviews_0 = gsviews[0]
        section_1 = gsviews_0.get("sections")[1]
        reports_0 = section_1.get("reports")[0]

        saleItemCount = reports_0.get("saleItemCount")
        print(saleItemCount)
    except Exception as e:
        print(f"Exception as reporting data {e}")
        

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
