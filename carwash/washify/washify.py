import json
import os 
import datetime
import requests

current_file_path = os.path.dirname(os.path.abspath(__file__))
# print(current_file_path)

cookies_path = os.path.join(current_file_path,"cookies")

cookie_file_path = os.path.join(cookies_path,"cookie.json")
print(cookies_path)

proxy_url=None

proxy = {
    "http":proxy_url,
    "https":proxy_url
}




class washifyClient():
    def __init__(self,) -> None:
        pass

    def login(self,username,password,companyName,userType,proxy)->bool:
        """login fucntion

        Args:
            username (str):username
            password (str): password
            companyName (str): companyname
            userType (str): usertype
        """
        headers = {
                'accept': 'application/json, text/plain, */*',
                'accept-language': 'en-US,en;q=0.9',
                'content-type': 'application/json',
                'origin': 'https://washifyapi.com:1000',
                'priority': 'u=1, i',
                'referer': 'https://washifyapi.com:1000/',
                'sec-ch-ua': '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
                'sec-ch-ua-mobile': '?0',
                'sec-ch-ua-platform': '"Linux"',
                'sec-fetch-dest': 'empty',
                'sec-fetch-mode': 'cors',
                'sec-fetch-site': 'same-site',
                'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            }
        
        # json_data = {
        #         'username': 'Cameron',
        #         'password': 'Password1',
        #         'companyName': 'cleangetawayexpress',
        #         'userType': 'CWA',
        #     }
        
        json_data = {
                'username': username,
                'password': password,
                'companyName': companyName,
                'userType': userType,
            }
        
        try:
            response = requests.post('https://washifyapi.com:8298/api/AccountLogin/ValidateUserCredentials', headers=headers, json=json_data,proxies=proxy)
            if response.status_code==200:
                data = response.json()
                print(len(data['data']))
                mainCompanyID = data['data'].get("mainCompanyID")
                companyID = data['data'].get("companyID")
                serverID = data['data'].get("serverID")
                userRoleID = data['data'].get("userRoleID")
                userID= data['data'].get("userID")
                authToken= data['data'].get("authToken")
                userLocations= data['data'].get("userLocations")
                timeOffset= data['data'].get("timeOffset")
                # Store data in a dictionary
                extracted_data = {
                    "mainCompanyID": mainCompanyID,
                    "companyID": companyID,
                    "serverID": serverID,
                    "userRoleID": userRoleID,
                    "userID":userID,
                    "authToken":authToken,
                    "userLocations":userLocations,
                    "timeOffset":timeOffset

                }

                try:
                    with open(cookie_file_path, 'w') as json_file:
                        json.dump(extracted_data, json_file, indent=4)
                    print("Data saved to", cookie_file_path)
                    return True
                except Exception as e:
                    print("Failed to save data to JSON file. Error:", e)
        except Exception as e:
            print(f"Exception in login() {e}")
        return False

    def get_common_data(self):
        common_data={}

        try:
            with open(cookie_file_path,'r') as f:
                cookie_data = json.load(f)
            companyID = cookie_data.get("companyID")
            serverID = cookie_data.get("serverID")
            userRoleID = cookie_data.get("userRoleID")
            userID     = cookie_data.get("userID")
            authToken   = cookie_data.get("authToken")
            timeOffset  = cookie_data.get("timeOffset")

            if all([companyID,serverID,userRoleID,userID,authToken,timeOffset]):
                common_data['commonCompanyID'] = companyID
                common_data['commonServerID'] = serverID
                common_data['commonUserRoleID'] = userRoleID
                common_data['commonUserID'] = userID
                common_data['authKey'] = authToken
                common_data['commonTimeoffset'] =timeOffset
        except (FileExistsError,FileNotFoundError):
            print("Cookeis file not found !")
        
        except Exception as e:
            print(f"Exception : {e}")
        
        return common_data

    def check_login(self,proxy=proxy)->bool:
        "cheks is we can use previous login data or not"

        login_passed = False

        headers = {
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'en-US,en;q=0.9',
            'content-type': 'application/json',
            'dnt': '1',
            'origin': 'https://washifyapi.com:1000',
            'priority': 'u=1, i',
            'referer': 'https://washifyapi.com:1000/',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-site',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        }

        # with open(cookie_file_path,'r') as f:
        #     cookie_data = json.load(f)
        # companyID = cookie_data.get("companyID")
        # serverID = cookie_data.get("serverID")
        # userRoleID = cookie_data.get("userRoleID")
        # userID     = cookie_data.get("userID")
        # authToken   = cookie_data.get("authToken")
        # timeOffset  = cookie_data.get("timeOffset")

        json_data = {
            'CompanyID': 0,
            'UserLocations': '',
            'ServerID': 0,
            'UserRoleID': 0,
            'ID': 0,
            'CommonCompanySettings': self.get_common_data(),
        }
        try:
            response = requests.post('https://washifyapi.com:8298/api/UserRoles/GetRoleId', headers=headers, json=json_data,proxies=proxy)
            if response.status_code==200:
                msg = response.json().get("message")
                login_passed =True if msg=="Success" else False
        except Exception as e:
            print(f"Exception in check_login() : {e}")
        
        return login_passed

    def get_user_locations(self):
        data ={}

        headers = {
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'en-US,en;q=0.9',
            'content-type': 'application/json',
            'dnt': '1',
            'origin': 'https://washifyapi.com:1000',
            'priority': 'u=1, i',
            'referer': 'https://washifyapi.com:1000/',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-site',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        }

        json_data = {
            'companyID': 0,
            'ServerID': 0,
            'ID': 0,
            'CommonCompanySettings': self.get_common_data(),
        }
        try:
            response = requests.post('https://washifyapi.com:8298/api/CommonMethod/getUserLocations', headers=headers, json=json_data)
            if response.status_code==200:
                data = response.json().get("data")
                data = { location.get("locationName").split("-")[-1].strip() :location.get("locationID") for location in data if location.get("locationID")!=0}
        except Exception as e:
            print(f"Error in get_user_locations() : {e}")
        
        return data

    def get_car_count_report(self,location:list):
        data = None
        headers = {
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'en-US,en;q=0.9',
            'content-type': 'application/json',
            'dnt': '1',
            'origin': 'https://washifyapi.com:1000',
            'priority': 'u=1, i',
            'referer': 'https://washifyapi.com:1000/',
            'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-site',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        }
        # Get the current time
        current_time = datetime.datetime.now()
        # Define the desired date format
        date_format_string = '%m/%d/%Y'
        current_date_formatted = current_time.strftime(date_format_string)
        json_data = {
            'Locations': location,
            'StartDate': current_date_formatted,#'06/14/2024',
            'EndDate': current_date_formatted,#'06/14/2024',
            'ReportBy': 'Day',
            'GroupAll': True,
            'CommonCompanySettings': self.get_common_data(),
        }
        try:
            response = requests.post('https://washifyapi.com:8298/api/Reports/GetCarCountReport', headers=headers, json=json_data)
            if response.status_code==200:
                data = response.json().get("data")[0].get("carwashed")
        except Exception as e:
            print(f"Exception : {e}")

        return data
if __name__=="__main__":
    
# json_data = {
        #         'username': 'Cameron',
        #         'password': 'Password1',
        #         'companyName': 'cleangetawayexpress',
        #         'userType': 'CWA',
        #     }
    username = 'Cameron'
    password = 'Password1'
    companyName = 'cleangetawayexpress'
    userType = 'CWA'
    client  = washifyClient()
    # login = client.login(username=username,password=password,
    #                     companyName=companyName,userType=userType,proxy=proxy)
    # print(f"login check : {login}")
    is_logged_in = client.check_login(proxy=proxy)
    client_locations = client.get_user_locations()
    print(f"is logged in :{is_logged_in}")
    print(f"user lcoations : {client_locations}")
    if client_locations:
        for location_name,location_id in client_locations.items():
            print(location_name,":",client.get_car_count_report([location_id,]))



            
        
        