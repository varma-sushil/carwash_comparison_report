import json
import os 
import datetime
import requests
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

file_path="washift.xlsx"

def append_dict_to_excel(file_path, data, num_lines,add_headers=True):
    try:
        # Try to load an existing workbook
        workbook = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        # If the file does not exist, create a new workbook
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Sheet1'
    else:
        # If the file exists, get the active sheet
        sheet = workbook.active

    # Append blank lines
    for _ in range(num_lines):
        sheet.append([])

    if add_headers:
        # Append the header row with bold keys
        bold_font = Font(bold=True)
        header_row = list(data.keys())
        sheet.append(header_row)
        for cell in sheet[sheet.max_row]:
            cell.font = bold_font

    # Append the data row
    data_row = list(data.values())
    sheet.append(data_row)

    # Save the workbook
    workbook.save(file_path)


current_file_path = os.path.dirname(os.path.abspath(__file__))
# print(current_file_path)

cookies_path = os.path.join(current_file_path,"cookies")

cookie_file_path = os.path.join(cookies_path,"cookie.json")

data_path = os.path.join(current_file_path,"data")
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
    
    def get_financal_revenue_summary(self):
        data =None
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
            'Locations': [
                88,
                89,
                87,
                90,
            ],
            'StartDate': '06/10/2024 12:00 AM',
            'EndDate': '06/21/2024 11:59 PM',
            'LogOutDate': '06/21/2024 11:59 PM',
            'locationName': '',
            'ReportBy': '',
            'CommonCompanySettings': self.get_common_data(),
        }

        try:
            response = requests.post(
            'https://washifyapi.com:8298/api/Reports/GetRevenuReportFinancialRevenueSummary',
            headers=headers,
            json=json_data,
        )
            if response.status_code==200:
                data = response.json()
        
        except Exception as e:
            print(f"Error in get_financal_revenue_summary() {e}")

        return data
 
 
    
    def GetRevenuReportFinancialWashPackage(self,client_locations:list):
        "WASH PACKAGES"
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

        json_data = {
            'Locations': client_locations,
            'StartDate': '06/10/2024 12:00 AM',
            'EndDate': '06/21/2024 11:59 PM',
            'LogOutDate': '06/21/2024 11:59 PM',
            'locationName': '',
            'ReportBy': '',
            'CommonCompanySettings': self.get_common_data(),
        }
        try:
            response = requests.post(
                'https://washifyapi.com:8298/api/Reports/GetRevenuReportFinancialWashPackage',
                headers=headers,
                json=json_data,
            )
            if response.status_code==200:
                data = response.json()
        except Exception as e:
            print(f" Exception in GetRevenuReportFinancialWashPackage()  {e}")
            
        return data

    def GetRevenuReportFinancialWashPackage_formatter(self,data):
        "WASH PACKAGES formatter"
        wash_packages_all = []
        Wash_Packages_Unlimited_total = 0
        Wash_Packages_Virtual_Wash_total = 0
        Wash_Packages_Non_Unlimited_total = 0
        Wash_Packages_Total_total  = 0
        Wash_Packages_Amount_total = 0
        Wash_Packages_Total_Amount_total = 0
        try:
            data = data.get("data")
            financialWashPackage = data.get("financialWashPackage")
            for wash_package in financialWashPackage:
                wash_package_structure={}
                
                cUnlimited  = wash_package.get("cUnlimited")
                Wash_Packages_Unlimited_total+=cUnlimited
                
                virtualWashNumber  = wash_package.get("virtualWashNumber")
                Wash_Packages_Virtual_Wash_total+=virtualWashNumber
                
                nonUnlimited    =   wash_package.get("nonUnlimited")
                Wash_Packages_Non_Unlimited_total+=nonUnlimited
                
                total    =   wash_package.get("total")
                Wash_Packages_Total_total += total
                
                price  =  wash_package.get("price")
                Wash_Packages_Amount_total +=price
                
                amount   =  wash_package.get("amount")
                Wash_Packages_Total_Amount_total +=amount
                
                wash_package_structure['Wash_Packages_ServiceName'] = wash_package.get("serviceName")
                wash_package_structure["Wash_Packages_Unlimited"]   = cUnlimited
                wash_package_structure["Wash_Packages_Virtual_Wash"] = virtualWashNumber
                wash_package_structure["Wash_Packages_Non_Unlimited"]  = nonUnlimited
                wash_package_structure["Wash_Packages_Total"]   =   total
                wash_package_structure["Wash_Packages_Amount"]   =    price
                wash_package_structure["Wash_Packages_Total_Amount"]  =amount
                # print(type(wash_package.get("amount")))
                wash_packages_all.append(wash_package_structure)
            
            wash_package_total_structure={
                "Wash_Packages_ServiceName":"Total:",
                "Wash_Packages_Unlimited":Wash_Packages_Unlimited_total,
                "Wash_Packages_Virtual_Wash":Wash_Packages_Virtual_Wash_total,
                "Wash_Packages_Non_Unlimited":Wash_Packages_Non_Unlimited_total,
                "Wash_Packages_Total":Wash_Packages_Total_total,
                "Wash_Packages_Amount":Wash_Packages_Amount_total,
                "Wash_Packages_Total_Amount":Wash_Packages_Total_Amount_total   
            }
            
            wash_packages_all.append(wash_package_total_structure) #last ALl total row
                
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialWashPackage_formatter() {e}")
        return wash_packages_all



    def GetRevenuReportFinancialWashDiscounts(self,client_locations):
        "DISCOUNTS"
        data=None
        


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
            'Locations': client_locations,
            'StartDate': '06/10/2024 12:00 AM',
            'EndDate': '06/21/2024 11:59 PM',
            'LogOutDate': '06/21/2024 11:59 PM',
            'locationName': '',
            'ReportBy': '',
            'CommonCompanySettings': self.get_common_data(),
        }

        try:
            response = requests.post(
                'https://washifyapi.com:8298/api/Reports/GetRevenuReportFinancialWashDiscounts',
                headers=headers,
                json=json_data,
            )

            if response.status_code==200:
                data = response.json()
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialWashDiscounts()  {e}")

        return data
    
    def  GetRevenuReportFinancialWashDiscounts_formatter(self,data):
        "DISCOUNTS formatter"
        discount_all =[]
        Wash_Packages_Discount_Number_total  = 0
        Wash_Packages_Discount_Service_Price_total = 0
        Wash_Packages_Discount_Total_Discount = 0
        
        try:
            data = data.get("data")
            financialWashDiscounts = data.get("financialWashDiscounts")
            
            for wash_discount in financialWashDiscounts:
                wash_discount_structure = {}
                
                number   =  wash_discount.get("number")
                Wash_Packages_Discount_Number_total+=number
                
                discountPrice  = wash_discount.get("discountPrice")
                Wash_Packages_Discount_Service_Price_total+=discountPrice
                
                totalAmt   =  wash_discount.get("totalAmt")
                Wash_Packages_Discount_Total_Discount+=totalAmt
                
                
                wash_discount_structure["Wash_Packages_Discount_ServiceName"] = wash_discount.get("discountName")
                wash_discount_structure["Wash_Packages_Discount_Number"]      = number
                wash_discount_structure["Wash_Packages_Discount_Service Price ($)"] = discountPrice
                wash_discount_structure["Wash_Packages_Discount_Total Discount ($)"] = totalAmt
                
                discount_all.append(wash_discount_structure)
                
            #appending Grand total
            discounts_all_total_structure = {
                "Wash_Packages_Discount_ServiceName":"Total:",
                "Wash_Packages_Discount_Number":Wash_Packages_Discount_Number_total,
                "Wash_Packages_Discount_Service Price ($)":Wash_Packages_Discount_Service_Price_total,
                "Wash_Packages_Discount_Total Discount ($)":Wash_Packages_Discount_Total_Discount
            }
            discount_all.append(discounts_all_total_structure)
            
            
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialWashDiscounts_formatter() {e}")
          
        return discount_all

    def GetRevenuReportFinancialWashDiscounts_formatter2(self,data):
        "DISCOUNTS DISCOUNTS formatter2"
        discount_discount_all =[]
        DISCOUNTS_Number_total = 0
        DISCOUNTS_Price_total = 0
        DISCOUNTS_Revenue_total = 0
        
        try:
            
            data = data.get("data")

            financialWashDiscounts = data.get("financialWashDiscounts")     
            
            for wash_discount in financialWashDiscounts:
                
                number  = int(wash_discount.get("number",0))
                DISCOUNTS_Number_total+=number
                
                discountPrice  = float(wash_discount.get("discountPrice",0.0))
                DISCOUNTS_Price_total+=discountPrice
                
                totalAmt  =  float(wash_discount.get("totalAmt",0.0))
                DISCOUNTS_Revenue_total+=totalAmt
                
                discount_discount_structure = {}                                                             # Discount Discount
                discount_discount_structure["DISCOUNTS_Discount"] = wash_discount.get("discountName")
                discount_discount_structure["DISCOUNTS_Number"]      = number
                discount_discount_structure["DISCOUNTS_Price ($)"] = discountPrice
                discount_discount_structure["DISCOUNTS_Revenue"] = totalAmt
                
                discount_discount_all.append(discount_discount_structure)
                
            #Total table Discount discount
            discount_discount_total={
                "DISCOUNTS_Discount":"Total:",
                "DISCOUNTS_Number":DISCOUNTS_Number_total,
                "DISCOUNTS_Price ($)":DISCOUNTS_Price_total,
                "DISCOUNTS_Revenue":DISCOUNTS_Revenue_total
            }
            
            discount_discount_all.append(discount_discount_total)
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialWashDiscounts_fromatter2 {e}")
            
        return discount_discount_all

    # def GetRevenuReportFinancialWashDiscounts(self):
        
    #     data = None
        
    #     headers = {
    #         'accept': 'application/json, text/plain, */*',
    #         'accept-language': 'en-US,en;q=0.9',
    #         'content-type': 'application/json',
    #         'dnt': '1',
    #         'origin': 'https://washifyapi.com:1000',
    #         'priority': 'u=1, i',
    #         'referer': 'https://washifyapi.com:1000/',
    #         'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
    #         'sec-ch-ua-mobile': '?0',
    #         'sec-ch-ua-platform': '"Windows"',
    #         'sec-fetch-dest': 'empty',
    #         'sec-fetch-mode': 'cors',
    #         'sec-fetch-site': 'same-site',
    #         'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    #     }

    #     json_data = {
    #         'Locations': [
    #             88,
    #             89,
    #             87,
    #             90,
    #         ],
    #         'StartDate': '06/10/2024 12:00 AM',
    #         'EndDate': '06/21/2024 11:59 PM',
    #         'LogOutDate': '06/21/2024 11:59 PM',
    #         'locationName': '',
    #         'ReportBy': '',
    #         'CommonCompanySettings': self.get_common_data(),
    #     }
    #     try:
    #         response = requests.post(
    #             'https://washifyapi.com:8298/api/Reports/GetRevenuReportFinancialWashDiscounts',
    #             headers=headers,
    #             json=json_data,
    #         )
            
    #         if response.status_code==200:
    #             data = response.json()
    #     except Exception as e:
    #         print(f"Exception in GetRevenuReportFinancialWashDiscounts() {e}")

    #     return data
   
 
 
   
    
    def GetRevenuReportFinancialPackagesDiscount(self,client_locations):
        "WASH EXTRAS"
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

        json_data = {
            'Locations': client_locations,
            'StartDate': '06/10/2024 12:00 AM',
            'EndDate': '06/21/2024 11:59 PM',
            'LogOutDate': '06/21/2024 11:59 PM',
            'locationName': '',
            'ReportBy': '',
            'CommonCompanySettings': self.get_common_data(),
        }

        try:
            response = requests.post(
                'https://washifyapi.com:8298/api/Reports/GetRevenuReportFinancialPackagesDiscount',
                headers=headers,
                json=json_data,
            )
            if response.status_code==200:
                data = response.json()
        except Exception as e:
            print(f"Exception on GetRevenuReportFinancialPackagesDiscount() {e}")
            
        return data

    def GetRevenuReportFinancialPackagesDiscount_formatter(self,data):
        "WASH EXTRAS formatter"
        wash_extras_all = []
        Wash_Extras_Number_total=0
        Wash_Extras_Amount_total = 0
        Wash_Extras_Total_Amount_total = 0
        
        try:
            data = data.get("data")
            financialPackagesDiscount = data.get("financialPackagesDiscount") 
            
            for wash_extra in financialPackagesDiscount:
                wash_extra_structure={}
                
                number = int(wash_extra.get("number",0))
                # print("number:",type(number))
                Wash_Extras_Number_total+=number
                
                servicePrice = float(wash_extra.get("servicePrice",0.0))
                Wash_Extras_Amount_total+=servicePrice
                
                totalAmount = float(wash_extra.get("totalAmount",0.0))
                Wash_Extras_Total_Amount_total+=totalAmount
                
                wash_extra_structure["Wash_Extras_ServiceName"] = wash_extra.get("serviceName")
                wash_extra_structure["Wash_Extras_Number"]      = number
                wash_extra_structure["Wash_Extras_Amount ($)"]  = servicePrice
                wash_extra_structure["Wash_Extras_Total_Amount ($)"] = totalAmount
                
                wash_extras_all.append(wash_extra_structure)
            #adding final total value 
            wash_extras_total_structure = {
                "Wash_Extras_ServiceName":"Total:",
                "Wash_Extras_Number":Wash_Extras_Number_total,
                "Wash_Extras_Amount ($)":Wash_Extras_Amount_total,
                "Wash_Extras_Total_Amount ($)":Wash_Extras_Total_Amount_total
            }   
            
            wash_extras_all.append(wash_extras_total_structure)
                
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialPackagesDiscount_formatter()  {e}")

        return wash_extras_all





    def GetRevenuReportFinancialUnlimitedSales(self,client_locations):
        "UNLIMITED SALES"
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

        json_data = {
            'Locations': client_locations,
            'StartDate': '06/10/2024 12:00 AM',
            'EndDate': '06/21/2024 11:59 PM',
            'LogOutDate': '06/21/2024 11:59 PM',
            'locationName': '',
            'ReportBy': '',
            'CommonCompanySettings':self.get_common_data(),
        }

        try:
            response = requests.post(
                'https://washifyapi.com:8298/api/Reports/GetRevenuReportFinancialUnlimitedSales',
                headers=headers,
                json=json_data,
            )
            if response.status_code==200:
                data = response.json()
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialUnlimitedSales() {e}")

        return data

    def GetRevenuReportFinancialUnlimitedSales_formatter(self,data):
        "UNLIMITED SALES formatter"
        
        unlimited_sales_all =[]
        
        Unlimited_Sales_Number_total = 0
        Unlimited_Sales_Revenue_total =0
        try:
            data = data.get("data")
            financialUnlimitedSales  = data.get("financialUnlimitedSales")
            
            for sales_data in financialUnlimitedSales:
                sales_data_structure ={}
                # print(sales_data)
                number  = int(sales_data.get("number",0))
                Unlimited_Sales_Number_total+=number
                
                price  =  float(sales_data.get("price",0.0))
                Unlimited_Sales_Revenue_total+=price
                
                sales_data_structure["Unlimited_Sales"] = sales_data.get("unlimited_Sales")
                sales_data_structure["Unlimited_Sales_Service"] = sales_data.get("serviceName")
                sales_data_structure["Unlimited_Sales_Number"]  = number
                sales_data_structure["Unlimited_Sales_Revenue ($)"] = price 
                
                unlimited_sales_all.append(sales_data_structure)
              
            #final table total
            unlimited_sales__total = {
                "Unlimited_Sales":"Total:",
                "Unlimited_Sales_Service":"",
                "Unlimited_Sales_Number":Unlimited_Sales_Number_total,
                "Unlimited_Sales_Revenue ($)":Unlimited_Sales_Revenue_total
            }    
            
            unlimited_sales_all.append(unlimited_sales__total)
            
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialUnlimitedSales_formatter() {e}")
            
        return unlimited_sales_all

 
 
 

    def GetRevenuReportFinancialGiftcardsale(self,client_locations):
        "GIFT CARD SALES"
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

        json_data = {
            'Locations': client_locations,
            'StartDate': '06/10/2024 12:00 AM',
            'EndDate': '06/21/2024 11:59 PM',
            'LogOutDate': '06/21/2024 11:59 PM',
            'locationName': '',
            'ReportBy': '',
            'CommonCompanySettings': self.get_common_data(),
        }

        try:
            response = requests.post(
                'https://washifyapi.com:8298/api/Reports/GetRevenuReportFinancialGiftcardsale',
                headers=headers,
                json=json_data,
            )
            
            if response.status_code==200:
                data = response.json()
        except Exception as e:
            print("Exception in GetRevenuReportFinancialGiftcardsale() {e}")
            
        return data

    def GetRevenuReportFinancialGiftcardsale_formatter(self,data):
        "GIFT CARD SALES formatter "
        
        gift_card_sale_all =[]
        GIFT_CARD_SALESr_Amount_total = 0
        try:
            data = data.get("data")

            financialGiftcardsale = data.get("financialGiftcardsale")
            
            for gift_card in financialGiftcardsale:
                gift_card_sale_structure = {}
                
                price   =  float(gift_card.get("price",0.0))
                GIFT_CARD_SALESr_Amount_total+=price
                
                gift_card_sale_structure["GIFT_CARD_SALES_DATE"] = gift_card.get("date")  #giftcarsd sales
                gift_card_sale_structure["GIFT_CARD_SALES_TIME"] = gift_card.get("time")
                gift_card_sale_structure["GIFT_CARD_SALES_Card_Number"] = gift_card.get("coupanNumber")
                gift_card_sale_structure["GIFT_CARD_SALESr_Amount ($)"] = gift_card.get("price")
                gift_card_sale_structure["GIFT_CARD_SALES_Source"]      = gift_card.get("transactionFrom")
                
                gift_card_sale_all.append(gift_card_sale_structure)
                
            #final total table 
            giftcard_sale_total = {
                "GIFT_CARD_SALES_DATE":"Total:",
                "GIFT_CARD_SALES_TIME":"",
                "GIFT_CARD_SALES_Card_Number":"",
                "GIFT_CARD_SALESr_Amount ($)":GIFT_CARD_SALESr_Amount_total,
                "GIFT_CARD_SALES_Source":""
            }
            gift_card_sale_all.append(giftcard_sale_total)
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialGiftcardsale_formatter() {e}")
            
        return gift_card_sale_all



    # def GetRevenuReportFinancialWashDiscounts(self):
    #     "DISCOUNTS"
    #     data = None
        

    #     headers = {
    #         'accept': 'application/json, text/plain, */*',
    #         'accept-language': 'en-US,en;q=0.9',
    #         'content-type': 'application/json',
    #         'dnt': '1',
    #         'origin': 'https://washifyapi.com:1000',
    #         'priority': 'u=1, i',
    #         'referer': 'https://washifyapi.com:1000/',
    #         'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
    #         'sec-ch-ua-mobile': '?0',
    #         'sec-ch-ua-platform': '"Windows"',
    #         'sec-fetch-dest': 'empty',
    #         'sec-fetch-mode': 'cors',
    #         'sec-fetch-site': 'same-site',
    #         'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    #     }

    #     json_data = {
    #         'Locations': [
    #             88,
    #             89,
    #             87,
    #             90,
    #         ],
    #         'StartDate': '06/10/2024 12:00 AM',
    #         'EndDate': '06/21/2024 11:59 PM',
    #         'LogOutDate': '06/21/2024 11:59 PM',
    #         'locationName': '',
    #         'ReportBy': '',
    #         'CommonCompanySettings': {
    #             'commonCompanyID': 'WWhtk4RrRQ8KymvwT0BMaw==',
    #             'commonServerID': '+iMYUwYx079az+3TrcOsag==',
    #             'commonUserRoleID': 'Hud67uIVSx9QDTsSrUCfjg==',
    #             'commonUserID': 'bKN5Lonb9871/yWYGAEuAQ==',
    #             'authKey': 'jF9inrLiZCOgWlXjXi0Z13m4qUvjpsZcQujn6kXp6iE=',
    #             'commonTimeoffset': '120',
    #         },
    #     }

    #     try:
    #         response = requests.post(
    #             'https://washifyapi.com:8298/api/Reports/GetRevenuReportFinancialWashDiscounts',
    #             headers=headers,
    #             json=json_data,
    #         )
    #         if response.status_code==200:
    #             data = response.json()
    #     except Exception as e:
    #         print(f"Exception in GetRevenuReportFinancialWashDiscounts() {e}")
            
    #     return data

    def GetRevenuReportFinancialRevenueSummary(self,client_locations):
        "GIFT CARD REDEEMED"
        
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

        json_data = {
            'Locations': client_locations,
            'StartDate': '06/10/2024 12:00 AM',
            'EndDate': '06/21/2024 11:59 PM',
            'LogOutDate': '06/21/2024 11:59 PM',
            'locationName': '',
            'ReportBy': '',
            'CommonCompanySettings': self.get_common_data(),
        }
        try:
            response = requests.post(
                'https://washifyapi.com:8298/api/Reports/GetRevenuReportFinancialRevenueSummary',
                headers=headers,
                json=json_data,
            )
            
            if response.status_code==200:
                data = response.json()
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialRevenueSummary() {e}")

        return data
    
    def GetRevenuReportFinancialRevenueSummary_formatted(self,data):
        "GIFT CARD REDEEMED formatter"
        
        reedemed_giftcard_all =[]
        GIFT_CARD_REDEEMED_Amount_total = 0
        try:
            data = data.get("data")

            financialGiftcardRedeemed = data.get("financialGiftcardRedeemed")
            
            
            for reedemed_giftcard in financialGiftcardRedeemed:
                reedemed_giftcard_structure = {}
                # print(reedemed_giftcard)
                
                price = float(reedemed_giftcard.get("price",0.0))
                GIFT_CARD_REDEEMED_Amount_total+=price
                
                reedemed_giftcard_structure["GIFT_CARD_REDEEMED_DATE"] = reedemed_giftcard.get("date")
                reedemed_giftcard_structure["GIFT_CARD_REDEEMED_TIME"] = reedemed_giftcard.get("time")
                reedemed_giftcard_structure["GIFT_CARD_REDEEMED_CARD_NUMBER"] = reedemed_giftcard.get("coupanNumber")
                reedemed_giftcard_structure["GIFT_CARD_REDEEMED_Amount ($)"]  = price
                
                reedemed_giftcard_all.append(reedemed_giftcard_structure)
            #reedem giftcard total
            
            reedem_total_structure ={
                "GIFT_CARD_REDEEMED_DATE":"Total:",
                "GIFT_CARD_REDEEMED_TIME":"",
                "GIFT_CARD_REDEEMED_CARD_NUMBER":"",
                "GIFT_CARD_REDEEMED_Amount ($)":GIFT_CARD_REDEEMED_Amount_total
            }    
            reedemed_giftcard_all.append(reedem_total_structure)
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialRevenueSummary_formatter()  {e}")
            
        return reedemed_giftcard_all





    def GetRevenuReportFinancialPaymentNew(self,client_locations):
        "Payment"
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

        json_data = {
            'Locations': client_locations,
            'StartDate': '06/10/2024 12:00 AM',
            'EndDate': '06/21/2024 11:59 PM',
            'LogOutDate': '06/21/2024 11:59 PM',
            'locationName': '',
            'ReportBy': '',
            'CommonCompanySettings': self.get_common_data(),
        }

        try:
            response = requests.post(
                'https://washifyapi.com:8298/api/Reports/GetRevenuReportFinancialPaymentNew',
                headers=headers,
                json=json_data,
            )
            if response.status_code==200:
                data = response.json()
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialPaymentNew() {e}")
            
        return data

    def GetRevenuReportFinancialPaymentNew_formatter(self,data):
        "Payment formatter"
        
        payment_data_all =[]
        Payment_Cash_total = 0
        Payment_Credit_Card_total = 0
        Payment_Check_total = 0
        Payment_Invoice_total = 0
        Payment_ACH_total = 0
        Payment_Total_total = 0
        try:
            data = data.get('data')
            financialPaymentNew = data.get("financialPaymentNew")
            
            for payment in financialPaymentNew:
                payment_structure = {}
                # print(payment)

                cash = float(payment.get("cash",0.0))
                Payment_Cash_total+=cash
                
                creditCard = float(payment.get("creditCard",0.0))
                Payment_Credit_Card_total +=creditCard
                
                checkpayment  = float(payment.get("checkpayment",0.0))
                Payment_Check_total+=checkpayment
                
                invoiceCustomer = float(payment.get("invoiceCustomer",0.0))
                Payment_Invoice_total+=invoiceCustomer
                
                
                ach = float(payment.get("ach",0.0))
                Payment_ACH_total +=ach
                
                total_payment = sum([cash,creditCard,checkpayment,invoiceCustomer,ach])
                Payment_Total_total+=total_payment
                
                payment_structure["Payment_Location"] = payment.get("locationName")
                payment_structure["Payment_Cash"]     = cash
                payment_structure["Payment_Credit_Card"]  = creditCard
                payment_structure["Payment_Check"]     = checkpayment
                payment_structure["Payment_Invoice"]   = invoiceCustomer
                payment_structure["Payment_ACH"]       = ach
                payment_structure["Payment_Total ($)"] = total_payment  ##payment
                
                payment_data_all.append(payment_structure)
                  
            #total payments row 
            payment_total_structure = {
                "Payment_Location":"Total Payments:",
                "Payment_Cash":Payment_Cash_total,
                "Payment_Credit_Card":Payment_Credit_Card_total,
                "Payment_Check":Payment_Check_total,
                "Payment_Invoice":Payment_Invoice_total,
                "Payment_ACH":Payment_ACH_total,
                "Payment_Total ($)":Payment_Total_total
            }
            
            payment_data_all.append(payment_total_structure)
                  
        except Exception as e:
            print(f"Exception in GetRevenuReportFinancialPaymentNew_formatter()  {e}")


        return payment_data_all

    
    
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
        # for location_name,location_id in client_locations.items():
        #     print(location_name,":",client.get_car_count_report([location_id,]))
        print(client_locations)
        
        # response_reneue = client.GetRevenuReportFinancialWashPackage()
        # with open(f"{data_path}\GetRevenuReportFinancialWashPackage.json","w") as f:  #for wah package
        #     json.dump(response_reneue,f,indent=4)
        
        # response_reneue = client.GetRevenuReportFinancialWashDiscounts()
        # with open(f"{data_path}\GetRevenuReportFinancialWashDiscounts.json","w") as f:  #for Discount
        #     json.dump(response_reneue,f,indent=4)
            
        # response_reneue = client.GetRevenuReportFinancialPackagesDiscount()
        # with open(f"{data_path}\GetRevenuReportFinancialPackagesDiscount.json","w") as f:  #for Discount
        #     json.dump(response_reneue,f,indent=4)
        # response_reneue = client.GetRevenuReportFinancialUnlimitedSales()
        # with open(f"{data_path}\GetRevenuReportFinancialUnlimitedSales.json","w") as f:  #for Unlimited sales
        #     json.dump(response_reneue,f,indent=4)
            
        # response_reneue = client.GetRevenuReportFinancialGiftcardsale()
        # with open(f"{data_path}\GetRevenuReportFinancialGiftcardsale.json","w") as f:  #for Gift card sale
        #     json.dump(response_reneue,f,indent=4)
        # response_reneue = client.GetRevenuReportFinancialRevenueSummary()
        # with open(f"{data_path}\GetRevenuReportFinancialRevenueSummary.json","w") as f:  #for Unlimited sales
        #     json.dump(response_reneue,f,indent=4)
        
        # response_reneue = client.GetRevenuReportFinancialPaymentNew()
        # with open(f"{data_path}\GetRevenuReportFinancialPaymentNew.json","w") as f:  #for Payment
        #     json.dump(response_reneue,f,indent=4)
            
            
        ## Formatter logic
        wash_packages_response = client.GetRevenuReportFinancialWashPackage()
        wash_packages_data = client.GetRevenuReportFinancialWashPackage_formatter(wash_packages_response)  #first table 
        
        for index,data in enumerate(wash_packages_data):
            if index == 0:
                append_dict_to_excel(file_path,data,0)
            else:
                append_dict_to_excel(file_path,data,0,False)
            # print(data)
        # print(json.dumps(formatted_response,indent=4))

        wash_package_discount_response = client.GetRevenuReportFinancialWashDiscounts()
        washpack_discount_data = client.GetRevenuReportFinancialWashDiscounts_formatter(wash_package_discount_response)  #secound table 
        # print(json.dumps(formatted_response,indent=4))
        
        for index,data in enumerate(washpack_discount_data):
            if index == 0:
                append_dict_to_excel(file_path,data,2)
            else:
                append_dict_to_excel(file_path,data,0,False)

        wash_extra_response = client.GetRevenuReportFinancialPackagesDiscount()
        wash_extra_data = client.GetRevenuReportFinancialPackagesDiscount_formatter(wash_extra_response)  #3rd table 
        # print(json.dumps(formatted_response,indent=4))
        for index,data in enumerate(wash_extra_data):
            if index == 0:
                append_dict_to_excel(file_path,data,2)
            else:
                append_dict_to_excel(file_path,data,0,False)

        unlimited_sales_response = client.GetRevenuReportFinancialUnlimitedSales()
        unlimited_sales_data  = client.GetRevenuReportFinancialUnlimitedSales_formatter(unlimited_sales_response) #unlimited sales
        # print(unlimited_sales_data)
        # print(json.dumps(formatted_response,indent=4))
        for index,data in enumerate( unlimited_sales_data):
            if index == 0:
                append_dict_to_excel(file_path,data,2)
            else:
                append_dict_to_excel(file_path,data,0,False)
            
        giftcard_sales_response = client.GetRevenuReportFinancialGiftcardsale()
        giftcards_sales_data = client.GetRevenuReportFinancialGiftcardsale_formatter(giftcard_sales_response)  #4rd table  gift card sale
        # print(formatted_response) 
        # print(json.dumps(formatted_response,indent=4))
        
        for index,data in enumerate(giftcards_sales_data):
            if index == 0:
                append_dict_to_excel(file_path,data,2)
            else:
                append_dict_to_excel(file_path,data,0,False)
        
        
        discount_discount_response = client.GetRevenuReportFinancialWashDiscounts()
        discount_discount_data = client.GetRevenuReportFinancialWashDiscounts_formatter2(discount_discount_response)  #5rd table    Discount discount
        # print(json.dumps(formatted_response,indent=4))
        
        for index,data in enumerate(discount_discount_data):
            if index == 0:
                append_dict_to_excel(file_path,data,2)
            else:
                append_dict_to_excel(file_path,data,0,False)
        
        
        giftcard_reedemed_response = client.GetRevenuReportFinancialRevenueSummary()
        giftcard_reedemed_data = client.GetRevenuReportFinancialRevenueSummary_formatted(giftcard_reedemed_response)  #6rd table   Discount discount
        # print(json.dumps(formatted_response,indent=4))
        for index,data in enumerate(giftcard_reedemed_data):
            if index == 0:
                append_dict_to_excel(file_path,data,2)
            else:
                append_dict_to_excel(file_path,data,0,False)
        
        
        
        # response = client.GetRevenuReportFinancialRevenueSummary()   #duplicate
        # formatted_response = client.GetRevenuReportFinancialRevenueSummary_formatted(response)  #8rd table  gift card reedem
        # print(json.dumps(formatted_response,indent=4))
        
        
        payment_response  = client.GetRevenuReportFinancialPaymentNew()
        payment_data = client.GetRevenuReportFinancialPaymentNew_formatter(payment_response)  #8rd table payment location
        # print(json.dumps(payment_data,indent=4))
        for index,data in enumerate(payment_data):
            if index == 0:
                append_dict_to_excel(file_path,data,2)
            else:
                append_dict_to_excel(file_path,data,0,False)
        
        
        
        
        
## need to chanege site locations dynamic and and time stamp also dynamic and need to use type casting and need to write xl conversion code