from washify import washifyClient
import os 
import json
import sys
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

# Add the carwash directory to the sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

username = 'Cameron'
password = 'Password1'
companyName = 'cleangetawayexpress'
userType = 'CWA'

file_path="washify_test.xlsx"

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


def write_dictionary_to_xlshet(dictionary_data,file_name,is_first_dictionary=True):
    for index,data in enumerate(dictionary_data):
                if index == 0:
                    column_gap = 0 if is_first_dictionary else 2
                    append_dict_to_excel(file_name,data,column_gap)
                else:
                    append_dict_to_excel(file_name,data,0,False)

def generate_weekly_report(file_path):
    "This will generate weekly report"
    
    try:
        client  = washifyClient()
        is_logged_in = client.check_login()

        if not is_logged_in:
            login = client.login(username=username,password=password,companyName=companyName,userType=userType)
            print(f"doing relogin : {login}")
        client_locations = client.get_user_locations()
        
        client_locations_number_codes =list(client_locations.values())
        
        if client_locations_number_codes:
            wash_packages_response = client.GetRevenuReportFinancialWashPackage(client_locations_number_codes)
            wash_packages_data = client.GetRevenuReportFinancialWashPackage_formatter(wash_packages_response)  #first table 
            
            write_dictionary_to_xlshet(wash_packages_data,file_path) #first dictionary
            
            wash_package_discount_response = client.GetRevenuReportFinancialWashDiscounts(client_locations_number_codes)
            washpack_discount_data = client.GetRevenuReportFinancialWashDiscounts_formatter(wash_package_discount_response)
            
            write_dictionary_to_xlshet(washpack_discount_data,file_path,False)
            
            wash_extra_response = client.GetRevenuReportFinancialPackagesDiscount(client_locations_number_codes)
            wash_extra_data = client.GetRevenuReportFinancialPackagesDiscount_formatter(wash_extra_response)
            
            write_dictionary_to_xlshet(wash_extra_data,file_path,False)
            
            unlimited_sales_response = client.GetRevenuReportFinancialUnlimitedSales(client_locations_number_codes)
            unlimited_sales_data  = client.GetRevenuReportFinancialUnlimitedSales_formatter(unlimited_sales_response) 
            
            write_dictionary_to_xlshet(unlimited_sales_data,file_path,False)
            
            giftcard_sales_response = client.GetRevenuReportFinancialGiftcardsale(client_locations_number_codes)
            giftcards_sales_data = client.GetRevenuReportFinancialGiftcardsale_formatter(giftcard_sales_response) 
            
            write_dictionary_to_xlshet(giftcards_sales_data,file_path,False)

            discount_discount_response = client.GetRevenuReportFinancialWashDiscounts(client_locations_number_codes)
            discount_discount_data = client.GetRevenuReportFinancialWashDiscounts_formatter2(discount_discount_response) 
            
            write_dictionary_to_xlshet(discount_discount_data,file_path,False)       
            
            giftcard_reedemed_response = client.GetRevenuReportFinancialRevenueSummary(client_locations_number_codes)
            giftcard_reedemed_data = client.GetRevenuReportFinancialRevenueSummary_formatted(giftcard_reedemed_response)
            
            write_dictionary_to_xlshet(giftcard_reedemed_data,file_path,False)
            
            payment_response  = client.GetRevenuReportFinancialPaymentNew(client_locations_number_codes)
            payment_data = client.GetRevenuReportFinancialPaymentNew_formatter(payment_response) 
            
            write_dictionary_to_xlshet(payment_data,file_path,False)
        
    except Exception as e:
        print(f"Exception generate_weeklyrepoer washify {e}")
    


if __name__=="__main__":
    generate_weekly_report(file_path)