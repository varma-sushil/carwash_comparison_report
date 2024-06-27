from datetime import datetime, timedelta
import os
import json

import pandas as pd 
from datetime import datetime, timedelta
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

from sitewatch4 import sitewatchClient
import sys
# Add the carwash directory to the sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))



tg_message=[]

current_folder_path = os.path.dirname(os.path.abspath(__file__))

cookies_path = os.path.join(current_folder_path,"cookies")

xlfile_path = os.path.join(current_folder_path,"sitewash_data.xlsx")

site_watch_latest_json = os.path.join(current_folder_path,"sitewatch_data_latest.json")

differenec_json = os.path.join(current_folder_path,"difference.json")

sites_df = pd.read_excel(xlfile_path)

orginations_lst = ['SPKLUS-001', 'SPKLUS-002', 
                   'SPKLUS-003', 'SPKLUS-004', 
                   'SPKLUS-005', 'SPKLUS-006', 
                   'SPKLUS-007', 'SPKLUS-008', 
                   'SPKLUS-009', 'SPKLUS-012', 
                   'SPKLUS-013', 'SPKLUS-014', 
                   'SPKLUS-015', 'SUDZWL-002']

client_names = ['Belair', 'Evans', 'North Augusta', 
                'Greenwood', 'Grovetown 1', 'Windsor Springs',
                'Furys Ferry (Martinez)', 'Peach Orchard Rd.', 
                'Grovetown 2', 'Cicero', 'Matteson', 'Sparkle Express ', 
                "Fuller's Calumet City "]

# wash_sales_lst= []
# wash_packages_lst =[]
# wash_extra_service_lst =[]
# gross_wash_sales_lst = []
# less_wash_sales_rdmd_lst =[]
# less_wash_discounts_lst=[]
# less_loyality_disc_lst = []

# net_site_sales_lst = []

# arm_plans_sold_lst  = []

# arm_plans_recharged_lst = []

# arm_plans_reedemed_lst = []

# arm_plans_terminated_lst = []

# prepaid_sold_lst = []

# less_prepaid_reedemed_lst = []

# online_sold_lst = []

# less_online_reedemed_lst = []

# free_washes_issued_lst = []

# less_paidouts_lst = []

# total_to_account_for_lst =[]

# deposits_lst = []

# total_xpt_cash_lst = []

# house_accounts_lst =[]

# over_short_lst = []

# cash_lst = []

# xpt_acceptors_lst = []

# xpt_dispensers_lst = []

# total_lst = []

# credit_card_list = []

# other_tenders_lst = []

# xpt_balancing_lst = []

# report_balance_lst = []

# picture_mismatch_lst = []

def wash_sales(section):
    wash_sales_lst = []
    
    # reports = section.get("reports")
    # for report in reports:
    #     wash_sales_structure ={
    #         "Wash_sales_Description":report.get("description"),
    #         "Wash_sales_price":report.get("price"),
    #         "Wash_sales_quantity":report.get("quantity"),
    #         "Wash_sales_amount":report.get("amount")
    #     }
    #     wash_sales_lst.append(wash_sales_structure)
    
    subtotals = section.get("subtotals")[0]
    
    # for subtotal in subtotals:
    #     wash_sales_structure ={
    #         "Wash_sales_Description":subtotal.get("description"),
    #         "Wash_sales_price":subtotal.get("price"),
    #         "Wash_sales_quantity":subtotal.get("quantity"),
    #         "Wash_sales_amount":subtotal.get("amount")
    #     }
    #     wash_sales_lst.append(wash_sales_structure)
    return subtotals.get("quantity")
        
    return wash_sales_lst

def wash_packages(section):
    wash_packages_lst = []

    reports = section.get("reports")
    subtotals = section.get("subtotals")
    
    for report in reports:
        wash_package_structure = {
            "Wash_packages_Description":report.get("description"),
            "Wash_packages_price":report.get("price"),
            "Wash_packages_quantity":report.get("quantity"),
            "Wash_packages_amount":report.get("amount"),
        }
        wash_packages_lst.append(wash_package_structure)
    
    for subtotal in subtotals:
        wash_package_structure = {
            "Wash_packages_Description":subtotal.get("description"),
            "Wash_packages_price":subtotal.get("price"),
            "Wash_packages_quantity":subtotal.get("quantity"),
            "Wash_packages_amount":subtotal.get("amount"),
        }
        wash_packages_lst.append(wash_package_structure)

    return wash_packages_lst
 
def wash_extra_services(section):
    wash_extra_service_lst = []
    reports = section.get("reports")
    subtotals  = section.get("subtotals")
    
    for report in reports:
        wash_extra_structure = {
            "Wash_Extra_Services_Description":report.get("description"),
            "Wash_Extra_Services_price":report.get("price"),
            "Wash_Extra_Services_quantity":report.get("quantity"),
            "Wash_Extra_Services_amout":report.get("amount")
        }
        wash_extra_service_lst.append(wash_extra_structure)
        
    for total in subtotals:
        wash_extra_structure = {
            "Wash_Extra_Services_Description":total.get("description"),
            "Wash_Extra_Services_price":total.get("price"),
            "Wash_Extra_Services_quantity":total.get("quantity"),
            "Wash_Extra_Services_amout":total.get("amount")
        }
        wash_extra_service_lst.append(wash_extra_structure)  
        
    return wash_extra_service_lst  

def gross_wash_sales(section):
    gross_wash_sales_lst = []
    gross_wash_sale_structure = {
            "Gross_Wash_Sales":section.get("totalAmount")
        }
        
    gross_wash_sales_lst.append(gross_wash_sale_structure)
    
    return gross_wash_sales_lst

def less_free_wash_rdmd(section):
    less_wash_sales_rdmd_lst = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    
    for report in reports:
        less_free_wash_rmd_structure = {
            "Less_free_wash_rdmd_Description":report.get("description"),
            "Less_free_wash_rdmd_price":report.get("price"),
            "Less_free_wash_rdmd_quantity":report.get("quantity"),
            "Less_free_wash_rdmd_amount":report.get("amount"),
        }
        less_wash_sales_rdmd_lst.append(less_free_wash_rmd_structure)
    
    for subtotal in subtotals:
        less_free_wash_rmd_structure = {
            "Less_free_wash_rdmd_Description":subtotal.get("description"),
            "Less_free_wash_rdmd_price":subtotal.get("price"),
            "Less_free_wash_rdmd_quantity":subtotal.get("quantity"),
            "Less_free_wash_rdmd_amount":subtotal.get("amount"),
        }
        less_wash_sales_rdmd_lst.append(less_free_wash_rmd_structure)
    return less_wash_sales_rdmd_lst

def less_wash_discounts(section):
    less_wash_discounts_lst = []
    reports  = section.get("reports")  
    subtotals  = section.get("subtotals")  
    for report in reports:
        less_wash_discount_structure = {
            "Less_Wash_Discounts_Description":report.get("description"),
            "Less_Wash_Discounts_Price":report.get("price"),
            "Less_Wash_Discounts_quantity":report.get("quantity"),
            "Less_Wash_Discounts_amount":report.get("amount")
        }
        less_wash_discounts_lst.append(less_wash_discount_structure)

    for subtotal in subtotals:
        less_wash_discount_structure = {
            "Less_Wash_Discounts_Description":subtotal.get("description"),
            "Less_Wash_Discounts_Price":subtotal.get("price"),
            "Less_Wash_Discounts_quantity":subtotal.get("quantity"),
            "Less_Wash_Discounts_amount":subtotal.get("amount")
        }
        less_wash_discounts_lst.append(less_wash_discount_structure)    

    return less_wash_discounts_lst
    

def less_loyality_disc(section):
    less_loyality_disc_lst = []

    reports = section.get("reports")
    subtotals = section.get("subtotals")
    
    for report in reports:
        less_loyality_disc_structure = {
            "Less_Loyalty_disc_description":report.get("description"),
            "Less_Loyalty_disc_price":report.get("price"),
            "Less_Loyalty_disc_quantity":report.get("quantity"),
            "Less_Loyalty_disc_amount":report.get("amount")
        }
        less_loyality_disc_lst.append(less_loyality_disc_structure)
        
    for subtotal in subtotals:
        less_loyality_disc_structure = {
            "Less_Loyalty_disc_description":subtotal.get("description"),
            "Less_Loyalty_disc_price":subtotal.get("price"),
            "Less_Loyalty_disc_quantity":subtotal.get("quantity"),
            "Less_Loyalty_disc_amount":subtotal.get("amount")
        }
        less_loyality_disc_lst.append(less_loyality_disc_structure)

    return less_loyality_disc_lst

def net_site_sales(section):
    # net_site_sales_lst = []
    # net_site_sales_structue={
    #             "Net_site_sales":section.get("totalAmount")
    #         }
    # net_site_sales_lst.append(net_site_sales_structue)

    return section.get("totalAmount")
 
def arm_plans_sold(section):
    # arm_plans_sold_lst = []
    # reports = section.get("reports")
    # subtotals  = section.get("subtotals")
    
    # for report in reports:
    #     arm_plans_sold_structure = {
    #         "Arm_plan_sold_description":report.get("description"),
    #         "Arm_plan_sold_price":report.get("price"),
    #         "Arm_plan_sold_quantity":report.get("quantity"),
    #         "Arm_plan_sold_amount":report.get("amount")
    #     }
    #     arm_plans_sold_lst.append(arm_plans_sold_structure)
        
    # for subtotal in subtotals:
    #     arm_plans_sold_structure = {
    #         "Arm_plan_sold_description":subtotal.get("description"),
    #         "Arm_plan_sold_price":subtotal.get("price"),
    #         "Arm_plan_sold_quantity":subtotal.get("quantity"),
    #         "Arm_plan_sold_amount":subtotal.get("amount")
    #     }
    #     arm_plans_sold_lst.append(arm_plans_sold_structure)
        
    return section.get("totalQuantity",0.0)   


def arm_plans_recharged(section):
    arm_plans_recharged_lst = []
    reports  = section.get("reports")   
    subtotals  = section.get("subtotals")   
    
    for report in reports:
        arm_plans_recharged_structure = {
            "Arm_plan_recharged_description":report.get("description"),
            "Arm_plan_recharged_price":report.get("price"),
            "Arm_plan_recharged_quantity":report.get("quantity"),
            "Arm_plan_recharged_amount":report.get("amount")
        } 
        arm_plans_recharged_lst.append(arm_plans_recharged_structure)
        
    for subtotal in subtotals:
        arm_plans_recharged_structure = {
            "Arm_plan_recharged_description":subtotal.get("description"),
            "Arm_plan_recharged_price":subtotal.get("price"),
            "Arm_plan_recharged_quantity":subtotal.get("quantity"),
            "Arm_plan_recharged_amount":subtotal.get("amount")
        } 
        arm_plans_recharged_lst.append(arm_plans_recharged_structure)
    
    return arm_plans_recharged_lst

def arm_planes_reedemed(section):
    data={}
    
    # arm_plans_reedemed_lst = []
    # reports  = section.get("reports")
    # subtotals = section.get("subtotals")     
    
    # for report in reports:
    #     arm_plans_reedemed_structure = {
    #         "Arm_plan_redeemed_description":report.get("description"),
    #         "Arm_plan_redeemed_price":report.get("price"),
    #         "Arm_plan_redeemed_quantity":report.get("quantity"),
    #         "Arm_plan_redeemed_amount":report.get("amount"),

    #     }
    
    #     arm_plans_reedemed_lst.append(arm_plans_reedemed_structure)
    # for subtotal in subtotals:
    #     arm_plans_reedemed_structure = {
    #         "Arm_plan_redeemed_description":subtotal.get("description"),
    #         "Arm_plan_redeemed_price":subtotal.get("price"),
    #         "Arm_plan_redeemed_quantity":subtotal.get("quantity"),
    #         "Arm_plan_redeemed_amount":subtotal.get("amount"),

    #     }
    
    #     arm_plans_reedemed_lst.append(arm_plans_reedemed_structure)    

    data["arm_plans_reedemed_cnt"] = section.get("totalQuantity")
    data["arm_plans_reedemed_amt"] = section.get("totalAmount",0.0)*(-1)
    return data
    
    
def arm_plans_terminated(section):
    arm_plans_terminated_lst = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    for report in reports:
        arm_plan_terminated_structure ={
            "Arm_plans_terminated_description":report.get("description"),
            "Arm_plans_terminated_price":report.get("price"),
            "Arm_plans_terminated_quantity":report.get("quantity"),
            "Arm_plans_terminated_amount":report.get("amount")
        }
        arm_plans_terminated_lst.append(arm_plan_terminated_structure)
    for subtotal in subtotals:
        arm_plan_terminated_structure ={
            "Arm_plans_terminated_description":subtotal.get("description"),
            "Arm_plans_terminated_price":subtotal.get("price"),
            "Arm_plans_terminated_quantity":subtotal.get("quantity"),
            "Arm_plans_terminated_amount":subtotal.get("amount")
        }
        arm_plans_terminated_lst.append(arm_plan_terminated_structure)
           
    return arm_plans_terminated_lst


def prepaid_sold(section):
    prepaid_sold_lst = []
    reports = section.get("reports")
    subtotals  = section.get("subtotals")
    
    for report in reports:
        prepaid_sold_structure = {
            "Prepaids_Sold_description":report.get("description"),
            "Prepaids_Sold_price":report.get("price"),
            "Prepaids_Sold_quantity":report.get("quantity"),
            "Prepaids_Sold_amount":report.get("amount"),
        }
        prepaid_sold_lst.append(prepaid_sold_structure)
        
    for subtotal in subtotals:
        prepaid_sold_structure = {
            "Prepaids_Sold_description":subtotal.get("description"),
            "Prepaids_Sold_price":subtotal.get("price"),
            "Prepaids_Sold_quantity":subtotal.get("quantity"),
            "Prepaids_Sold_amount":subtotal.get("amount"),
        }
        prepaid_sold_lst.append(prepaid_sold_structure)
        
    return prepaid_sold_lst
    
def less_prepaid_reedemed(section):
    less_prepaid_reedemed_lst = []
    
    reports  = section.get("reports")
    subtotals = section.get("subtotals")
    for report in reports:
        less_prepaid__reedemed_structure ={
            "Less_prepaids_redeemed_description":report.get("description"),
            "Less_prepaids_redeemed_price":report.get("price"),
            "Less_prepaids_redeemed_quantity":report.get("quantity"),
            "Less_prepaids_redeemed_amount":report.get("amount")
        }
        
        less_prepaid_reedemed_lst.append(less_prepaid__reedemed_structure)
        
    for subtotal in subtotals:
        less_prepaid__reedemed_structure ={
            "Less_prepaids_redeemed_description":subtotal.get("description"),
            "Less_prepaids_redeemed_price":subtotal.get("price"),
            "Less_prepaids_redeemed_quantity":subtotal.get("quantity"),
            "Less_prepaids_redeemed_amount":subtotal.get("amount")
        }
        
        less_prepaid_reedemed_lst.append(less_prepaid__reedemed_structure)
    
    return less_prepaid_reedemed_lst
 
def online_sold(section):
    data = []
    reports = section.get("reports")
    subtotals= section.get("subtotals")
    
    for report in  reports:
        online_sold_structure ={
            "online_sold_description":report.get("description"),
            "online_sold_price":report.get("price"),
            "online_sold_quantity":report.get("quantity"),
            "online_sold_amount":report.get("amount")
        }
        data.append(online_sold_structure)
        
    for subtotal in  subtotals:
        online_sold_structure ={
            "online_sold_description":subtotal.get("description"),
            "online_sold_price":subtotal.get("price"),
            "online_sold_quantity":subtotal.get("quantity"),
            "online_sold_amount":subtotal.get("amount")
        }
        data.append(online_sold_structure)
    
    return data
        
def less_online_reedemed(section):
    data = [] 
    reports =section.get("reports")
    subtotals= section.get("subtotals")
    
    for report in reports:
        less_online_reedemed_strcture = {
            "Less_online_redeemed_description":report.get("description"),
            "Less_online_redeemed_price":report.get("price"),
            "Less_online_redeemed_quantity":report.get("quantity"),
            "Less_online_redeemed_amount":report.get("amount")
        }
        data.append(less_online_reedemed_strcture)
    for subtotal in subtotals:
        less_online_reedemed_strcture = {
            "Less_online_redeemed_description":subtotal.get("description"),
            "Less_online_redeemed_price":subtotal.get("price"),
            "Less_online_redeemed_quantity":subtotal.get("quantity"),
            "Less_online_redeemed_amount":subtotal.get("amount")
        }
        data.append(less_online_reedemed_strcture)
    return data
    

def free_wash_issued(section):
    data=[]
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    for report in reports:
        free_wash_issued_structure = {
            "Free_washes_issued_description":report.get("description"),
            "Free_washes_issued_price":report.get("price"),
            "Free_washes_issued_quantity":report.get("quantity"),
            "Free_washes_issued_amount":report.get("amount")
        }
        
        data.append(free_wash_issued_structure)
    for subtotal in subtotals:
        free_wash_issued_structure = {
            "Free_washes_issued_description":subtotal.get("description"),
            "Free_washes_issued_price":subtotal.get("price"),
            "Free_washes_issued_quantity":subtotal.get("quantity"),
            "Free_washes_issued_amount":subtotal.get("amount")
        }
        
        data.append(free_wash_issued_structure)
        
    return data

def less_paidouts(section):
    data=[]
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    
    for report in reports:
        less_paidouts_structure = {
            "Less_paidouts_description":report.get("description"),
            "Less_paidouts_price":report.get("price"),
            "Less_paidouts_quantity":report.get("quantity"),
            "Less_paidouts_amount":report.get("amount")
        }
        data.append(less_paidouts_structure)
    for total in subtotals:
        less_paidouts_structure = {
            "Less_paidouts_description":total.get("description"),
            "Less_paidouts_price":total.get("price"),
            "Less_paidouts_quantity":total.get("quantity"),
            "Less_paidouts_amount":total.get("amount")
        }
        data.append(less_paidouts_structure)
    
    return data

def total_to_account_for(section):
    # data = []
    

    # total_to_account_for_structure = {
    #     "TOTAL_TO_ACCOUNT_FOR:":section.get("totalAmount")
    # }

    # data.append(total_to_account_for_structure)
    
    return section.get("totalAmount",0.0)

def deposits(section):
    data = []
    reports  = section.get("reports")
    subtotals  = section.get("subtotals")
    for report in reports:
        deposit_structure = {
            "Deposits_description":report.get("description"),
            "Deposits_price":report.get("price"),
            "Deposits_quantity":report.get("quantity"),
            "Deposits_amount":report.get("amount")
        }
        data.append(deposit_structure)
    
    for subtotal in subtotals:
        deposit_structure = {
            "Deposits_description":subtotal.get("description"),
            "Deposits_price":subtotal.get("price"),
            "Deposits_quantity":subtotal.get("quantity"),
            "Deposits_amount":subtotal.get("amount")
        }
        data.append(deposit_structure)
        
    return data

def total_xpt_cash(section):
    data =[]
    total_xpt_structure ={
        "TOTAL XPT CASH:":section.get("totalAmount")
    }
    data.append(total_xpt_structure)
    
    return data

def house_accounts(section):
    data = []
    reports = section.get('reports')
    subtotals  = section.get('subtotals')
    for report in reports:
        house_accounts_structure ={
            "House_accounts_description":report.get("description"),
            "House_accounts_price":report.get("price"),
            "House_accounts_quantity":report.get("quantity"),
            "House_accounts_amount":report.get("amount")
        }
        data.append(house_accounts_structure)
        
    for subtotal in subtotals:
        house_accounts_structure ={
            "House_accounts_description":subtotal.get("description"),
            "House_accounts_price":subtotal.get("price"),
            "House_accounts_quantity":subtotal.get("quantity"),
            "House_accounts_amount":subtotal.get("amount")
        }
        data.append(house_accounts_structure)
        
    return data
    
# def over_short(section):

def cash(section):
    data=[]
    cash_structure ={
        "CASH:":section.get("totalAmount")
    }
    data.append(cash_structure)

    return data

def xpt_acceptors(section):
    data = []
    xpt_acceptor_structure = {
        "XPT ACCEPTORS:":section.get("totalAmount")
    }
    
    data.append(xpt_acceptor_structure)
    
    return data
  
def xpt_dispensers(section):
    data = []
    xpt_dispenser_structure ={
        "XPT DISPENSERS:":section.get("totalAmount")
    }  
    
    data.append(xpt_dispenser_structure)
    
    return data

def total_function(section):
    data = []
    total_structure = {
        "TOTAL:":section.get("totalAmount")
    }
    data.append(total_structure)
    
    return data

def credit_card(section):
    data = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    
    for report in reports:
        credit_card_structure = {
            "Credit_Card_description":report.get("description"),
            "Credit_Card_price":report.get("price"),
            "Credit_Card_quantity":report.get("quantity"),
            "Credit_Card_amount":report.get("amount")
        }
        data.append(credit_card_structure)
    
    for subtotal in subtotals:
        credit_card_structure = {
            "Credit_Card_description":subtotal.get("description"),
            "Credit_Card_price":subtotal.get("price"),
            "Credit_Card_quantity":subtotal.get("quantity"),
            "Credit_Card_amount":subtotal.get("amount")
        }
        data.append(credit_card_structure)
    return data

def other_tenders(section):
    data = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    for report in reports:
        other_tender_structure = {
            "Other_tenders_description":report.get("description"),
            "Other_tenders_price":report.get("price"),
            "Other_tenders_quantity":report.get("quantity"),
            "Other_tenders_amount":report.get("amount")
        }
        data.append(other_tender_structure)
    for subtotal in subtotals:
        other_tender_structure = {
            "Other_tenders_description":subtotal.get("description"),
            "Other_tenders_price":subtotal.get("price"),
            "Other_tenders_quantity":subtotal.get("quantity"),
            "Other_tenders_amount":subtotal.get("amount")
        }
        data.append(other_tender_structure)
        
    return data

def xpt_balancing(section):
    data = []
    xpt_balancing_structure = {
        "XPT BALANCING:":section.get("totalAmount")
    }
    data.append(xpt_balancing_structure)
    
    return data


def report_balance(section):
    data = []
    report_balance_structure = {
        "REPORT BALANCE:":section.get("totalAmount")
    }  
    
    data.append(report_balance_structure)
    
    return data

def picture_mismatch(section):
    data = []
    picture_mismatch_structure = {
        "PICTURE MISMATCH:":section.get("totalCount")
    }
    data.append(picture_mismatch_structure)
    
    return data


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
    return ""
    for index,data in enumerate(dictionary_data):
                if index == 0:
                    column_gap = 0 if is_first_dictionary else 2
                    append_dict_to_excel(file_name,data,column_gap)
                else:
                    append_dict_to_excel(file_name,data,0,False)


def report_data_extractor(report_data):
    single_site_report = {}
    if report_data:
        gsviews =  report_data.get("gsviews")
        gsviews_0 = gsviews[0]

        sections = gsviews_0.get("sections")
        for section in sections:
            text        = section.get("text")
            
            
            if text=="WASH SALES-":
                wash_sales_ret = wash_sales(section)#,file_path
                print(f"Washsales:{wash_sales_ret}")
                single_site_report["car_count"] = wash_sales_ret #car count
                
                    
                
            elif text=="WASH PACKAGES-":
                wash_packages(section)
                    
            elif text=="WASH EXTRA SERVICES-":
                wash_extra_services(section)
                    
            elif text=="GROSS WASH SALES-":
                gross_wash_sales(section)
                
            
            elif text=="LESS FREE WASH RDMD-":
                less_free_wash_rdmd(section)
                    
            
            elif text=="LESS WASH DISCOUNTS-":
                less_wash_discounts(section)
                    
            elif text=="LESS LOYALTY DISC-":
                less_loyality_disc(section)
            
            elif text == "NET SITE SALES:":      
                net_site_sales_value = net_site_sales(section)
                single_site_report['net_sales']=net_site_sales_value
            
            elif text=="ARM PLANS SOLD-":
                arm_plans_sold_cnt= arm_plans_sold(section)
                single_site_report["arm_plans_sold_cnt"] = arm_plans_sold_cnt
                
                
            elif text=="ARM PLANS RECHARGED-":
                arm_plans_recharged(section)
                
                
            elif text=="ARM PLANS REDEEMED-":
                arm_plans_reedemed_value= arm_planes_reedemed(section)
                single_site_report["arm_plans_reedemed_cnt"] = arm_plans_reedemed_value.get("arm_plans_reedemed_cnt")
                single_site_report["arm_plans_reedemed_amt"] =arm_plans_reedemed_value.get("arm_plans_reedemed_amt")
                    
            elif text=="ARM PLANS TERMINATED-":
                arm_plans_terminated(section)
            
            elif text=="PREPAIDS SOLD-":
                prepaid_sold(section)
            
            elif text=="LESS PREPAIDS REDEEMED-":
                less_prepaid_reedemed(section)
                
            elif text == "ONLINE SOLD-":
                online_sold(section)
                
            elif text == "LESS ONLINE REDEEMED-":
                less_online_reedemed(section)
                
            elif text=="FREE WASHES ISSUED-":
                free_wash_issued(section)
                
            elif text=="LESS PAIDOUTS:":
                less_paidouts(section)
                
            elif text=="TOTAL TO ACCOUNT FOR:":
                total_revenue_val = total_to_account_for(section)
                single_site_report['total_revenue'] = total_revenue_val
            
            elif text=="DEPOSITS-":
                deposits(section)
                
            elif text=="TOTAL XPT CASH:":
                total_xpt_cash(section)
                
            elif text=="HOUSE ACCOUNTS-":
                house_accounts(section)
                
            # elif text =="OVER / SHORT (-)":
            #     over_short_lst = over_short(section)
            
            elif text=="CASH:":
                cash(section)
                
            elif text=="XPT ACCEPTORS:":
                xpt_acceptors(section)
                
            elif text =="XPT DISPENSERS:":
                xpt_dispensers(section)
                
            elif text =="TOTAL:":
                total_function(section)
                
            elif text=="CREDIT CARD:":
                credit_card(section)
                
                
            elif text=="OTHER TENDERS:":
                other_tenders(section)
                
            elif text=="XPT BALANCING:":
                xpt_balancing(section)
                
            elif text=="REPORT BALANCE:":
                report_balance(section)
                
            elif text=="PICTURE MISMATCH:":
                picture_mismatch(section)
            
    return single_site_report

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


def generate_weekly_report(path,monday_date_str,friday_date_str,saturday_date_str, sunday_date_str):
    site_watch_report={}
    
    
    for index,site in sites_df.iterrows():
        # all_dictionaries_lst = []
        combined_data = {}
        
        site_dict = site.to_dict()
        cookiesfile_name = f"{(site_dict.get('Organization')).strip().replace('-','_')}.pkl"
        # print(cookiesfile_name)
        print(site_dict)

        cookies_file = os.path.join(cookies_path,cookiesfile_name)
    
        
        client = sitewatchClient(cookies_file=cookies_file)
        employCode = site_dict.get("employee")
        password = site_dict.get("password")
        locationCode = site_dict.get("Organization")
        client_name = site_dict.get("client_name2")
        remember = 1
        file_path=f"sitewatch_{client_name.strip().replace(' ','_')}_{monday_date_str}_{sunday_date_str}.xlsx"
        
        file_path = os.path.join(path,file_path)
        session_chek = client.check_session_auth(timeout=30)

        if not session_chek:
            token = client.login(employeeCode=employCode,password=password,locationCode=locationCode,remember=1)
            print(token)
        
        session_chek = client.check_session_auth(timeout=30)
        if session_chek:
            reportOn=site_dict.get("site")
            id=site_dict.get("id")
            idname=site_dict.get("id_name")
            request_id1 = client.get_general_sales_report_request_id(reportOn,id,idname,monday_date_str,friday_date_str)
            request_id1_1 =client.get_activity_by_date_proft_request_id(reportOn,monday_date_str,friday_date_str) #for labour hours 

            if request_id1 and request_id1_1:
                report_data = client.get_report(reportOn,request_id1)
                # print(f"report data: {report_data}")
                extracted_data1= report_data_extractor(report_data)
                car_count_monday_to_friday = extracted_data1.get("car_count",0)
                arm_plans_reedemed_monday_to_friday_cnt = extracted_data1.get("arm_plans_reedemed_cnt",0)
                arm_plans_reedemed_monday_to_friday_amt = extracted_data1.get("arm_plans_reedemed_amt")
                retail_car_count_monday_to_friday=abs(car_count_monday_to_friday - arm_plans_reedemed_monday_to_friday_cnt)
                
                net_sales_amt= extracted_data1.get("net_sales",0.0)
                retail_revenue__monday_to_friday = round((net_sales_amt - arm_plans_reedemed_monday_to_friday_amt),2)
                
                total_revenue_val = round(extracted_data1.get("total_revenue",0.0),2)
                
                arm_plans_sold_cnt1 = extracted_data1.get("arm_plans_sold_cnt")
                labour_hours_monday_to_friday=client.get_labour_hours(reportOn,request_id1_1)
                cars_per_labour_hour_monday_to_friday = round((car_count_monday_to_friday/labour_hours_monday_to_friday),2)
                
                mon_fri_data = {
                    "car_count_monday_to_friday":car_count_monday_to_friday,
                    "arm_plans_reedemed_monday_to_friday_cnt":arm_plans_reedemed_monday_to_friday_cnt,
                    "retail_car_count_monday_to_friday":retail_car_count_monday_to_friday,
                    "retail_revenue_monday_to_friday":retail_revenue__monday_to_friday,
                    "total_revenue_monday_to_friday":total_revenue_val,
                    "labour_hours_monday_to_friday":labour_hours_monday_to_friday,
                    "cars_per_labour_hour_monday_to_friday":cars_per_labour_hour_monday_to_friday
                }
                combined_data.update(mon_fri_data)
            
            request_id2 = client.get_general_sales_report_request_id(reportOn,id,idname,saturday_date_str,sunday_date_str)
            request_id2_2 =client.get_activity_by_date_proft_request_id(reportOn,saturday_date_str,sunday_date_str) #for labour hours
             
            if request_id2 and request_id2_2:
                report_data = client.get_report(reportOn,request_id2)
                # print(f"data2:{report_data}")
                extracted_data2= report_data_extractor(report_data)
                car_count_saturday_sunday = extracted_data2.get("car_count",0)
                arm_plans_reedemed_saturday_sunday_cnt = extracted_data2.get("arm_plans_reedemed_cnt",0)
                arm_plans_reedemed_saturday_sunday_amt = extracted_data2.get("arm_plans_reedemed_amt")
                
                retail_car_count_saturday_sunday =abs(car_count_saturday_sunday- arm_plans_reedemed_saturday_sunday_cnt)
                
                net_sales_amt2= extracted_data2.get("net_sales",0.0)
                retail_revenue__saturday_sunday = round((net_sales_amt2 - arm_plans_reedemed_saturday_sunday_amt),2)
                
                total_revenue_val2 = round(extracted_data2.get("total_revenue",0.0),2)
                
                arm_plans_sold_cnt2 = extracted_data2.get("arm_plans_sold_cnt")
                
                labour_hours_saturday_sunday=client.get_labour_hours(reportOn,request_id2_2)
                cars_per_labour_hour_saturday_sunday = round((car_count_saturday_sunday/labour_hours_saturday_sunday),2)
                
                sat_sun_data = {
                    "car_count_saturday_sunday":car_count_saturday_sunday,
                    "arm_plans_reedemed_saturday_sunday":arm_plans_reedemed_saturday_sunday_cnt,
                    "retail_car_count_saturday_sunday":retail_car_count_saturday_sunday,
                    "retail_revenue_saturday_sunday":retail_revenue__saturday_sunday,
                    "total_revenue_saturday_sunday":total_revenue_val2,
                    "labour_hours_saturday_sunday":labour_hours_saturday_sunday,
                    "cars_per_labour_hour_saturday_sunday":cars_per_labour_hour_saturday_sunday
                }
                
                combined_data.update(sat_sun_data)
                
                arm_plans_sold_total_cnt = sum([arm_plans_sold_cnt1,arm_plans_sold_cnt2])
                total_arm_planmembers_cnt = sum([arm_plans_sold_cnt1,arm_plans_sold_cnt2,
                                                                  arm_plans_reedemed_monday_to_friday_cnt,
                                                                  arm_plans_reedemed_saturday_sunday_cnt])
                
                total_retail_car_cnt = sum([retail_car_count_monday_to_friday,retail_car_count_saturday_sunday])
                conversion_rate  = round((arm_plans_sold_total_cnt/total_retail_car_cnt)*100,2)
                #combines values update 
                #combined_data["total_cars"] = sum([car_count_monday_to_friday,car_count_saturday_sunday])
                combined_data["total_revenue"] = sum([total_revenue_val,total_revenue_val2])
                combined_data["arm_plans_sold_cnt"] = arm_plans_sold_total_cnt
                combined_data["total_arm_planmembers_cnt"] = total_arm_planmembers_cnt
                combined_data["conversion_rate"]= conversion_rate
            # print(f"combined data:{combined_data}")
            site_watch_report[client_name]=combined_data
                        
    
    return site_watch_report

if __name__=="__main__":
    import pandas as pd
    monday_date_str, sunday_date_str = get_week_dates()
    print(monday_date_str,sunday_date_str)
    monday_date_str="2024-06-03"
    friday_date_str = "2024-06-07"
    saturday_date_str = "2024-06-08"
    sunday_date_str="2024-06-09"  #YMD
    
    report = generate_weekly_report("",monday_date_str,friday_date_str,saturday_date_str, sunday_date_str)
    print("\n"*6)
    print(report)
    with open("sitewatch_report.json","w") as f:
        json.dump(report,f,indent=4)
    # with open("SPKLUS-002.json",'r') as f:
    #     data = json.load(f)
        
    # df = pd.read_json("SPKLUS-002.json")
    
    # df.to_excel("sitewatch.xlsx")