from datetime import datetime, timedelta
import os
import json
import time

import pandas as pd 
from datetime import datetime, timedelta
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
import xlsxwriter

from sitewatch4 import sitewatchClient
from sitewatch4 import generate_past_4_weeks_days
from sitewatch4 import generate_past_4_week_days_full
import random

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
    data=None
    arm_plans_recharged_lst = []
    reports  = section.get("reports")   
    subtotals  = section.get("subtotals")   
    
    # for report in reports:
    #     arm_plans_recharged_structure = {
    #         "Arm_plan_recharged_description":report.get("description"),
    #         "Arm_plan_recharged_price":report.get("price"),
    #         "Arm_plan_recharged_quantity":report.get("quantity"),
    #         "Arm_plan_recharged_amount":report.get("amount")
    #     } 
    #     arm_plans_recharged_lst.append(arm_plans_recharged_structure)
    data = subtotals[0].get("amount")
    
    return data
    
        
    # for subtotal in subtotals:
    #     arm_plans_recharged_structure = {
    #         "Arm_plan_recharged_description":subtotal.get("description"),
    #         "Arm_plan_recharged_price":subtotal.get("price"),
    #         "Arm_plan_recharged_quantity":subtotal.get("quantity"),
    #         "Arm_plan_recharged_amount":subtotal.get("amount")
    #     } 
    #     arm_plans_recharged_lst.append(arm_plans_recharged_structure)
    
    # return arm_plans_recharged_lst

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
                arm_reachrged_amt =arm_plans_recharged(section)
                print("arm plans rechanged:",arm_reachrged_amt)
                single_site_report["arm_plans_recharged_amt"]=arm_reachrged_amt
                
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
    
    # Find the current week's Friday, Saturday, and Sunday dates
    current_week_friday = current_week_monday + timedelta(days=4)
    current_week_saturday = current_week_monday + timedelta(days=5)
    current_week_sunday = current_week_monday + timedelta(days=6)
    
    # Format the dates in yyyy-mm-dd format
    monday_date_str = current_week_monday.strftime("%Y-%m-%d")
    friday_date_str = current_week_friday.strftime("%Y-%m-%d")
    saturday_date_str = current_week_saturday.strftime("%Y-%m-%d")
    sunday_date_str = current_week_sunday.strftime("%Y-%m-%d")
    
    return monday_date_str, friday_date_str, saturday_date_str, sunday_date_str



def generate_weekly_report(path,monday_date_str,friday_date_str,saturday_date_str, sunday_date_str):
    site_watch_report={}
    
    is_location_code_success=False
    success_location_code=None,None
    location_codes_a = ['SPKLUS-001', 'SPKLUS-002', 'SPKLUS-003', 'SPKLUS-004', 'SPKLUS-005', 'SPKLUS-006', 'SPKLUS-007', 'SPKLUS-008', 'SPKLUS-009']
    location_codes_b = ['SPKLUS-012', 'SPKLUS-013', 'SPKLUS-014', 'SPKLUS-015']
    for index,site in sites_df.iterrows():
        # all_dictionaries_lst = []
        site_dict = site.to_dict()
        slno=site_dict.get("slno")
        locationCode = site_dict.get("Organization")
        
        
        while True:
                      
            try:
            
                combined_data = {}
                
                
                cookiesfile_name = f"{(site_dict.get('Organization')).strip().replace('-','_')}.pkl"
                # print(cookiesfile_name)
                print(site_dict)
                

                cookies_file = os.path.join(cookies_path,cookiesfile_name)
            
                
                client = sitewatchClient(cookies_file=cookies_file)
                employCode = site_dict.get("employee")
                password = site_dict.get("password")
                
                
                if is_location_code_success:
                    locationCode_old,slno_old =success_location_code
                    location_a_range =range(1,10)
                    location_b_range =range(10,14)
                    if slno_old in location_a_range and slno in location_a_range: 
                        locationCode=locationCode_old
                    
                    elif slno_old in location_b_range and slno in location_b_range:
                        locationCode=locationCode_old
                        
                print(f"\n location code used :{locationCode}")
                client_name = site_dict.get("client_name2")
                remember = 1
                # file_path=f"sitewatch_{client_name.strip().replace(' ','_')}_{monday_date_str}_{sunday_date_str}.xlsx"
                
                # file_path = os.path.join(path,file_path)
                session_chek = client.check_session_auth(timeout=60)

                if not session_chek:
                    token = client.login(employeeCode=employCode,password=password,locationCode=locationCode,remember=1)
                    print(token)
                
                session_chek = client.check_session_auth(timeout=60)
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
                        arm_plans_reedemed_monday_to_friday_amt = abs(extracted_data1.get("arm_plans_reedemed_amt",0))#need abs because will nee to mek it postive 
                        arm_plans_recharged_amt_monday_to_friday_amt =extracted_data1.get("arm_plans_recharged_amt",0)
                        
                        retail_car_count_monday_to_friday=(car_count_monday_to_friday - arm_plans_reedemed_monday_to_friday_cnt)
                        
                        net_sales_amt= extracted_data1.get("net_sales",0.0)
                        # retail_revenue__monday_to_friday = round((net_sales_amt - arm_plans_reedemed_monday_to_friday_amt),2)
                        total_revenue_val = round(extracted_data1.get("total_revenue",0.0),2)
                        
                        if client_name=="Sudz - Beverly":
                            retail_revenue__monday_to_friday = round((total_revenue_val - arm_plans_recharged_amt_monday_to_friday_amt),2)
                        else:
                            retail_revenue__monday_to_friday = round((net_sales_amt - arm_plans_reedemed_monday_to_friday_amt),2)
                            
                            
                        arm_plans_sold_cnt1 = extracted_data1.get("arm_plans_sold_cnt")
                        labour_hours_monday_to_friday=client.get_labour_hours(reportOn,request_id1_1)
                        cars_per_labour_hour_monday_to_friday = round((car_count_monday_to_friday/labour_hours_monday_to_friday),2) if labour_hours_monday_to_friday !=0 else "" 
                        
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
                        arm_plans_reedemed_saturday_sunday_amt = extracted_data2.get("arm_plans_reedemed_amt",0)
                        arm_plans_recharged_amt_saturday_sunday_amt =extracted_data2.get("arm_plans_recharged_amt",0)
                        
                        retail_car_count_saturday_sunday =(car_count_saturday_sunday- arm_plans_reedemed_saturday_sunday_cnt)
                        
                        net_sales_amt2= extracted_data2.get("net_sales",0.0)
                        
                        total_revenue_val2 = round(extracted_data2.get("total_revenue",0.0),2)
                        
                        #print(f"retail rev : total rev {total_revenue_val2}-{arm_plans_recharged_amt_saturday_sunday_amt} ={total_revenue_val2-arm_plans_recharged_amt_saturday_sunday_amt}")
                        
                        if client_name=="Sudz - Beverly":
                            retail_revenue__saturday_sunday = round((total_revenue_val2 - arm_plans_recharged_amt_saturday_sunday_amt),2)
                        else:
                            retail_revenue__saturday_sunday = round((net_sales_amt2- arm_plans_reedemed_saturday_sunday_amt),2)
                            
                        arm_plans_sold_cnt2 = extracted_data2.get("arm_plans_sold_cnt")
                        
                        labour_hours_saturday_sunday=client.get_labour_hours(reportOn,request_id2_2)
                        cars_per_labour_hour_saturday_sunday = round((car_count_saturday_sunday/labour_hours_saturday_sunday),2) if labour_hours_saturday_sunday !=0 else ""
                        
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
                        # total_arm_planmembers_cnt = sum([arm_plans_sold_cnt1,arm_plans_sold_cnt2,
                        #                                                 arm_plans_reedemed_monday_to_friday_cnt,
                        #                                                 arm_plans_reedemed_saturday_sunday_cnt])
                        
                        total_retail_car_cnt = sum([retail_car_count_monday_to_friday,retail_car_count_saturday_sunday])
                        conversion_rate  = round((arm_plans_sold_total_cnt/total_retail_car_cnt)*100,2) if total_retail_car_cnt !=0 else ""
                        #combines values update 
                        #combined_data["total_cars"] = sum([car_count_monday_to_friday,car_count_saturday_sunday])
                        total_members_req_id = client.get_plan_analysis_request_id(sunday_date_str,reportOn)
                        
                        total_arm_planmembers_cnt = client.get_total_plan_members(total_members_req_id,reportOn)
                        
                        #Past 4 weeks logic 
                        past_week_day1, past_week_day2= generate_past_4_weeks_days(monday_date_str)
                        
                        request_id3 = client.get_general_sales_report_request_id(reportOn,id,idname,past_week_day1, past_week_day2)
                        report_data3 = client.get_report(reportOn,request_id3)
                        extracted_data3= report_data_extractor(report_data3)
                        past_4_week_cnt = extracted_data3.get("car_count",0)
                        past_4_week_arm_plans_sold =extracted_data3.get("arm_plans_sold_cnt")
                        past_4_weeks_arm_plans_reedemed_cnt = extracted_data3.get("arm_plans_reedemed_cnt",0)
                        past_4_weeks_retail_car_count = past_4_week_cnt - past_4_weeks_arm_plans_reedemed_cnt
                        past_4_weeks_total_revenue = extracted_data3.get("total_revenue")
                        
                        print(f"past week cnt : {past_4_week_cnt}")
                        past_4_weeks_conversion_rate = (past_4_week_arm_plans_sold/past_4_weeks_retail_car_count)*100
                        
                        #Past 4 weeks splieed days mon-day st -sun
                        full_weeks_lst = generate_past_4_week_days_full(monday_date_str)
                        
                        combined_data["past_4_week_car_cnt_mon_fri"]=0
                        combined_data["past_4_week_labour_hours_mon_fri"]=0
                        
                        combined_data["past_4_week_car_cnt_sat_sun"]=0
                        combined_data["past_4_week_labour_hours_sat_sun"]=0
                        for single_week in full_weeks_lst:
                            mon = single_week[0]
                            fri = single_week[1]
                            sat =single_week[2]
                            sun = single_week[3]
                            request_id4 = client.get_general_sales_report_request_id(reportOn,id,idname,mon, fri)
                            report_data4 = client.get_report(reportOn,request_id4)
                            extracted_data4 = report_data_extractor(report_data4)
                            
                            request_id4_2 =client.get_activity_by_date_proft_request_id(reportOn,mon, fri) #for labour hours
                            labour_hours_mon_fri=client.get_labour_hours(reportOn,request_id4_2)
                            
                            request_id5_2 =client.get_activity_by_date_proft_request_id(reportOn,sat, sun) #for labour hours
                            labour_hours_sat_sun=client.get_labour_hours(reportOn,request_id5_2)
                            
                            request_id5 = client.get_general_sales_report_request_id(reportOn,id,idname,sat, sun)
                            report_data5 = client.get_report(reportOn,request_id5)
                            extracted_data5 = report_data_extractor(report_data5)
                            
                            
                            car_cnt_mon_friday = extracted_data4.get("car_count",0)
                            car_cnt_sat_sun    = extracted_data5.get("car_count",0)
                            combined_data["past_4_week_car_cnt_mon_fri"] = combined_data.get("past_4_week_car_cnt_mon_fri",0)+car_cnt_mon_friday
                            combined_data["past_4_week_car_cnt_sat_sun"] = combined_data.get("past_4_week_car_cnt_sat_sun",0)+ car_cnt_sat_sun
                            combined_data["past_4_week_labour_hours_mon_fri"] = combined_data.get("past_4_week_labour_hours_mon_fri",0) + labour_hours_mon_fri
                            combined_data["past_4_week_labour_hours_sat_sun"] = combined_data.get("past_4_week_labour_hours_sat_sun",0) + labour_hours_sat_sun
                            
                            
                            
                            
                            
                            
                            
                        
                        combined_data["total_revenue"] = sum([total_revenue_val,total_revenue_val2])
                        combined_data["arm_plans_sold_cnt"] = arm_plans_sold_total_cnt
                        combined_data["total_arm_planmembers_cnt"] = total_arm_planmembers_cnt
                        combined_data["conversion_rate"]= conversion_rate
                        
                        combined_data["past_4_week_cnt"]=past_4_week_cnt #total car count past 4weeks 
                        combined_data["past_4_week_conversion_rate"] = past_4_weeks_conversion_rate
                        combined_data["past_4_weeks_total_revenue"] = past_4_weeks_total_revenue
                        combined_data["past_4_weeks_arm_plans_sold_cnt"]= past_4_week_arm_plans_sold
                        combined_data["past_4_weeks_retail_car_count"] = past_4_weeks_retail_car_count
                    print(f"combined data:{combined_data}")
                    site_watch_report[client_name]=combined_data
                    if combined_data:
                        is_location_code_success=True
                        success_location_code = locationCode,slno
                    if is_location_code_success:
                        break #break loop
                
                # if not is_location_code_success:#session not success
                else:
                    location_a_range =range(1,10)
                    location_b_range =range(10,14)
                    print("\n switching lcoation code")
                    if slno in location_a_range: 
                        locationCode=random.choice(location_codes_a)
                        print(f"\n retrying with new location code {locationCode} for {slno}")
                        success_location_code = locationCode,slno
                    
                    elif  slno in location_b_range:
                        locationCode=random.choice(location_codes_b)
                        success_location_code = locationCode,slno
                        print(f"\n retrying with new location code {locationCode} for {slno}")
                        
                    print("sleep for 5 secound before next retry")
                    time.sleep(15)
                    
                    
            except Exception as e:
                print(f"Excetion for this loctaion {client_name} {e}")                 
    
    return site_watch_report

#xl maps fucntions 

def do_sum(xl_map,start_index,range):
    "This will do row based some "
    total=0
    for i in range:
        val = xl_map[start_index][i+1]
        if isinstance(val,float) or isinstance(val,int):
            total+=val
    
    return total

def do_sum_location(xl_map,location:list):
    "This will take array of row ,col"
    total =0
    for row,col in location:
        val = xl_map[row][col]
        if isinstance(val,float) or isinstance(val,int):
            total+=val
    return total

#Xl maps fucntions ends

def Average_retail_visit__GA_SC_fucntion(retail_revenue_monday_to_friday_GA_SC,retail_revenue_saturday_to_sunday_GA_SC,
                                         retail_car_count_monday_to_friday_GA_SC,retail_car_count_saturday_to_sunday_GA_SC):    
    
    result = 0
    try:
        Average_retail_visit__GA_SC = sum([retail_revenue_monday_to_friday_GA_SC,retail_revenue_saturday_to_sunday_GA_SC])/sum(
        [retail_car_count_monday_to_friday_GA_SC,retail_car_count_saturday_to_sunday_GA_SC ]
    )
        result = Average_retail_visit__GA_SC
    except Exception as e:
        print(f"Exception in Average_retail_visit__GA_SC_fucntion() {e}")
        

    return result 

def Average_memeber_visit_GA_SC_function(Total_revenue_GA_SC,
        retail_revenue_monday_to_friday_GA_SC,retail_revenue_saturday_to_sunday_GA_SC,
        retail_car_count_monday_to_friday_GA_SC,retail_car_count_saturday_to_sunday_GA_SC,
        total_cars_GA_SC
        ):
    result = 0
    
    try:
        
        Average_memeber_visit_GA_SC = (Total_revenue_GA_SC - sum([retail_revenue_monday_to_friday_GA_SC , retail_revenue_saturday_to_sunday_GA_SC]))/(total_cars_GA_SC - sum([
        retail_car_count_monday_to_friday_GA_SC,retail_car_count_saturday_to_sunday_GA_SC]))
        
        result = Average_memeber_visit_GA_SC
        
    except Exception as e:
        print(f"Exception in Average_memeber_visit_GA_SC_function() {e}")
    
    return result

def cost_per_labour_hour_monday_to_friday_GA_SC_fucntion(car_count_monday_to_friday_GA_SC,staff_hours_monday_to_friday_GA_SC):
    result = 0
    try:
        cost_per_labour_hour_monday_to_friday_GA_SC = car_count_monday_to_friday_GA_SC/staff_hours_monday_to_friday_GA_SC
        result =cost_per_labour_hour_monday_to_friday_GA_SC
    except Exception as e:
        print(f"Exception in cost_per_labour_hour_monday_to_friday_GA_SC() {e}")

    return result


def cost_per_labour_hour_saturday_to_sunday_GA_SC_function(car_count_saturday_to_sunday_GA_SC,staff_hours_saturday_to_sunday_GA_SC):
    result = 0
    
    try:
        cost_per_labour_hour_saturday_to_sunday_GA_SC = car_count_saturday_to_sunday_GA_SC/staff_hours_saturday_to_sunday_GA_SC
        result = cost_per_labour_hour_saturday_to_sunday_GA_SC
    except Exception as e:
        print(f"Exception in staff_hours_saturday_to_sunday_GA_SC_fucntion() {e}")
        
    return result

def Total_cars_per_man_hour_GA_SC_function(total_cars_GA_SC, staff_hours_monday_to_friday_GA_SC,staff_hours_saturday_to_sunday_GA_SC):
    result =0 
    try:
            Total_cars_per_man_hour_GA_SC = total_cars_GA_SC/sum([
            staff_hours_monday_to_friday_GA_SC,staff_hours_saturday_to_sunday_GA_SC])
            result = Total_cars_per_man_hour_GA_SC
    
    except Exception as e:
        print(f"Exception in  Total_cars_per_man_hour_GA_SC_function() {e}")
    
    return result
 
def Conversion_rate_GA_SC_function(Total_club_plans_sold_GA_SC,retail_car_count_monday_to_friday_GA_SC,
                                   retail_car_count_saturday_to_sunday_GA_SC):
    result = 0 
    try:
        Conversion_rate_GA_SC = Total_club_plans_sold_GA_SC/sum([
        retail_car_count_monday_to_friday_GA_SC,retail_car_count_saturday_to_sunday_GA_SC])     
        
        result =  Conversion_rate_GA_SC
    except Exception as e:
        print(f"Exceptionin Conversion_rate_GA_SC_function() {e}")
    
    return result

def Conversion_rate_Total_function(Total_club_plans_sold_Total,
                                   retail_car_count_monday_to_friday_Total,retail_car_count_saturday_to_sunday_Total):
    result = 0
    
    try:
        
        Conversion_rate_Total = Total_club_plans_sold_Total/sum(
        [retail_car_count_monday_to_friday_Total,retail_car_count_saturday_to_sunday_Total])  
        
        result = Conversion_rate_Total
    except Exception as e:
        print(f"Exception in Conversion_rate_Total_function()  {e}") 
    
    return result

def Conversion_rate_ILL_function(Total_club_plans_sold_ILL,
                                 retail_car_count_monday_to_friday_ILL,retail_car_count_saturday_to_sunday_ILL):
    result = 0
    
    try:
        Conversion_rate_ILL = Total_club_plans_sold_ILL/sum([retail_car_count_monday_to_friday_ILL,retail_car_count_saturday_to_sunday_ILL])         
        result = Conversion_rate_ILL
    except Exception as e:
        print(f"Exception in Conversion_rate_ILL_function() {e}")
        
    return result

def Total_cars_per_man_hour_total_function(Total_cars_Total,
                                           staff_hours_monday_to_friday_Total,
                                           staff_hours_saturday_to_sunday_Total):
    result = 0
    
    try:
        Total_cars_per_man_hour_total = Total_cars_Total/sum(
        [staff_hours_monday_to_friday_Total,staff_hours_saturday_to_sunday_Total])
        
        result = Total_cars_per_man_hour_total
    
    except Exception as e:
        print(f"Exception in Total_cars_per_man_hour_total_function() {e}")
        
    return result


def Total_cars_per_man_hour_ILL_function(total_cars_in_ILL,
                                         staff_hours_monday_to_friday_ILL,
                                         staff_hours_saturday_to_sunday_ILL):
    result = 0
    
    try:
        Total_cars_per_man_hour_ILL = total_cars_in_ILL/(sum([
        staff_hours_monday_to_friday_ILL+staff_hours_saturday_to_sunday_ILL]))

        result = Total_cars_per_man_hour_ILL
    
    except Exception as e:
        print(f"Exception Total_cars_per_man_hour_ILL_function() {e}")
    
    return result


def cost_per_labour_hour_saturday_to_sunday_ILL_function(car_count_saturday_to_sunday_ILL,staff_hours_saturday_to_sunday_ILL):
    result = 0
    
    try:
        cost_per_labour_hour_saturday_to_sunday_ILL = car_count_saturday_to_sunday_ILL/staff_hours_saturday_to_sunday_ILL
        
        result = cost_per_labour_hour_saturday_to_sunday_ILL
    
    except Exception as e:
        print(f"Exception cost_per_labour_hour_saturday_to_sunday_ILL_function() {e}")
        
    return result


def cost_per_labour_hour_monday_to_friday_Total_function(car_count_monday_to_friday_Total,staff_hours_monday_to_friday_Total):
    result = 0
    
    try:
        cost_per_labour_hour_monday_to_friday_Total = car_count_monday_to_friday_Total/staff_hours_monday_to_friday_Total
        result = cost_per_labour_hour_monday_to_friday_Total
    
    except Exception as e:
        print(f"Exception cost_per_labour_hour_monday_to_friday_Total_function() {e}")
    
    return result

def cost_per_labour_hour_monday_to_friday_ILL_function(car_count_monday_to_friday_ILL,staff_hours_monday_to_friday_ILL):
    result = 0
    
    try:
        cost_per_labour_hour_monday_to_friday_ILL = car_count_monday_to_friday_ILL/staff_hours_monday_to_friday_ILL
        result =  cost_per_labour_hour_monday_to_friday_ILL
    
    except Exception as e:
        print(f"Exception cost_per_labour_hour_monday_to_friday_ILL_function() {e}")
        
    return result

def Average_memeber_visit_Total_function(Total_revenue_Total,retail_revenue_monday_to_friday_Total
                        ,retail_revenue_saturday_to_sunday_Total,Total_cars_Total,retail_car_count_monday_to_friday_Total,
                        retail_car_count_saturday_to_sunday_Total):
    result = 0
    
    try:
        Average_memeber_visit_Total = (Total_revenue_Total -sum([retail_revenue_monday_to_friday_Total,retail_revenue_saturday_to_sunday_Total]))/(Total_cars_Total - sum([
        retail_car_count_monday_to_friday_Total,retail_car_count_saturday_to_sunday_Total]))
        result =  Average_memeber_visit_Total
    
    except Exception as e:
        print(f"Exception Average_memeber_visit_Total_function() {e}")  
    
    return result    

def Average_memeber_visit_ILL_function(Total_revenue_ILL,retail_revenue_monday_to_friday_ILL,
                                       retail_revenue_saturday_to_sunday_ILL,total_cars_in_ILL,
                                       retail_car_count_monday_to_friday_ILL,retail_car_count_saturday_to_sunday_ILL):
    result = 0
    
    try:
        Average_memeber_visit_ILL = (Total_revenue_ILL - sum([retail_revenue_monday_to_friday_ILL,retail_revenue_saturday_to_sunday_ILL]))/(total_cars_in_ILL - sum([
        retail_car_count_monday_to_friday_ILL,retail_car_count_saturday_to_sunday_ILL]))
        
        result = Average_memeber_visit_ILL
    
    except Exception as e:
        print(f"Exception Average_memeber_visit_ILL_function() {e}")
    
    return result

def Average_retail_visit_Total_function(retail_revenue_monday_to_friday_Total,
                                        retail_revenue_saturday_to_sunday_Total,
                                        retail_car_count_monday_to_friday_Total,
                                        retail_car_count_saturday_to_sunday_Total):
    result = 0
    
    try:
        Average_retail_visit_Total = sum([retail_revenue_monday_to_friday_Total,retail_revenue_saturday_to_sunday_Total])/sum([
        retail_car_count_monday_to_friday_Total,retail_car_count_saturday_to_sunday_Total])
        
        result = Average_retail_visit_Total
    
    except Exception as e:
        print(f"Exception Average_retail_visit_Total_function() {e}")
        
    return result
            

def Average_retail_visit_IL_function(retail_revenue_monday_to_friday_ILL,
                                     retail_revenue_saturday_to_sunday_ILL,
                                     retail_car_count_monday_to_friday_ILL,
                                     retail_car_count_saturday_to_sunday_ILL):
    result = 0
    
    try:
        Average_retail_visit_IL = sum([retail_revenue_monday_to_friday_ILL,retail_revenue_saturday_to_sunday_ILL])/sum(
        [retail_car_count_monday_to_friday_ILL,retail_car_count_saturday_to_sunday_ILL])

        result = Average_retail_visit_IL
    
    except Exception as e:
        print(f"Exception Average_retail_visit_IL_function() {e}")
    
    return result



def update_place_to_xlmap(xl_map,place_index,place_dictionary)->list:
    "Will return updates place dictionary"
    xl_map[2][place_index] = place_dictionary.get("car_count_monday_to_friday")
    xl_map[3][place_index]=place_dictionary.get("car_count_saturday_sunday")
    xl_map[4][place_index]=place_dictionary.get("retail_car_count_monday_to_friday")
    xl_map[5][place_index]=place_dictionary.get("retail_car_count_saturday_sunday")
    
    xl_map[7][place_index]=place_dictionary.get("retail_revenue_monday_to_friday")
    xl_map[8][place_index]=place_dictionary.get("retail_revenue_saturday_sunday")
    xl_map[9][place_index]=place_dictionary.get("total_revenue_monday_to_friday")
    xl_map[10][place_index]=place_dictionary.get("total_revenue_saturday_sunday")
    
    xl_map[14][place_index]=place_dictionary.get("labour_hours_monday_to_friday")
    xl_map[15][place_index]=place_dictionary.get("labour_hours_saturday_sunday")
    xl_map[16][place_index]=place_dictionary.get("cars_per_labour_hour_monday_to_friday")
    xl_map[17][place_index]=place_dictionary.get("cars_per_labour_hour_saturday_sunday")
    
    xl_map[19][place_index] = place_dictionary.get("arm_plans_sold_cnt")
    xl_map[20][place_index]= place_dictionary.get("conversion_rate")
    xl_map[21][place_index] = place_dictionary.get("total_arm_planmembers_cnt")
        
    return xl_map

def prepare_xlmap(data,comment="The comment section",filename="test.xlsx",sheet_name="sheet1"):
    # Load the existing workbook using openpyxl
    try:
        workbook = openpyxl.load_workbook(filename)
        worksheet = workbook.create_sheet(sheet_name)
    except Exception as _:
        print(f"creating neww xl file !! {filename}")
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = sheet_name
    
    
    xl_map = [
    [""],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ["","","","","","","","","","","","","","","","","","","","","","","",],
    ]
    
    cols = ["    ","Totals","ILL",
                "GA / SC","Sudz - Beverly",'Fuller-Calumet',
                "Fuller-Cicero","Fuller-Matteson","Fuller-Elgin",
                "Splash-Peoria","Getaway-Macomb","Getaway-Morton",
                "Getaway-Ottawa","Getaway-Peru","Sparkle-Belair",
                "Sparkle-Evans","Sparkle-Furrys Ferry","Sparkle-Greenwood",
                "Sparkle-Grovetown 1","Sparkle-Grovetown 2","Sparkle-North Augusta",
                "Sparkle-Peach Orchard","Sparkle-Windsor Spring"]

    index_names =["Car Count Mon - Fri","Car Count Sat - Sun","Retail Car Count Mon - Fri",
                "Retail Car Count Sat - Sun","Total Cars","Retail Revenue Mon - Fri",
                "Retail Revenue Sat - Sun","Total Revenue Mon - Fri","Total Revenue Sat - Sun",
                "Total Revenue","Avg. Retail Visit","Avg. Member Visit",
                "Staff Hours Mon - Fri","Staff Hours Sat - Sun","Cars Per Labor Hour Mon - Fri",
                "Cars Per Labor Hour Sat & Sun","Total Cars Per Man Hour","Total Club Plans Sold",
                "Conversion Rate","Total Club Plan Members"]

    xl_map[0][0] = comment #to maps
    
    
    
    for col in range(len(cols)): #column names to maps
        xl_map[1][col]=cols[col]
        
    for index  in range(len(index_names)):  #index names to map
        xl_map[index+2][0]=index_names[index]
    
    Fuller_Cicero = data.get("Fuller-Cicero")  

    if Fuller_Cicero:
        Fuller_Cicero_index=6
        update_place_to_xlmap(xl_map,Fuller_Cicero_index,Fuller_Cicero)
    Fuller_Matteson = data.get("Fuller-Matteson")
    
    if Fuller_Matteson:
        Fuller_Matteson_index=7
        update_place_to_xlmap(xl_map,Fuller_Matteson_index,Fuller_Matteson)
    
    Fuller_Elgin = data.get("Fuller-Elgin")
    
    if Fuller_Elgin:
        Fuller_Elgin_index = 8
        update_place_to_xlmap(xl_map,Fuller_Elgin_index,Fuller_Elgin)
        
      
    Sparkle_Belair = data.get("Sparkle-Belair")
    
    if Sparkle_Belair:
        Sparkle_Belair_index=14

        update_place_to_xlmap(xl_map,Sparkle_Belair_index,Sparkle_Belair)
        
    Sparkle_Evans = data.get("Sparkle-Evans")
    
    if Sparkle_Evans:
        Sparkle_Evans_index=15

        update_place_to_xlmap(xl_map,Sparkle_Evans_index,Sparkle_Evans)

    Sparkle_North_Augusta = data.get("Sparkle-North Augusta")
    
    if Sparkle_North_Augusta:
        Sparkle_North_Augusta_index=20

        update_place_to_xlmap(xl_map,Sparkle_North_Augusta_index,Sparkle_North_Augusta)
    
    Sparkle_Greenwood = data.get("Sparkle-Greenwood")
    
    if Sparkle_Greenwood:
        Sparkle_Greenwood_index=17

        update_place_to_xlmap(xl_map,Sparkle_Greenwood_index,Sparkle_Greenwood)
    
    
    Sparkle_Grovetown_1 = data.get("Sparkle-Grovetown 1")
    
    if Sparkle_Grovetown_1:
        Sparkle_Grovetown_1_index=18

        update_place_to_xlmap(xl_map,Sparkle_Grovetown_1_index,Sparkle_Grovetown_1)
        
    Sparkle_Windsor_Spring = data.get("Sparkle-Windsor Spring")    
    
    if Sparkle_Windsor_Spring:
        Sparkle_Windsor_Spring_index=22 
        
        update_place_to_xlmap(xl_map,Sparkle_Windsor_Spring_index,Sparkle_Windsor_Spring)
        
    Sparkle_Furrys_Ferry = data.get("Sparkle-Furrys Ferry")
    
    if Sparkle_Furrys_Ferry:
        Sparkle_Furrys_Ferry_index= 16
        
        update_place_to_xlmap(xl_map,Sparkle_Furrys_Ferry_index,Sparkle_Furrys_Ferry)      

    Sparkle_Peach_Orchard = data.get("Sparkle-Peach Orchard")
    
    if Sparkle_Peach_Orchard:
        Sparkle_Peach_Orchard_index=21
        
        update_place_to_xlmap(xl_map,Sparkle_Peach_Orchard_index,Sparkle_Peach_Orchard)
        
    Sparkle_Grovetown_2 =  data.get("Sparkle-Grovetown 2")   
      
    if Sparkle_Grovetown_2:
        Sparkle_Grovetown_2_index =19
        
        update_place_to_xlmap(xl_map,Sparkle_Grovetown_2_index,Sparkle_Grovetown_2)
        
    Fuller_Calumet = data.get("Fuller-Calumet")
    
    if Fuller_Calumet:
        Fuller_Calumet_index= 5
        
        update_place_to_xlmap(xl_map,Fuller_Calumet_index,Fuller_Calumet)
        
        
    Sudz_Beverly = data.get("Sudz - Beverly")
    
    if Sudz_Beverly:
        Sudz_Beverly_index = 4
        
        update_place_to_xlmap(xl_map,Sudz_Beverly_index,Sudz_Beverly)
        
    ##Hamilton sites 
    
    Getaway_Macomb = data.get("Getaway-Macomb")
    
    if Getaway_Macomb:
        Getaway_Macomb_index = 10
        update_place_to_xlmap(xl_map,Getaway_Macomb_index,Getaway_Macomb)
        
    Getaway_Morton = data.get("Getaway-Morton")
    
    if Getaway_Morton:
        Getaway_Morton_index=11
        update_place_to_xlmap(xl_map,Getaway_Morton_index,Getaway_Morton)
        
    Getaway_Ottawa = data.get("Getaway-Ottawa")
    
    if Getaway_Ottawa:
        Getaway_Ottawa_index = 12
        update_place_to_xlmap(xl_map,Getaway_Ottawa_index,Getaway_Ottawa)
        
    Getaway_Peru = data.get("Getaway-Peru")    
    
    if Getaway_Peru:
        Getaway_Peru_index = 13
        update_place_to_xlmap(xl_map,Getaway_Peru_index,Getaway_Peru)
        
    
    #Hamilton 
    
    Splash_Peoria = data.get("Splash-Peoria")
    
    if Splash_Peoria:
        Splash_Peoria_index = 9
        update_place_to_xlmap(xl_map,Splash_Peoria_index,Splash_Peoria)

    #computation for first three columns 
    # sum([ xl_map[2][i+1] for i in range(3,13)])
    
    #first index
    car_count_monday_to_friday_ILL = do_sum(xl_map,2,range(3,13))
    xl_map[2][2] = car_count_monday_to_friday_ILL
    
    car_count_monday_to_friday_GA_SC = do_sum(xl_map,2,range(13,22))
    xl_map[2][3] = car_count_monday_to_friday_GA_SC
    
    car_count_monday_to_friday_Total = sum([car_count_monday_to_friday_ILL,car_count_monday_to_friday_GA_SC])
    
    xl_map[2][1]=car_count_monday_to_friday_Total
    
    
    
    #secound index 
    car_count_saturday_to_sunday_ILL = do_sum(xl_map,3,range(3,13))
    xl_map[3][2] = car_count_saturday_to_sunday_ILL
    
    car_count_saturday_to_sunday_GA_SC = do_sum(xl_map,3,range(13,22))
    xl_map[3][3] = car_count_saturday_to_sunday_GA_SC
    
    car_count_saturday_to_sunday_Total = sum([car_count_saturday_to_sunday_ILL, car_count_saturday_to_sunday_GA_SC])
    
    xl_map[3][1] = car_count_saturday_to_sunday_Total
    
    #third row 
    
    retail_car_count_monday_to_friday_ILL = do_sum(xl_map,4,range(3,13))
    xl_map[4][2] = retail_car_count_monday_to_friday_ILL
    
    retail_car_count_monday_to_friday_GA_SC = do_sum(xl_map,4,range(13,22))
    xl_map[4][3] = retail_car_count_monday_to_friday_GA_SC
    
    retail_car_count_monday_to_friday_Total = sum([retail_car_count_monday_to_friday_ILL,retail_car_count_monday_to_friday_GA_SC])
    
    xl_map[4][1]=retail_car_count_monday_to_friday_Total
    
    # Reatil car cound satruday to sunday
    
    retail_car_count_saturday_to_sunday_ILL = do_sum(xl_map,5,range(3,13))
    xl_map[5][2] = retail_car_count_saturday_to_sunday_ILL
    
    retail_car_count_saturday_to_sunday_GA_SC = do_sum(xl_map,5,range(13,22))
    xl_map[5][3] = retail_car_count_saturday_to_sunday_GA_SC
    
    retail_car_count_saturday_to_sunday_Total = sum([retail_car_count_saturday_to_sunday_ILL,retail_car_count_saturday_to_sunday_GA_SC])
    
    xl_map[5][1]=retail_car_count_saturday_to_sunday_Total
    
    loc_1 = [[2,3],[3,3]]
    total_cars_GA_SC=do_sum_location(xl_map,loc_1)
    
    xl_map[6][3]=total_cars_GA_SC
    
    loc_2=[[2,2],[3,2]]
    total_cars_in_ILL = do_sum_location(xl_map,loc_2)
    
    xl_map[6][2] = total_cars_in_ILL
    
    loc_3=[[6,3],[6,2]]
    
    Total_cars_Total = do_sum_location(xl_map,loc_3)
    xl_map[6][1]=Total_cars_Total
    
    #Retail Revenue Monday to Friday
    retail_revenue_monday_to_friday_ILL = do_sum(xl_map,7,range(3,13))
    xl_map[7][2] = retail_revenue_monday_to_friday_ILL
    
    retail_revenue_monday_to_friday_GA_SC = do_sum(xl_map,7,range(13,22))
    xl_map[7][3] = retail_revenue_monday_to_friday_GA_SC
    
    retail_revenue_monday_to_friday_Total = sum([retail_revenue_monday_to_friday_ILL,retail_revenue_monday_to_friday_GA_SC])
    
    xl_map[7][1]=retail_revenue_monday_to_friday_Total
    
    
    #Reatil Revenue Saturda to Sunday
    retail_revenue_saturday_to_sunday_ILL = do_sum(xl_map,8,range(3,13))
    xl_map[8][2] = retail_revenue_saturday_to_sunday_ILL
    
    retail_revenue_saturday_to_sunday_GA_SC = do_sum(xl_map,8,range(13,22))
    xl_map[8][3] = retail_revenue_saturday_to_sunday_GA_SC
    
    retail_revenue_saturday_to_sunday_Total = sum([retail_revenue_saturday_to_sunday_ILL,retail_revenue_saturday_to_sunday_GA_SC])
    
    xl_map[8][1]=retail_revenue_saturday_to_sunday_Total
    
    #Total Revnue Monday to Friday
    Total_revenue_monday_to_friday_ILL = do_sum(xl_map,9,range(3,13))
    xl_map[9][2] = Total_revenue_monday_to_friday_ILL
    
    Total_revenue_monday_to_friday_GA_SC = do_sum(xl_map,9,range(13,22))
    xl_map[9][3] = Total_revenue_monday_to_friday_GA_SC
    
    Total_revenue_monday_to_friday = sum([Total_revenue_monday_to_friday_ILL,Total_revenue_monday_to_friday_GA_SC])
    
    xl_map[9][1]=Total_revenue_monday_to_friday
    
    # Total Revenue Saturday to Sunday
    Total_revenue_saturday_to_sunday_ILL = do_sum(xl_map,10,range(3,13))
    xl_map[10][2] = Total_revenue_saturday_to_sunday_ILL
    
    Total_revenue_saturday_to_sunday_GA_SC = do_sum(xl_map,10,range(13,22))
    xl_map[10][3] = Total_revenue_saturday_to_sunday_GA_SC
    
    Total_revenue_saturday_sunday = sum([Total_revenue_saturday_to_sunday_ILL,Total_revenue_saturday_to_sunday_GA_SC])
    
    xl_map[10][1]=Total_revenue_saturday_sunday 
    
    # Total Revenue 
    loc_4=[[9,2],[10,2]]
    Total_revenue_ILL = do_sum_location(xl_map,loc_4)
    xl_map[11][2] =Total_revenue_ILL
    
    loc_5=[[9,3],[10,3]]
    Total_revenue_GA_SC = do_sum_location(xl_map,loc_5)
    xl_map[11][3]= Total_revenue_GA_SC
    
    Total_revenue_Total = sum([Total_revenue_ILL,Total_revenue_GA_SC])
    xl_map[11][1] =Total_revenue_Total
    
    #Average Retail Visit

    
    
    Average_retail_visit_IL_val =Average_retail_visit_IL_function(retail_revenue_monday_to_friday_ILL,
                                     retail_revenue_saturday_to_sunday_ILL,
                                     retail_car_count_monday_to_friday_ILL,
                                     retail_car_count_saturday_to_sunday_ILL)
    xl_map[12][2]=round(Average_retail_visit_IL_val,2)
    
    

    Average_retail_visit__GA_SC_val = Average_retail_visit__GA_SC_fucntion(retail_revenue_monday_to_friday_GA_SC
                                                                            ,retail_revenue_saturday_to_sunday_GA_SC,
                                                                            retail_car_count_monday_to_friday_GA_SC,
                                                                            retail_car_count_saturday_to_sunday_GA_SC)
    xl_map[12][3] = round(Average_retail_visit__GA_SC_val,2)
    
    Average_retail_visit_Total_val =Average_retail_visit_Total_function(retail_revenue_monday_to_friday_Total,
                                        retail_revenue_saturday_to_sunday_Total,
                                        retail_car_count_monday_to_friday_Total,
                                        retail_car_count_saturday_to_sunday_Total)
    xl_map[12][1] = round(Average_retail_visit_Total_val,2)
    
    #Average Member visit 
    
    Average_memeber_visit_ILL_val =Average_memeber_visit_ILL_function(Total_revenue_ILL,retail_revenue_monday_to_friday_ILL,
                                       retail_revenue_saturday_to_sunday_ILL,total_cars_in_ILL,
                                       retail_car_count_monday_to_friday_ILL,retail_car_count_saturday_to_sunday_ILL)
    xl_map[13][2] = round(Average_memeber_visit_ILL_val,2)
    
    Average_memeber_visit_GA_SC_val = Average_memeber_visit_GA_SC_function(Total_revenue_GA_SC,
        retail_revenue_monday_to_friday_GA_SC,retail_revenue_saturday_to_sunday_GA_SC,
        retail_car_count_monday_to_friday_GA_SC,retail_car_count_saturday_to_sunday_GA_SC,total_cars_GA_SC)
    
    xl_map[13][3] = round(Average_memeber_visit_GA_SC_val,2)
    
    Average_memeber_visit_Total_val =Average_memeber_visit_Total_function(Total_revenue_Total,retail_revenue_monday_to_friday_Total
                        ,retail_revenue_saturday_to_sunday_Total,Total_cars_Total,retail_car_count_monday_to_friday_Total,
                        retail_car_count_saturday_to_sunday_Total)
    
    xl_map[13][1] = round(Average_memeber_visit_Total_val,2)
    
    #Staff Hours Monday to Friday
    
    staff_hours_monday_to_friday_ILL = do_sum(xl_map,14,range(3,13))
    xl_map[14][2] = staff_hours_monday_to_friday_ILL
    
    staff_hours_monday_to_friday_GA_SC = do_sum(xl_map,14,range(13,22))
    
    xl_map[14][3] = staff_hours_monday_to_friday_GA_SC
    
    staff_hours_monday_to_friday_Total = sum([staff_hours_monday_to_friday_ILL,staff_hours_monday_to_friday_GA_SC])
    
    xl_map[14][1] = staff_hours_monday_to_friday_Total
    
    
    #Staff Hours Saturday to Sunday

    staff_hours_saturday_to_sunday_ILL = do_sum(xl_map,15,range(3,13))
    xl_map[15][2] = staff_hours_saturday_to_sunday_ILL
    
    staff_hours_saturday_to_sunday_GA_SC = do_sum(xl_map,15,range(13,22))
    
    xl_map[15][3] = staff_hours_saturday_to_sunday_GA_SC
    
    staff_hours_saturday_to_sunday_Total = sum([staff_hours_saturday_to_sunday_ILL,staff_hours_saturday_to_sunday_GA_SC])
    
    xl_map[15][1] = staff_hours_saturday_to_sunday_Total  
    
    # Cost per labour hour  Monday to friday
    cost_per_labour_hour_monday_to_friday_ILL_val =cost_per_labour_hour_monday_to_friday_ILL_function(car_count_monday_to_friday_ILL,staff_hours_monday_to_friday_ILL)
    xl_map[16][2] = round(cost_per_labour_hour_monday_to_friday_ILL_val,2)
    
    
    cost_per_labour_hour_monday_to_friday_GA_SC_val = cost_per_labour_hour_monday_to_friday_GA_SC_fucntion(car_count_monday_to_friday_GA_SC,staff_hours_monday_to_friday_GA_SC)
    xl_map[16][3] = round(cost_per_labour_hour_monday_to_friday_GA_SC_val,2)
    
    
    cost_per_labour_hour_monday_to_friday_Total_val =cost_per_labour_hour_monday_to_friday_Total_function(car_count_monday_to_friday_Total,staff_hours_monday_to_friday_Total)
    
    xl_map[16][1] = round(cost_per_labour_hour_monday_to_friday_Total_val,2)
    
    #Cost per laobour hour Saturday and Sunday
    cost_per_labour_hour_saturday_to_sunday_ILL_val = cost_per_labour_hour_saturday_to_sunday_ILL_function(car_count_saturday_to_sunday_ILL,staff_hours_saturday_to_sunday_ILL)
    xl_map[17][2] = round(cost_per_labour_hour_saturday_to_sunday_ILL_val,2)
    
    
    
    cost_per_labour_hour_saturday_to_sunday_GA_SC_val = cost_per_labour_hour_saturday_to_sunday_GA_SC_function(car_count_saturday_to_sunday_GA_SC,staff_hours_saturday_to_sunday_GA_SC)
    xl_map[17][3] = round(cost_per_labour_hour_saturday_to_sunday_GA_SC_val,2)
    
    cost_per_labour_hour_saturday_to_sunday_Total= car_count_saturday_to_sunday_Total/staff_hours_saturday_to_sunday_Total if staff_hours_saturday_to_sunday_Total!=0 else ""
    
    xl_map[17][1] = round(cost_per_labour_hour_saturday_to_sunday_Total,2) if cost_per_labour_hour_saturday_to_sunday_Total else ""
    
    # Total cars per man hour
    

    Total_cars_per_man_hour_ILL_val =Total_cars_per_man_hour_ILL_function(total_cars_in_ILL,
                                         staff_hours_monday_to_friday_ILL,
                                         staff_hours_saturday_to_sunday_ILL)
    xl_map[18][2] = round(Total_cars_per_man_hour_ILL_val,2)
    
    Total_cars_per_man_hour_GA_SC_val = Total_cars_per_man_hour_GA_SC_function(total_cars_GA_SC, staff_hours_monday_to_friday_GA_SC,
                                                                               staff_hours_saturday_to_sunday_GA_SC)
    
    xl_map[18][3] = round(Total_cars_per_man_hour_GA_SC_val,2)
    
    
    Total_cars_per_man_hour_total_val =Total_cars_per_man_hour_total_function(Total_cars_Total,
                                           staff_hours_monday_to_friday_Total,
                                           staff_hours_saturday_to_sunday_Total)
    xl_map[18][1] = round(Total_cars_per_man_hour_total_val,2)
    
    #Total club plans sold 
    Total_club_plans_sold_ILL = do_sum(xl_map,19,range(3,13))
    
    xl_map[19][2] = Total_club_plans_sold_ILL
    
    Total_club_plans_sold_GA_SC = do_sum(xl_map,19,range(13,22))
    
    xl_map[19][3] = Total_club_plans_sold_GA_SC
    
    Total_club_plans_sold_Total = sum([Total_club_plans_sold_ILL,Total_club_plans_sold_GA_SC])
    
    xl_map[19][1] = Total_club_plans_sold_Total
    
    
    #Conversion Rate 
    Conversion_rate_ILL_val =Conversion_rate_ILL_function(Total_club_plans_sold_ILL,
                                                          retail_car_count_monday_to_friday_ILL,
                                                          retail_car_count_saturday_to_sunday_ILL)
    xl_map[20][2] = round((Conversion_rate_ILL_val * 100),2)
    
    Conversion_rate_GA_SC_val =Conversion_rate_GA_SC_function(Total_club_plans_sold_GA_SC,retail_car_count_monday_to_friday_GA_SC,
                                   retail_car_count_saturday_to_sunday_GA_SC)
    
    xl_map[20][3]= round((Conversion_rate_GA_SC_val * 100),2)
    
    

    Conversion_rate_Total_val = Conversion_rate_Total_function(Total_club_plans_sold_Total,
                                                               retail_car_count_monday_to_friday_Total,retail_car_count_saturday_to_sunday_Total)
    xl_map[20][1] = round((Conversion_rate_Total_val * 100),2)
    
    
    #Total club plan members 
    Total_club_planmembers_ILL = do_sum(xl_map,21,range(3,13))
    
    xl_map[21][2] = Total_club_planmembers_ILL
    
    Total_club_planmembers_GA_SC = do_sum(xl_map,21,range(13,22))
    
    xl_map[21][3] = Total_club_planmembers_GA_SC 
    
    Total_club_planmembers_Total = sum([Total_club_planmembers_ILL,Total_club_planmembers_GA_SC])
    
    xl_map[21][1] = Total_club_planmembers_Total
    
    #Toatl Cars for all locations 
    total_cars_row = 6 
    for i in range(3,22):
        xl_map[total_cars_row][i+1]= do_sum_location(xl_map,[[2,i+1],[3,i+1]])
    
    
    revenue_row = 11
    
    #Total Revenue calculatio n for all locations 
    for i in range(3,22):
        xl_map[revenue_row][i+1]= do_sum_location(xl_map,[[9,i+1],[10,i+1]])
    
    average_retail_visit_row = 12
    for i in range(3,22):
        retail_revenue_mon_sun = do_sum_location(xl_map,location=[[7,i+1],[8,i+1]])
        retail_car_count_mon_sun = do_sum_location(xl_map,location=[[4,i+1],[5,i+1]])
        
        average_retail_visit_val = retail_revenue_mon_sun/retail_car_count_mon_sun if retail_car_count_mon_sun != 0 else ""
        xl_map[average_retail_visit_row][i+1]=   round(average_retail_visit_val,2) if average_retail_visit_val else ""
    
    
    average_member_visit_row = 13
    for i in range(3,22):
        total_revenue = xl_map[11][i+1]
        total_revenue = total_revenue if isinstance(total_revenue,int) or isinstance(total_revenue,float) else 0
        
        retail_revenue_mon_sun = do_sum_location(xl_map,location=[[7,i+1],[8,i+1]])
        
        total_cars = xl_map[6][i+1]
        total_cars = total_cars if isinstance(total_cars,int) or isinstance(total_cars,float) else 0
        
        retail_car_count_mon_sun  = do_sum_location(xl_map,location=[[4,i+1],[5,i+1]])
        
        average_member_visit_val = (total_revenue-retail_revenue_mon_sun)/(total_cars-retail_car_count_mon_sun) if total_cars-retail_car_count_mon_sun !=0 else ""
        
        xl_map[average_member_visit_row][i+1] = round(average_member_visit_val,2) if average_member_visit_val else ""
    
    #Total cars per man hour 
    total_cars_per_man_hour_row = 18 
    for i in range(3,22):
        total_cars = xl_map[6][i+1]
        total_cars = total_cars if isinstance(total_cars,int) or isinstance(total_cars,float) else 0
        staff_hours_monday_to_saturday = do_sum_location(xl_map,[[14,i+1],[15,i+1]])
        
        total_cars_per_man_hour_val = total_cars/staff_hours_monday_to_saturday if staff_hours_monday_to_saturday !=0 else ""
        
        xl_map[total_cars_per_man_hour_row][i+1] = round(total_cars_per_man_hour_val,2) if total_cars_per_man_hour_val else ""
 
        
      
    # Define cell styles (assuming they are the same as before)
    bg_color = PatternFill(start_color='0b3040', end_color='0b3040', fill_type='solid')
    font_color = Font(color='FFFFFF')

    bg_color_index = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    font_color_index = Font(color='000000')

    darkgreen_format = PatternFill(start_color='0ee85e', end_color='0ee85e', fill_type='solid')
    light_green_format = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
    darkred_format = PatternFill(start_color='FC0303', end_color='FC0303', fill_type='solid')
    lightred_format = PatternFill(start_color='D98484', end_color='D98484', fill_type='solid')
    yellow=  PatternFill(start_color='D0D48A', end_color='D0D48A', fill_type='solid')
    
    #writing to  actual sheet
    #first row comment section
    # Write comments in the first row
    # worksheet.append([comment])
    
    for row in range(len(xl_map)):
        for col in range(len(xl_map[row])):
            val = xl_map[row][col]
            cell = worksheet.cell(row=row+1, column=col+1, value=val)  # offset by 0 rows for header and comment
            
            # val = add_commas(val)
            # val = float(val)
            # print("type:",type(val),val)
            if row == 1 and col != 0:
                cell.fill = bg_color
                cell.font = font_color

            elif col == 0 and 1 < row < 22:
                cell.fill = bg_color_index
                cell.font = font_color_index

            elif val  and row == 12 and col > 0:
                if val >= 10:
                    cell.fill = darkgreen_format
                    
                elif val>=5 and val <10: # [5,9] (inclusive intervals)
                    cell.fill = light_green_format
                    
                elif val>=-5 and val <5:  # [-5,4]
                    cell.fill = lightred_format
        
                elif  val <=-10 and val<-5:
                    cell.fill = darkred_format

            elif val  and row in [16, 17, 18, 20] and col > 0:
                if val >= 20:
                    cell.fill = darkgreen_format
                elif val>=10 and val <20:
                    cell.fill = light_green_format
                elif val>=-10 and val <10:
                    cell.fill = lightred_format
                elif val <=-20 or val<-10:
                    cell.fill = darkred_format

            # elif row == 5 and col > 2:
            #     cell.value = ""

            # elif row in [17] and col > 2:
            #     cell.value = ""

            # elif row == 21:
            #     cell.value = ""

            # elif row > 21 and col > 0:
            #     cell.value = ""

    # Add legend or additional information below the table
    legend_start_row = 24

    legend_styles = {
        "Very Concerning": PatternFill(start_color="fc0303", end_color="fc0303", fill_type="solid"),
        "Concerning": PatternFill(start_color="d98484", end_color="d98484", fill_type="solid"),
        "Neutral": PatternFill(start_color="d0d48a", end_color="d0d48a", fill_type="solid"),
        "Positive": PatternFill(start_color="8ad493", end_color="8ad493", fill_type="solid"),
        "Very Positive": PatternFill(start_color="0ee85e", end_color="0ee85e", fill_type="solid"),
    }

    worksheet.cell(row=legend_start_row, column=1, value='Legend')
    worksheet.cell(row=legend_start_row + 1, column=1, value='Very Concerning').fill = legend_styles["Very Concerning"]
    worksheet.cell(row=legend_start_row + 2, column=1, value='Concerning').fill = legend_styles["Concerning"]
    worksheet.cell(row=legend_start_row + 3, column=1, value='Neutral').fill = legend_styles["Neutral"]
    worksheet.cell(row=legend_start_row + 4, column=1, value='Positive').fill = legend_styles["Positive"]
    worksheet.cell(row=legend_start_row + 5, column=1, value='Very Positive').fill = legend_styles["Very Positive"]
    
    #setting cloumn width as deafult 
    column_width = 20  # You can change this value to whatever width you need
    first_col_width =25
    for col in range(1, 24):  # Columns A to W are 1 to 23
        column_letter = get_column_letter(col)
        if col==1:
            worksheet.column_dimensions[column_letter].width = first_col_width
        else:
            worksheet.column_dimensions[column_letter].width = column_width
       
    
    # Define the border style


    thick_border = Border(
    left=Side(style='thick'),
    right=Side(style='thick'))
    
    thick_border_bottom = Border(
    left=Side(style='thick'),
    right=Side(style='thick'),
    bottom=Side(style='thick'))




    # Apply the border to a range of cells (e.g., A1:C3)
    # for row in worksheet.iter_rows(min_row=3, max_row=21, min_col=2, max_col=4):
    #     for cell in row:
    #         cell.border = thin_border

    for row in range(3,23):
        cell0 = worksheet.cell(row=row,column=1)
        cell1 = worksheet.cell(row=row,column=2)
        cell2 = worksheet.cell(row=row,column=3)
        cell3 = worksheet.cell(row=row,column=4)
        
        if row in [7,12,13,19,21,22]:
            cell0.border=thick_border_bottom
            cell1.border=thick_border_bottom
            cell2.border=thick_border_bottom
            cell3.border=thick_border_bottom
        else:
            cell0.border=thick_border
            cell1.border=thick_border
            cell2.border=thick_border
            cell3.border=thick_border
            
    for row in worksheet.iter_rows():
        for cell in row:
            row_index = cell.row
            if isinstance(cell.value, (int, float)) and row_index in [17,18,19]:
                cell.number_format = '#,##0.0'
            elif isinstance(cell.value, (int, float)) and row_index in [13,14,21]:
                cell.number_format = '#,##0.00'
            elif isinstance(cell.value, (int, float)): #:
                cell.number_format = '#,##0'
    #Doller sysmbol     
    for row in range(8,13):
        cell1 = worksheet.cell(row=row,column=2)
        cell2 = worksheet.cell(row=row,column=3)
        cell3 = worksheet.cell(row=row,column=4)
        
        # cell1.border=thick_border
        # cell2.border=thick_border
        # cell3.border=thick_border
        
        cells=[cell1,cell2,cell3]
        for cell in cells:
            if isinstance(cell.value, (int, float)) and cell.value >= 1000:
                cell.number_format = '"$"#,##0'     
    
    #finding past week averages for Total car count
    all_locations = [Sudz_Beverly,Fuller_Calumet,
                    Fuller_Cicero,Fuller_Matteson,
                    Fuller_Elgin,Splash_Peoria,Getaway_Macomb,Getaway_Morton,
                    Getaway_Ottawa,Getaway_Peru,Sparkle_Belair,
                    Sparkle_Evans,Sparkle_Furrys_Ferry,Sparkle_Greenwood,
                    Sparkle_Grovetown_1,Sparkle_Grovetown_2,Sparkle_North_Augusta,
                    Sparkle_Peach_Orchard,Sparkle_Windsor_Spring]
    
    
    lst_main = ["Totals","ILL",
                "GA / SC"]
    ga_sc = all_locations[10:]
    ill = all_locations[0:10]
    ill_past_4_weeks_car_cnt = [loc_data.get("past_4_week_cnt") for loc_data in ill if loc_data]
    ill_sum_past = sum(ill_past_4_weeks_car_cnt)/4
    current_ill_week_cnt = xl_map[6][2]
    ill_average_percent = ((current_ill_week_cnt - ill_sum_past)/ ill_sum_past)*100
    
    ga_sc_past_4_weeks_car_cnt = [loc_data.get("past_4_week_cnt") for loc_data in ga_sc if loc_data]
    ga_sc_sum_past = sum(ga_sc_past_4_weeks_car_cnt)/4
    current_ga_sc_week_cnt = xl_map[6][3]
    ga_sc_average_percent= ((current_ga_sc_week_cnt-ga_sc_sum_past)/ga_sc_sum_past)*100
    
    totals_past = ill_sum_past +ga_sc_sum_past
    totals_current = xl_map[6][1]
    total_average_percent = ((totals_current-totals_past)/totals_past)*100
    
    ill_total_revenue_past_4_weeks = [loc_data.get("past_4_weeks_total_revenue") for loc_data in ill if loc_data]
    ill_total_revenue_avg = sum(ill_total_revenue_past_4_weeks)/4
    ill_curent_revenue = xl_map[11][2]
    ill_avg_revenue_change  = ((ill_curent_revenue-ill_total_revenue_avg)/ill_total_revenue_avg)*100
    
    ga_sc_total_revenue_past_4_weeks = [loc_data.get("past_4_weeks_total_revenue") for loc_data in ga_sc if loc_data]
    ga_sc_avg_revenue = sum(ga_sc_total_revenue_past_4_weeks)/4
    ga_sc_curent_revenue  = xl_map[11][3]
    
    ga_sc_avg_revenue_change = ((ga_sc_curent_revenue - ga_sc_avg_revenue)/ga_sc_avg_revenue)*100
    
    total_revenue_past_4_avg_total = ill_total_revenue_avg + ga_sc_avg_revenue
    total_reveneu_curent = ill_curent_revenue + ga_sc_curent_revenue
    
    total_revenue_total_change = ((total_reveneu_curent - total_revenue_past_4_avg_total)/total_revenue_past_4_avg_total)*100
    
    
    ill_past_4_weeks_arm_plan_sold = [loc_data.get("past_4_weeks_arm_plans_sold_cnt") for loc_data in ill if loc_data]
    ill_past_4_weeks_arm_plan_sold_sum = sum(ill_past_4_weeks_arm_plan_sold)
    
    ill_past_4_retail_car_count = [loc_data.get("past_4_weeks_retail_car_count") for loc_data in ill if loc_data]
    ill_past_4_retail_car_count_sum = sum(ill_past_4_retail_car_count)
    
    ill_past_4_conversation_rate = (ill_past_4_weeks_arm_plan_sold_sum/ill_past_4_retail_car_count_sum)*100
    
    ill_current_conversation_rate = xl_map[20][2]
    ill_conversation_rate_change = ill_current_conversation_rate - ill_past_4_conversation_rate
    
    ga_sc_past_4_weeks_arm_plans_sold = [loc_data.get("past_4_weeks_arm_plans_sold_cnt") for loc_data in ga_sc if loc_data]
    ga_sc_past_4_weeks_arm_plans_sold_sum = sum(ga_sc_past_4_weeks_arm_plans_sold)
    
    ga_sc_past_4_weeks_retail_car_count = [loc_data.get("past_4_weeks_retail_car_count") for loc_data in ga_sc if loc_data]
    ga_sc_past_4_weeks_retail_car_count_sum = sum(ga_sc_past_4_weeks_retail_car_count)
    
    ga_sc_past_4_conversation_rate = ( ga_sc_past_4_weeks_arm_plans_sold_sum/ga_sc_past_4_weeks_retail_car_count_sum)*100
    
    ga_sc_current_conversation_rate  = xl_map[20][3]
    ga_sc_conversation_change = ga_sc_current_conversation_rate - ga_sc_past_4_conversation_rate 
    
    past_total_arm_plans_sold = ill_past_4_weeks_arm_plan_sold_sum + ga_sc_past_4_weeks_arm_plans_sold_sum
    
    past_total_reatil_car_count = ill_past_4_retail_car_count_sum + ga_sc_past_4_weeks_retail_car_count_sum
    
    past_total_conversation_change = (past_total_arm_plans_sold/past_total_reatil_car_count)*100
    
    current_total_conversation_rate = xl_map[20][1]
    
    total_conversation_change = current_total_conversation_rate - past_total_conversation_change
    
    
    
    
    colours = darkgreen_format,light_green_format,darkred_format,lightred_format
    print(f"ill avg : {ill_average_percent}")
    print("ill avg revenue:",ill_avg_revenue_change)
    print("ga sc avg revenue :",ga_sc_avg_revenue_change)
    print("total revenue total change:",total_revenue_total_change)
    print(f"ga_sc average :{ga_sc_average_percent}")
    print(f"total_average : {total_average_percent}")
    print("ill conversation change :",ill_conversation_rate_change)
    print("gasc conversation chane :",ga_sc_conversation_change)
    print("total conversation change :",total_conversation_change)
    set_colour(ill_average_percent,7,3,worksheet,colours) #for ill
    set_colour(ga_sc_average_percent,7,4,worksheet,colours) #for gasc total
    set_colour(total_average_percent,7,2,worksheet,colours) #for total total
    
    set_colour(ill_avg_revenue_change,12,3,worksheet,colours) #ill 
    set_colour(ga_sc_avg_revenue_change,12,4,worksheet,colours) #ga sc
    set_colour(total_revenue_total_change,12,2,worksheet,colours)
    set_colour(ill_conversation_rate_change,21,3,worksheet,colours)
    set_colour(ga_sc_conversation_change,21,4,worksheet,colours)
    set_colour(total_conversation_change,21,2,worksheet,colours)
    
    
    
    loc_names = ["Sudz - Beverly",'Fuller-Calumet',
                "Fuller-Cicero","Fuller-Matteson","Fuller-Elgin",
                "Splash-Peoria","Getaway-Macomb","Getaway-Morton",
                "Getaway-Ottawa","Getaway-Peru","Sparkle-Belair",
                "Sparkle-Evans","Sparkle-Furrys Ferry","Sparkle-Greenwood",
                "Sparkle-Grovetown 1","Sparkle-Grovetown 2","Sparkle-North Augusta",
                "Sparkle-Peach Orchard","Sparkle-Windsor Spring"]
    
    for index,place_dictionary in enumerate(all_locations):
        current_week_total_cars = xl_map[6][index+4]
        past_4_week_total_cars = place_dictionary.get("past_4_week_cnt")
        change_in_total_car_count_percent  = chnage_total_car_count_fun(current_week_total_cars,past_4_week_total_cars)
        set_colour(change_in_total_car_count_percent,7,index+5,worksheet,colours) #for total cars 
        
        print(f"{loc_names[index]}=>chnage car count  {change_in_total_car_count_percent}")
        change_in_conversationrate = place_dictionary.get("conversion_rate") - place_dictionary.get("past_4_week_conversion_rate")
        set_colour(change_in_conversationrate,21,index+5,worksheet,colours) #conversation rate colours
        
        print(f"{loc_names[index]}=>chnage conversation rate   {change_in_conversationrate}")
        current_revenue_total = place_dictionary.get("total_revenue")
        past_4_week_revenue_total = place_dictionary.get("past_4_weeks_total_revenue")
        change_in_total_revenue = chnage_total_revenue_fun(current_revenue_total,past_4_week_revenue_total)
        print(f"{loc_names[index]}=>chnage total revenue    {change_in_total_revenue}")
        set_colour(change_in_total_revenue,12,index+5,worksheet,colours)
        
        print("\n"*2)
        
    
    

    #applying bold font
    # Define a bold font style
    bold_font = Font(bold=True)
    
    for row in worksheet.iter_rows(min_row=7, max_row=7, min_col=1, max_col=4):
        for cell in row:
            cell.font = bold_font

    for row in worksheet.iter_rows(min_row=12, max_row=13, min_col=1, max_col=4):
        for cell in row:
            cell.font = bold_font
            
    for row in worksheet.iter_rows(min_row=19, max_row=19, min_col=1, max_col=4):
        for cell in row:
            cell.font = bold_font
            
    for row in worksheet.iter_rows(min_row=21, max_row=21, min_col=1, max_col=4):
        for cell in row:
            cell.font = bold_font
    # Save the modified workbook
    workbook.save(filename)


if __name__=="__main__":
    # import pandas as pd
    # monday_date_str, sunday_date_str = get_week_dates()
    # print(monday_date_str,sunday_date_str)
    monday_date_str="2024-06-10"
    friday_date_str = "2024-06-14"
    saturday_date_str = "2024-06-15"
    sunday_date_str="2024-06-16"  #YMD
    
    monday_date_str="2024-07-01"
    friday_date_str = "2024-07-05"
    saturday_date_str = "2024-07-06"
    sunday_date_str="2024-07-07"  #Y-M-D
    
    report = generate_weekly_report("",monday_date_str,friday_date_str,saturday_date_str, sunday_date_str)
    print("\n"*6)
    print(report)
    with open("sitewatch_report.json","w") as f:
        json.dump(report,f,indent=4)
    # with open("SPKLUS-002.json",'r') as f:
    #     data = json.load(f)
        
    # df = pd.read_json("SPKLUS-002.json")
    
    # df.to_excel("sitewatch.xlsx")
    
    # with open("sitewatch_report_old.json",'r') as f:
    #     data=json.load(f)
    # data = report
    #prepare_xlmap(data)
    