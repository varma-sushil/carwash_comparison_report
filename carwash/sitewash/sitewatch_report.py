import os
import json
import time
import traceback
import pandas as pd
import random
import logging
import sys

from sitewatch4 import sitewatchClient

# Add the carwash directory to the sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))


MAX_RETRIES = 100

current_folder_path = os.path.dirname(os.path.abspath(__file__))
cookies_path = os.path.join(current_folder_path, "cookies")
xlfile_path = os.path.join(current_folder_path, "sitewash_data.xlsx")
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
                'Grovetown 2', 'Cicero', 'Matteson', 'Sparkle Express',
                "Fuller's Calumet City "]


def get_total_arm_plan_members(client, reportOn, end_date):
    total_arm_planmembers_cnt = None
    while True:
        total_members_req_id = client.get_plan_analysis_request_id(
                                        end_date, reportOn)
        total_arm_planmembers_cnt = client.get_total_plan_members(
                                        total_members_req_id, reportOn)

        if total_arm_planmembers_cnt or total_arm_planmembers_cnt == 0:
            break

        print("retrying for report data")
        time.sleep(5)

    return total_arm_planmembers_cnt


def wash_sales(section):

    subtotals = section.get("subtotals")[0]

    return subtotals.get("quantity")


def wash_packages(section):
    wash_packages_lst = []

    reports = section.get("reports")
    subtotals = section.get("subtotals")

    for report in reports:
        wash_package_structure = {
            "Wash_packages_Description": report.get("description"),
            "Wash_packages_price": report.get("price"),
            "Wash_packages_quantity": report.get("quantity"),
            "Wash_packages_amount": report.get("amount"),
        }
        wash_packages_lst.append(wash_package_structure)

    for subtotal in subtotals:
        wash_package_structure = {
            "Wash_packages_Description": subtotal.get("description"),
            "Wash_packages_price": subtotal.get("price"),
            "Wash_packages_quantity": subtotal.get("quantity"),
            "Wash_packages_amount": subtotal.get("amount"),
        }
        wash_packages_lst.append(wash_package_structure)

    return wash_packages_lst


def wash_extra_services(section):
    wash_extra_service_lst = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")

    for report in reports:
        wash_extra_structure = {
            "Wash_Extra_Services_Description": report.get("description"),
            "Wash_Extra_Services_price": report.get("price"),
            "Wash_Extra_Services_quantity": report.get("quantity"),
            "Wash_Extra_Services_amout": report.get("amount")
        }
        wash_extra_service_lst.append(wash_extra_structure)

    for total in subtotals:
        wash_extra_structure = {
            "Wash_Extra_Services_Description": total.get("description"),
            "Wash_Extra_Services_price": total.get("price"),
            "Wash_Extra_Services_quantity": total.get("quantity"),
            "Wash_Extra_Services_amout": total.get("amount")
        }
        wash_extra_service_lst.append(wash_extra_structure)

    return wash_extra_service_lst


def gross_wash_sales(section):
    gross_wash_sales_lst = []
    gross_wash_sale_structure = {
            "Gross_Wash_Sales": section.get("totalAmount")
        }

    gross_wash_sales_lst.append(gross_wash_sale_structure)

    return gross_wash_sales_lst


def less_free_wash_rdmd(section):
    less_wash_sales_rdmd_lst = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")

    for report in reports:
        less_free_wash_rmd_structure = {
            "Less_free_wash_rdmd_Description": report.get("description"),
            "Less_free_wash_rdmd_price": report.get("price"),
            "Less_free_wash_rdmd_quantity": report.get("quantity"),
            "Less_free_wash_rdmd_amount": report.get("amount"),
        }
        less_wash_sales_rdmd_lst.append(less_free_wash_rmd_structure)

    for subtotal in subtotals:
        less_free_wash_rmd_structure = {
            "Less_free_wash_rdmd_Description": subtotal.get("description"),
            "Less_free_wash_rdmd_price": subtotal.get("price"),
            "Less_free_wash_rdmd_quantity": subtotal.get("quantity"),
            "Less_free_wash_rdmd_amount": subtotal.get("amount"),
        }
        less_wash_sales_rdmd_lst.append(less_free_wash_rmd_structure)
    return less_wash_sales_rdmd_lst


def less_wash_discounts(section):
    less_wash_discounts_lst = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    for report in reports:
        less_wash_discount_structure = {
            "Less_Wash_Discounts_Description": report.get("description"),
            "Less_Wash_Discounts_Price": report.get("price"),
            "Less_Wash_Discounts_quantity": report.get("quantity"),
            "Less_Wash_Discounts_amount": report.get("amount")
        }
        less_wash_discounts_lst.append(less_wash_discount_structure)

    for subtotal in subtotals:
        less_wash_discount_structure = {
            "Less_Wash_Discounts_Description": subtotal.get("description"),
            "Less_Wash_Discounts_Price": subtotal.get("price"),
            "Less_Wash_Discounts_quantity": subtotal.get("quantity"),
            "Less_Wash_Discounts_amount": subtotal.get("amount")
        }
        less_wash_discounts_lst.append(less_wash_discount_structure)

    return less_wash_discounts_lst


def less_loyality_disc(section):
    less_loyality_disc_lst = []

    reports = section.get("reports")
    subtotals = section.get("subtotals")

    for report in reports:
        less_loyality_disc_structure = {
            "Less_Loyalty_disc_description": report.get("description"),
            "Less_Loyalty_disc_price": report.get("price"),
            "Less_Loyalty_disc_quantity": report.get("quantity"),
            "Less_Loyalty_disc_amount": report.get("amount")
        }
        less_loyality_disc_lst.append(less_loyality_disc_structure)

    for subtotal in subtotals:
        less_loyality_disc_structure = {
            "Less_Loyalty_disc_description": subtotal.get("description"),
            "Less_Loyalty_disc_price": subtotal.get("price"),
            "Less_Loyalty_disc_quantity": subtotal.get("quantity"),
            "Less_Loyalty_disc_amount": subtotal.get("amount")
        }
        less_loyality_disc_lst.append(less_loyality_disc_structure)

    return less_loyality_disc_lst


def net_site_sales(section):

    return section.get("totalAmount")


def arm_plans_sold(section):

    return section.get("totalQuantity", 0.0)


def arm_plans_recharged(section):
    data = None
    arm_plans_recharged_lst = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")

    data = subtotals[0].get("amount")

    return data


def arm_planes_reedemed(section):
    data = {}

    data["arm_plans_reedemed_cnt"] = section.get("totalQuantity")
    data["arm_plans_reedemed_amt"] = section.get("totalAmount", 0.0)*(-1)
    return data


def arm_plans_terminated(section):
    arm_plans_terminated_lst = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    for report in reports:
        arm_plan_terminated_structure = {
            "Arm_plans_terminated_description": report.get("description"),
            "Arm_plans_terminated_price": report.get("price"),
            "Arm_plans_terminated_quantity": report.get("quantity"),
            "Arm_plans_terminated_amount": report.get("amount")
        }
        arm_plans_terminated_lst.append(arm_plan_terminated_structure)
    for subtotal in subtotals:
        arm_plan_terminated_structure ={
            "Arm_plans_terminated_description": subtotal.get("description"),
            "Arm_plans_terminated_price": subtotal.get("price"),
            "Arm_plans_terminated_quantity": subtotal.get("quantity"),
            "Arm_plans_terminated_amount": subtotal.get("amount")
        }
        arm_plans_terminated_lst.append(arm_plan_terminated_structure)

    return arm_plans_terminated_lst


def prepaid_sold(section):
    prepaid_sold_lst = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")

    for report in reports:
        prepaid_sold_structure = {
            "Prepaids_Sold_description": report.get("description"),
            "Prepaids_Sold_price": report.get("price"),
            "Prepaids_Sold_quantity": report.get("quantity"),
            "Prepaids_Sold_amount": report.get("amount"),
        }
        prepaid_sold_lst.append(prepaid_sold_structure)

    for subtotal in subtotals:
        prepaid_sold_structure = {
            "Prepaids_Sold_description": subtotal.get("description"),
            "Prepaids_Sold_price": subtotal.get("price"),
            "Prepaids_Sold_quantity": subtotal.get("quantity"),
            "Prepaids_Sold_amount": subtotal.get("amount"),
        }
        prepaid_sold_lst.append(prepaid_sold_structure)

    return prepaid_sold_lst


def less_prepaid_reedemed(section):
    less_prepaid_reedemed_lst = []

    reports = section.get("reports")
    subtotals = section.get("subtotals")
    for report in reports:
        less_prepaid__reedemed_structure = {
            "Less_prepaids_redeemed_description": report.get("description"),
            "Less_prepaids_redeemed_price": report.get("price"),
            "Less_prepaids_redeemed_quantity": report.get("quantity"),
            "Less_prepaids_redeemed_amount": report.get("amount")
        }

        less_prepaid_reedemed_lst.append(less_prepaid__reedemed_structure)

    for subtotal in subtotals:
        less_prepaid__reedemed_structure = {
            "Less_prepaids_redeemed_description": subtotal.get("description"),
            "Less_prepaids_redeemed_price": subtotal.get("price"),
            "Less_prepaids_redeemed_quantity": subtotal.get("quantity"),
            "Less_prepaids_redeemed_amount": subtotal.get("amount")
        }

        less_prepaid_reedemed_lst.append(less_prepaid__reedemed_structure)

    return less_prepaid_reedemed_lst


def online_sold(section):
    data = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")

    for report in reports:
        online_sold_structure = {
            "online_sold_description": report.get("description"),
            "online_sold_price": report.get("price"),
            "online_sold_quantity": report.get("quantity"),
            "online_sold_amount": report.get("amount")
        }
        data.append(online_sold_structure)

    for subtotal in subtotals:
        online_sold_structure = {
            "online_sold_description": subtotal.get("description"),
            "online_sold_price": subtotal.get("price"),
            "online_sold_quantity": subtotal.get("quantity"),
            "online_sold_amount": subtotal.get("amount")
        }
        data.append(online_sold_structure)

    return data


def less_online_reedemed(section):
    data = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")

    for report in reports:
        less_online_reedemed_strcture = {
            "Less_online_redeemed_description": report.get("description"),
            "Less_online_redeemed_price": report.get("price"),
            "Less_online_redeemed_quantity": report.get("quantity"),
            "Less_online_redeemed_amount": report.get("amount")
        }
        data.append(less_online_reedemed_strcture)
    for subtotal in subtotals:
        less_online_reedemed_strcture = {
            "Less_online_redeemed_description": subtotal.get("description"),
            "Less_online_redeemed_price": subtotal.get("price"),
            "Less_online_redeemed_quantity": subtotal.get("quantity"),
            "Less_online_redeemed_amount": subtotal.get("amount")
        }
        data.append(less_online_reedemed_strcture)
    return data


def free_wash_issued(section):
    data = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    for report in reports:
        free_wash_issued_structure = {
            "Free_washes_issued_description": report.get("description"),
            "Free_washes_issued_price": report.get("price"),
            "Free_washes_issued_quantity": report.get("quantity"),
            "Free_washes_issued_amount": report.get("amount")
        }

        data.append(free_wash_issued_structure)
    for subtotal in subtotals:
        free_wash_issued_structure = {
            "Free_washes_issued_description": subtotal.get("description"),
            "Free_washes_issued_price": subtotal.get("price"),
            "Free_washes_issued_quantity": subtotal.get("quantity"),
            "Free_washes_issued_amount": subtotal.get("amount")
        }

        data.append(free_wash_issued_structure)

    return data


def less_paidouts(section):
    data = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")

    for report in reports:
        less_paidouts_structure = {
            "Less_paidouts_description": report.get("description"),
            "Less_paidouts_price": report.get("price"),
            "Less_paidouts_quantity": report.get("quantity"),
            "Less_paidouts_amount": report.get("amount")
        }
        data.append(less_paidouts_structure)
    for total in subtotals:
        less_paidouts_structure = {
            "Less_paidouts_description": total.get("description"),
            "Less_paidouts_price": total.get("price"),
            "Less_paidouts_quantity": total.get("quantity"),
            "Less_paidouts_amount": total.get("amount")
        }
        data.append(less_paidouts_structure)

    return data


def total_to_account_for(section):

    return section.get("totalAmount", 0.0)


def deposits(section):
    data = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    for report in reports:
        deposit_structure = {
            "Deposits_description": report.get("description"),
            "Deposits_price": report.get("price"),
            "Deposits_quantity": report.get("quantity"),
            "Deposits_amount": report.get("amount")
        }
        data.append(deposit_structure)

    for subtotal in subtotals:
        deposit_structure = {
            "Deposits_description": subtotal.get("description"),
            "Deposits_price": subtotal.get("price"),
            "Deposits_quantity": subtotal.get("quantity"),
            "Deposits_amount": subtotal.get("amount")
        }
        data.append(deposit_structure)

    return data


def total_xpt_cash(section):
    data = []
    total_xpt_structure = {
        "TOTAL XPT CASH:": section.get("totalAmount")
    }
    data.append(total_xpt_structure)

    return data


def house_accounts(section):
    data = []
    reports = section.get('reports')
    subtotals = section.get('subtotals')
    for report in reports:
        house_accounts_structure = {
            "House_accounts_description": report.get("description"),
            "House_accounts_price": report.get("price"),
            "House_accounts_quantity": report.get("quantity"),
            "House_accounts_amount": report.get("amount")
        }
        data.append(house_accounts_structure)

    for subtotal in subtotals:
        house_accounts_structure = {
            "House_accounts_description": subtotal.get("description"),
            "House_accounts_price": subtotal.get("price"),
            "House_accounts_quantity": subtotal.get("quantity"),
            "House_accounts_amount": subtotal.get("amount")
        }
        data.append(house_accounts_structure)

    return data


def cash(section):
    data = []
    cash_structure = {
        "CASH:": section.get("totalAmount")
    }
    data.append(cash_structure)

    return data


def xpt_acceptors(section):
    data = []
    xpt_acceptor_structure = {
        "XPT ACCEPTORS:": section.get("totalAmount")
    }

    data.append(xpt_acceptor_structure)

    return data


def xpt_dispensers(section):
    data = []
    xpt_dispenser_structure = {
        "XPT DISPENSERS:": section.get("totalAmount")
    }

    data.append(xpt_dispenser_structure)

    return data


def total_function(section):
    data = []
    total_structure = {
        "TOTAL:": section.get("totalAmount")
    }
    data.append(total_structure)

    return data


def credit_card(section):
    data = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")

    for report in reports:
        credit_card_structure = {
            "Credit_Card_description": report.get("description"),
            "Credit_Card_price": report.get("price"),
            "Credit_Card_quantity": report.get("quantity"),
            "Credit_Card_amount": report.get("amount")
        }
        data.append(credit_card_structure)

    for subtotal in subtotals:
        credit_card_structure = {
            "Credit_Card_description": subtotal.get("description"),
            "Credit_Card_price": subtotal.get("price"),
            "Credit_Card_quantity": subtotal.get("quantity"),
            "Credit_Card_amount": subtotal.get("amount")
        }
        data.append(credit_card_structure)
    return data


def other_tenders(section):
    data = []
    reports = section.get("reports")
    subtotals = section.get("subtotals")
    for report in reports:
        other_tender_structure = {
            "Other_tenders_description": report.get("description"),
            "Other_tenders_price": report.get("price"),
            "Other_tenders_quantity": report.get("quantity"),
            "Other_tenders_amount": report.get("amount")
        }
        data.append(other_tender_structure)
    for subtotal in subtotals:
        other_tender_structure = {
            "Other_tenders_description": subtotal.get("description"),
            "Other_tenders_price": subtotal.get("price"),
            "Other_tenders_quantity": subtotal.get("quantity"),
            "Other_tenders_amount": subtotal.get("amount")
        }
        data.append(other_tender_structure)

    return data


def xpt_balancing(section):
    data = []
    xpt_balancing_structure = {
        "XPT BALANCING:": section.get("totalAmount")
    }
    data.append(xpt_balancing_structure)

    return data


def report_balance(section):
    data = []
    report_balance_structure = {
        "REPORT BALANCE:": section.get("totalAmount")
    }

    data.append(report_balance_structure)

    return data


def picture_mismatch(section):
    data = []
    picture_mismatch_structure = {
        "PICTURE MISMATCH:": section.get("totalCount")
    }
    data.append(picture_mismatch_structure)

    return data


def report_data_extractor(report_data):
    single_site_report = {}
    if report_data:
        gsviews = report_data.get("gsviews")
        gsviews_0 = gsviews[0]

        sections = gsviews_0.get("sections")
        for section in sections:
            text = section.get("text")

            if text == "WASH SALES-":
                wash_sales_ret = wash_sales(section)  # file_path
                print(f"Washsales:{wash_sales_ret}")
                single_site_report["car_count"] = wash_sales_ret  # car count

            elif text == "WASH PACKAGES-":
                wash_packages(section)

            elif text == "WASH EXTRA SERVICES-":
                wash_extra_services(section)

            elif text == "GROSS WASH SALES-":
                gross_wash_sales(section)

            elif text == "LESS FREE WASH RDMD-":
                less_free_wash_rdmd(section)

            elif text == "LESS WASH DISCOUNTS-":
                less_wash_discounts(section)

            elif text == "LESS LOYALTY DISC-":
                less_loyality_disc(section)

            elif text == "NET SITE SALES:":
                net_site_sales_value = net_site_sales(section)
                single_site_report['net_sales'] = net_site_sales_value

            elif text == "ARM PLANS SOLD-":
                arm_plans_sold_cnt = arm_plans_sold(section)
                single_site_report["arm_plans_sold_cnt"] = arm_plans_sold_cnt

            elif text == "ARM PLANS RECHARGED-":
                arm_reachrged_amt = arm_plans_recharged(section)
                print("arm plans rechanged:", arm_reachrged_amt)
                single_site_report["arm_plans_recharged_amt"] = arm_reachrged_amt

            elif text == "ARM PLANS REDEEMED-":
                arm_plans_reedemed_value = arm_planes_reedemed(section)
                single_site_report["arm_plans_reedemed_cnt"] = arm_plans_reedemed_value.get("arm_plans_reedemed_cnt")
                single_site_report["arm_plans_reedemed_amt"] = arm_plans_reedemed_value.get("arm_plans_reedemed_amt")

            elif text == "ARM PLANS TERMINATED-":
                arm_plans_terminated(section)

            elif text == "PREPAIDS SOLD-":
                prepaid_sold(section)

            elif text == "LESS PREPAIDS REDEEMED-":
                less_prepaid_reedemed(section)

            elif text == "ONLINE SOLD-":
                online_sold(section)

            elif text == "LESS ONLINE REDEEMED-":
                less_online_reedemed(section)

            elif text == "FREE WASHES ISSUED-":
                free_wash_issued(section)

            elif text == "LESS PAIDOUTS:":
                less_paidouts(section)

            elif text == "TOTAL TO ACCOUNT FOR:":
                total_revenue_val = total_to_account_for(section)
                single_site_report['total_revenue'] = total_revenue_val

            elif text == "DEPOSITS-":
                deposits(section)

            elif text == "TOTAL XPT CASH:":
                total_xpt_cash(section)

            elif text == "HOUSE ACCOUNTS-":
                house_accounts(section)

            elif text == "CASH:":
                cash(section)

            elif text == "XPT ACCEPTORS:":
                xpt_acceptors(section)

            elif text == "XPT DISPENSERS:":
                xpt_dispensers(section)

            elif text == "TOTAL:":
                total_function(section)

            elif text == "CREDIT CARD:":
                credit_card(section)

            elif text == "OTHER TENDERS:":
                other_tenders(section)

            elif text == "XPT BALANCING:":
                xpt_balancing(section)

            elif text == "REPORT BALANCE:":
                report_balance(section)

            elif text == "PICTURE MISMATCH:":
                picture_mismatch(section)

    return single_site_report


def generate_report(path, start_date_current_year, end_date_current_year, start_date_last_year, end_date_last_year):
    logger = logging.getLogger(__name__)
    logger.info("started main script")
    site_watch_report = {}

    is_location_code_success = False
    success_location_code = None, None
    location_codes_a = ['SPKLUS-001', 'SPKLUS-002', 'SPKLUS-003', 'SPKLUS-004',
                        'SPKLUS-005', 'SPKLUS-006', 'SPKLUS-007', 'SPKLUS-008',
                        'SPKLUS-009']
    location_codes_b = ['SPKLUS-012', 'SPKLUS-013', 'SPKLUS-014', 'SPKLUS-015']

    for index, site in sites_df.iterrows():
        # all_dictionaries_lst = []
        site_dict = site.to_dict()
        slno = site_dict.get("slno")
        locationCode = site_dict.get("Organization")
        location_retry = 0
        while True:

            if location_retry >= MAX_RETRIES:
                logger.info(f"Max retries reached givingup on retry for location : {site_dict}")
                break

            try:
                combined_data = {}

                cookiesfile_name = f"{(site_dict.get('Organization')).strip().replace('-','_')}.pkl"
                # print(cookiesfile_name)
                print(site_dict)
                logging.info(f"{site_dict}")

                cookies_file = os.path.join(cookies_path, cookiesfile_name)

                client = sitewatchClient(cookies_file=cookies_file)
                employCode = site_dict.get("employee")
                password = site_dict.get("password")

                if is_location_code_success:
                    locationCode_old, slno_old = success_location_code
                    location_a_range = range(1, 10)
                    location_b_range = range(10, 14)
                    if slno_old in location_a_range and slno in location_a_range:
                        locationCode = locationCode_old

                    elif slno_old in location_b_range and slno in location_b_range:
                        locationCode = locationCode_old

                print(f"\n location code used :{locationCode}")
                logger.info(f"location code used :{locationCode}")
                client_name = site_dict.get("client_name2")
                session_chek = client.check_session_auth(timeout=60)

                if not session_chek:
                    token = client.login(employeeCode=employCode, password=password, locationCode=locationCode, remember=1)
                    print(token)

                session_chek = client.check_session_auth(timeout=60)
                if session_chek:
                    reportOn = site_dict.get("site")
                    id = site_dict.get("id")
                    idname = site_dict.get("id_name")
                    request_id1 = client.get_general_sales_report_request_id(reportOn, id, idname, start_date_current_year, end_date_current_year)
                    request_id1_1 = client.get_activity_by_date_proft_request_id(reportOn, start_date_current_year, end_date_current_year)  # for labour hours

                    if request_id1 and request_id1_1:
                        report_data = client.get_report(reportOn,request_id1)
                        # print(f"report data: {report_data}")
                        extracted_data1= report_data_extractor(report_data)
                        car_count_current_year = extracted_data1.get("car_count",0)
                        arm_plans_reedemed_current_year_cnt = extracted_data1.get("arm_plans_reedemed_cnt",0)
                        arm_plans_reedemed_current_year_amt = abs(extracted_data1.get("arm_plans_reedemed_amt",0))#need abs because will nee to mek it postive
                        arm_plans_recharged_amt_current_year_amt =extracted_data1.get("arm_plans_recharged_amt",0)

                        retail_car_count_current_year = (car_count_current_year - arm_plans_reedemed_current_year_cnt)

                        net_sales_amt = extracted_data1.get("net_sales",0.0)
                        total_revenue_current_year = round(extracted_data1.get("total_revenue",0.0),2)

                        if client_name=="Sudz - Beverly":
                            retail_revenue__current_year = round((total_revenue_current_year - arm_plans_recharged_amt_current_year_amt),2)
                        else:
                            retail_revenue__current_year = round((net_sales_amt - arm_plans_reedemed_current_year_amt),2)

                        arm_plans_sold_total_cnt_current_year = extracted_data1.get("arm_plans_sold_cnt")
                        labour_hours_current_year=client.get_labour_hours(reportOn,request_id1_1)
                        cars_per_labour_hour_current_year = round((car_count_current_year/labour_hours_current_year),2) if labour_hours_current_year !=0 else ""

                        current_year_data = {
                            "car_count_current_year":car_count_current_year,
                            "arm_plans_reedemed_current_year_cnt":arm_plans_reedemed_current_year_cnt,
                            "retail_car_count_current_year":retail_car_count_current_year,
                            "retail_revenue_current_year":retail_revenue__current_year,
                            "total_revenue_current_year":total_revenue_current_year,
                            "labour_hours_current_year":labour_hours_current_year,
                            "cars_per_labour_hour_current_year":cars_per_labour_hour_current_year,
                            "arm_plans_sold_cnt_current_year": arm_plans_sold_total_cnt_current_year
                        }
                        combined_data.update(current_year_data)

                    request_id2 = client.get_general_sales_report_request_id(reportOn,id,idname,start_date_last_year, end_date_last_year)
                    request_id2_2 =client.get_activity_by_date_proft_request_id(reportOn,start_date_last_year, end_date_last_year) #for labour hours

                    if request_id2 and request_id2_2:
                        report_data = client.get_report(reportOn,request_id2)
                        # print(f"data2:{report_data}")
                        extracted_data2= report_data_extractor(report_data)
                        car_count_last_year = extracted_data2.get("car_count",0)
                        arm_plans_reedemed_last_year_cnt = extracted_data2.get("arm_plans_reedemed_cnt",0)
                        arm_plans_reedemed_last_year_amt = abs(extracted_data2.get("arm_plans_reedemed_amt",0))
                        arm_plans_recharged_amt_last_year_amt =extracted_data2.get("arm_plans_recharged_amt",0)

                        retail_car_count_last_year =(car_count_last_year- arm_plans_reedemed_last_year_cnt)

                        net_sales_amt2= extracted_data2.get("net_sales",0.0)
                        total_revenue_last_year = round(extracted_data2.get("total_revenue",0.0),2)

                        if client_name=="Sudz - Beverly":
                            retail_revenue__last_year = round((total_revenue_last_year - arm_plans_recharged_amt_last_year_amt),2)
                        else:
                            retail_revenue__last_year = round((net_sales_amt2- arm_plans_reedemed_last_year_amt),2)

                        arm_plans_sold_total_cnt_last_year = extracted_data2.get("arm_plans_sold_cnt")

                        labour_hours_last_year=client.get_labour_hours(reportOn,request_id2_2)
                        cars_per_labour_hour_last_year = round((car_count_last_year/labour_hours_last_year),2) if labour_hours_last_year !=0 else ""

                        last_year_data = {
                            "car_count_last_year":car_count_last_year,
                            "arm_plans_reedemed_last_year":arm_plans_reedemed_last_year_cnt,
                            "retail_car_count_last_year":retail_car_count_last_year,
                            "retail_revenue_last_year":retail_revenue__last_year,
                            "total_revenue_last_year":total_revenue_last_year,
                            "labour_hours_last_year":labour_hours_last_year,
                            "cars_per_labour_hour_last_year":cars_per_labour_hour_last_year,
                            "arm_plans_sold_cnt_last_year": arm_plans_sold_total_cnt_last_year
                        }

                        combined_data.update(last_year_data)

                        conversion_rate_current_year  = round((arm_plans_sold_total_cnt_current_year/retail_car_count_current_year)*100,2) if retail_car_count_current_year !=0 else ""
                        conversion_rate_last_year  = round((arm_plans_sold_total_cnt_last_year/retail_car_count_last_year)*100,2) if retail_car_count_last_year !=0 else ""

                        total_arm_planmembers_cnt_current_year = get_total_arm_plan_members(client,reportOn,end_date_current_year) # parent year end date
                        total_arm_planmembers_cnt_last_year = get_total_arm_plan_members(client,reportOn,end_date_last_year) # parent year end date

                        combined_data["total_arm_planmembers_cnt_current_year"] = total_arm_planmembers_cnt_current_year
                        combined_data["total_arm_planmembers_cnt_last_year"] = total_arm_planmembers_cnt_last_year
                        combined_data["conversion_rate_current_year"]= conversion_rate_current_year
                        combined_data["conversion_rate_last_year"]= conversion_rate_last_year


                    print(f"combined data:{combined_data}")
                    logger.info(f"combined data:{combined_data}")
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
                    logger.info("switching lcoation code")
                    if slno in location_a_range:
                        locationCode=random.choice(location_codes_a)
                        print(f"\n retrying with new location code {locationCode} for {slno}")
                        logger.info(f"retrying with new location code {locationCode} for {slno}")
                        success_location_code = locationCode,slno

                    elif  slno in location_b_range:
                        locationCode=random.choice(location_codes_b)
                        success_location_code = locationCode,slno
                        print(f"\n retrying with new location code {locationCode} for {slno}")
                        logger.info(f"retrying with new location code {locationCode} for {slno}")

                    print("sleep for 30 secound before next retry")
                    logger.info("sleep for 30 secound before next retry")
                    time.sleep(30)

            except Exception as e:
                print(f"Excetion for this loctaion {client_name} {e} {traceback.print_exc() }")
                logger.info(f"Excetion for this loctaion {client_name} {e} {traceback.print_exc() }")

            location_retry+=1

    return site_watch_report


if __name__=="__main__":

    start_date_current_year="2024-11-01"
    end_date_current_year="2024-11-14"

    start_date_last_year="2023-11-01"
    end_date_last_year="2023-11-14"

    report = generate_report("",start_date_current_year,end_date_current_year,start_date_last_year, end_date_last_year)
    print("\n"*6)
    print(report)
    with open("sitewatch_report_full.json","w") as f:
        json.dump(report,f,indent=4)
