from washify import washifyClient
import json
import traceback
import logging


username = 'Cameron'
password = 'Password1'
companyName = 'cleangetawayexpress'
userType = 'CWA'


def conversion_rate_washify(arm_plans_sold,retail_car_count):
    rate = 0.0

    try:
        rate = arm_plans_sold/retail_car_count
        rate= rate*100
        rate = round(rate,2)

    except Exception as e:
        print(f"Exception in conversion_rate_washify() {e}")
    return rate


def generate_report(file_path, start_date_current_year,end_date_current_year,start_date_last_year, end_date_last_year):
    logger = logging.getLogger(__name__)
    logger.info("started main washify")
    "This will generate weekly report"
    final_report = {}

    try:
        client  = washifyClient()
        is_logged_in = client.check_login(proxy=None)

        if not is_logged_in:
            login = client.login(username=username,password=password,companyName=companyName,userType=userType)
            print(f"doing relogin : {login}")
            logger.info(f"doing relogin : {login}")
        client_locations = client.get_user_locations()

        # client_locations_number_codes =list(client_locations.values())
        print(f"client lcoations {client_locations.items()}")
        logger.info(f"client lcoations {client_locations.items()}")

        if client_locations:

            for location_name,location_code in client_locations.items():
                single_site_report = {}
                print(location_code,location_name)
                ## -----------current year start  to ---- end  ----------------##
                car_count_report_current_year_report = client.get_car_count_report([location_code],start_date_current_year,end_date_current_year)
                retail_revenue_summary_report_current_year = client.GetRevenuReportFinancialRevenueSummary([location_code],start_date_current_year,end_date_current_year)
                total_arm_plans1 = client.GetRevenuReportFinancialUnlimitedSales([location_code],start_date_current_year,end_date_current_year)
                retail_revenue_current_year = retail_revenue_summary_report_current_year.get("netPrice",0)
                total_revenue_current_year = retail_revenue_summary_report_current_year.get("total",0.0)
                labour_hours_current_year = car_count_report_current_year_report.get("totalhrs")

                car_count_current_year_cnt = car_count_report_current_year_report.get("car_count")
                print(car_count_report_current_year_report)
                print("retail revenue  report :",retail_revenue_summary_report_current_year)

                cars_per_labour_hour_current_year = round((car_count_current_year_cnt/labour_hours_current_year),2) if labour_hours_current_year !=0 else ""

                retail_car_count_current_year = car_count_report_current_year_report.get("retail_car_count")

                single_site_report["car_count_current_year"]=car_count_current_year_cnt

                #single_site_report["arm_plans_reedemed_monday_to_friday_cnt"] = ""  #update
                single_site_report["retail_car_count_current_year"] = retail_car_count_current_year
                single_site_report["retail_revenue_current_year"] = retail_revenue_current_year
                single_site_report["total_revenue_current_year"] = total_revenue_current_year
                single_site_report["labour_hours_current_year"] = labour_hours_current_year
                single_site_report["cars_per_labour_hour_current_year"] = cars_per_labour_hour_current_year

                ## -----------Last year start  to ---- end  ----------------##

                car_count_report_last_year_report = client.get_car_count_report([location_code],start_date_last_year,end_date_last_year)
                retail_revenue_summary_report_last_year = client.GetRevenuReportFinancialRevenueSummary([location_code],start_date_last_year,end_date_last_year)
                total_arm_plans2 = client.GetRevenuReportFinancialUnlimitedSales([location_code],start_date_last_year,end_date_last_year)
                retail_revenue_last_year = retail_revenue_summary_report_last_year.get("netPrice",0)
                total_revenue_last_year = retail_revenue_summary_report_last_year.get("total",0.0)
                #print(car_count_report_sat_sun_report)
                car_count_last_year_cnt = car_count_report_last_year_report.get("car_count",0)
                labour_hours_last_year = car_count_report_last_year_report.get("totalhrs")

                cars_per_labour_hour_last_year = round((car_count_last_year_cnt/labour_hours_last_year),2) if labour_hours_last_year != 0 else ""

                print("retail revenue  report :",retail_revenue_summary_report_last_year)

                retail_car_count_last_year = car_count_report_last_year_report.get("retail_car_count")

                single_site_report["car_count_last_year"]=car_count_report_last_year_report.get("car_count")
                #single_site_report["arm_plans_reedemed_saturday_sunday"] = "" #update
                single_site_report["retail_car_count_last_year"] = retail_car_count_last_year
                single_site_report["retail_revenue_last_year"] = retail_revenue_last_year
                single_site_report["total_revenue_last_year"] = total_revenue_last_year
                single_site_report["labour_hours_last_year"] = labour_hours_last_year
                single_site_report["cars_per_labour_hour_last_year"] = cars_per_labour_hour_last_year

                total_arm_planmembers_cnt = client.get_club_plan_members(location_code)

                single_site_report["arm_plans_sold_cnt_current_year"] = total_arm_plans1
                single_site_report["arm_plans_sold_cnt_last_year"] = total_arm_plans2
                single_site_report["total_arm_planmembers_cnt_current_year"] = total_arm_planmembers_cnt
                single_site_report["conversion_rate_current_year"] = conversion_rate_washify(total_arm_plans1,retail_car_count_current_year)
                single_site_report["conversion_rate_last_year"] = conversion_rate_washify(total_arm_plans2,retail_car_count_last_year)

                if "1631" in location_name: # 1631 E Jackson St
                    final_report["Getaway-Macomb"] = single_site_report
                elif "1821" in location_name:
                    final_report["Getaway-Morton"]=  single_site_report
                elif "2950" in location_name:
                    final_report["Getaway-Ottawa"] = single_site_report
                elif "4234" in location_name:
                    final_report["Getaway-Peru"]   = single_site_report

    except Exception as e:
        print(f"Exception generate_weeklyrepoer washify {e},{traceback.format_exc()}")
        logger.info(f"Exception generate_weeklyrepoer washify {e},{traceback.format_exc()}")
    logger.info("sitewatch final data")
    logger.info(f"{final_report}")

    return final_report


if __name__=="__main__":

    start_date_current_year=  "11/01/2024"
    end_date_current_year=  "11/14/2024"

    start_date_last_year=  "11/01/2023"
    end_date_last_year=  "11/14/2023"

    data = generate_report('', start_date_current_year,end_date_current_year,start_date_last_year, end_date_last_year)

    print(data)

    with open("test_washify_data.json","w") as f:
        json.dump(data,f,indent=4)
