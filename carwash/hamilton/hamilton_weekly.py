from datetime import datetime, timedelta
from hamilton import hamiltonClient
import os


def get_week_dates():
    # Get the current date
    today = datetime.today()
    
    # Find the current week's Monday date
    current_week_monday = today - timedelta(days=today.weekday())
    
    # Find the current week's Sunday date
    current_week_sunday = current_week_monday + timedelta(days=6)
    
    # Format the dates in dd/mm/yyyy format
    monday_date_str = current_week_monday.strftime("%m/%d/%Y")
    sunday_date_str = current_week_sunday.strftime("%m/%d/%Y")
    
    return monday_date_str, sunday_date_str


def generate_weekly_report(file_path,monday_date_str, sunday_date_str):
    proxy_url = None
    proxy = {"http": proxy_url, "https": proxy_url}
    client = hamiltonClient()
    login_data = {"UserName": "CR@Sparklecw.com", "Password": "CameronRay1"}
    login = client.login(login_data, proxy)
    # print(f"login:{login}")
    # daily_report = client.get_daily_report(proxy)
    # print(f"daily :{daily_report}")
    rev = client.get_revenue(monday_date_str, sunday_date_str)
    # file_name_with_path = os.path.join(file_path,f"hamilton_{monday_date_str}_{sunday_date_str}.csv")
    client.to_csv(rev,file_path)

if __name__=="__main__":
    print(get_week_dates())