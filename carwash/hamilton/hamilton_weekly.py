from datetime import datetime, timedelta
import os


def get_week_dates():
    # Get the current date
    today = datetime.today()
    
    # Find the current week's Monday date
    current_week_monday = today - timedelta(days=today.weekday())
    
    # Find the current week's Friday, Saturday, and Sunday dates
    current_week_friday = current_week_monday + timedelta(days=4)
    current_week_saturday = current_week_monday + timedelta(days=5)
    current_week_sunday = current_week_monday + timedelta(days=6)
    
    # Format the dates in mm/dd/yyyy format
    monday_date_str = current_week_monday.strftime("%Y-%m-%d")
    friday_date_str = current_week_friday.strftime("%Y-%m-%d")
    saturday_date_str = current_week_saturday.strftime("%Y-%m-%d")
    sunday_date_str = current_week_sunday.strftime("%Y-%m-%d")
    
    return monday_date_str, friday_date_str, saturday_date_str, sunday_date_str

def generate_past_4_weeks_days(date_str):
    # Convert the string date to a datetime object
    date_format = "%Y-%m-%d"
    monday = datetime.strptime(date_str, date_format)
    
    # Subtract one day
    one_day_before = monday - timedelta(days=1)
    four_weeks_before = monday - timedelta(days=(7*4))

    # Format the dates in "dd/mm/yyyy" format
    formatted_date = one_day_before.strftime("%Y-%m-%d")
    four_weeks_before_fmt = four_weeks_before.strftime("%Y-%m-%d")

    print("One day before the current date:", formatted_date)
    print("4 weeks before day :", four_weeks_before_fmt)

    return four_weeks_before_fmt, formatted_date

def generate_past_4_week_days_full(mondaystr):
    "this will generat mon,fri,sat stunda list of dates or all 4 weeks "
    # Convert the string date to a datetime object
    date_format = "%Y-%m-%d"
    input_date = datetime.strptime(mondaystr, date_format)
    
    # Calculate one day before the input date and four weeks (28 days) before the input date
    one_day_before = input_date - timedelta(days=1)
    four_weeks_before = input_date - timedelta(days=7*4)

    # Initialize a list to store the required days
    required_days = []

    # Iterate through the range of dates
    current_date = four_weeks_before
    while current_date <= one_day_before:
        # Check if the current date is a Monday, Friday, Saturday, or Sunday
        if current_date.weekday() in [0, 4, 5, 6]:  # 0=Monday, 4=Friday, 5=Saturday, 6=Sunday
            required_days.append(current_date.strftime(date_format))
        current_date += timedelta(days=1)

    # Print the result

    full_days = [required_days[i:i + 4] for i in range(0, len(required_days), 4)]
    
    #     print(day)

    return full_days

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