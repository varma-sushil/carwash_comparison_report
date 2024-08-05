from datetime import datetime, timedelta

def get_week_dates_for_current(monday_str=None):
    """
    Get the dates for the Monday, Friday, Saturday, and Sunday of the week 
    containing the specified date. If no date is specified, use today's date. Y-m-d
    
    Parameters:
    - monday_str (str): Date string in "Y-m-d" format. Optional.
    
    Returns:
    - tuple: Dates for the current week's Monday, Friday, Saturday, and Sunday in yyyy-mm-dd format.
    """
    if monday_str:
        today = datetime.strptime(monday_str, "%Y-%m-%d")
    else:
        today = datetime.today()
    
    # Find the current week's Monday date
    current_week_monday = today - timedelta(days=today.weekday())
    
    # Find the current week's Friday, Saturday, and Sunday dates
    current_week_friday = current_week_monday + timedelta(days=4)
    current_week_saturday = current_week_monday + timedelta(days=5)
    current_week_sunday = current_week_monday + timedelta(days=6)
    
    # Format the dates in yyyy-mm-dd format
    return (current_week_monday,
            current_week_friday,
            current_week_saturday,
            current_week_sunday)

def generate_past_4_week_days_full(monday_str):
    "this will generat mon,fri,sat stunda list of dates or all 4 weeks "
    # Convert the string date to a datetime object
    if monday_str:
        input_date = datetime.strptime(monday_str, "%Y-%m-%d")
    
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
            required_days.append(current_date)
        current_date += timedelta(days=1)

    # Print the result

    full_days = [required_days[i:i + 4] for i in range(0, len(required_days), 4)]
    
    #     print(day)

    return full_days

def format_date_sitewatch(dates):
    """
    Format a tuple of dates in the format 'yyyy-mm-dd'.
    
    Parameters:
    - dates (tuple): A tuple of datetime objects.
    
    Returns:
    - tuple: A tuple of formatted date strings.
    """
    if isinstance(dates,tuple):
        return tuple(date.strftime("%Y-%m-%d") for date in dates)
    else:
        fmt_days =[]
        for week in dates:
            fmt_days.append([date.strftime("%Y-%m-%d") for date in week])
        return fmt_days

def format_date_hamilton(dates):
    """
    Format a tuple of dates in the format 'yyyy-mm-dd'.
    
    Parameters:
    - dates (tuple): A tuple of datetime objects.
    
    Returns:
    - tuple: A tuple of formatted date strings.
    """

    if isinstance(dates,tuple):
        return tuple(date.strftime("%Y-%m-%d") for date in dates)
    else:
        fmt_days =[]
        for week in dates:
            fmt_days.append([date.strftime("%Y-%m-%d") for date in week])
        return fmt_days

def format_date_washify(dates):
    """
    Format a tuple of dates in the format 'yyyy-mm-dd'.
    
    Parameters:
    - dates (tuple): A tuple of datetime objects.
    
    Returns:
    - tuple: A tuple of formatted date strings.
    """

    if isinstance(dates,tuple):
        return tuple(date.strftime("%m/%d/%Y") for date in dates)
    else:
        fmt_days =[]
        for week in dates:
            fmt_days.append([date.strftime("%m/%d/%Y") for date in week])
        return fmt_days



if __name__=="__main__":
    print(get_week_dates_for_current())  # Use today's date
    dates = get_week_dates_for_current("2024-07-31")  # Use the specified date
    print(type(dates))
    fmt_dates = format_date_sitewatch(dates)
    past_4_weeks_days_full = generate_past_4_week_days_full(fmt_dates[0])
    print(format_date_washify(past_4_weeks_days_full))