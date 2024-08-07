from datetime import datetime, timedelta


def generate_weekdays_historical(start_year, start_month, start_day, end_year, end_month, end_day):
    weekdays = {
        0: "Monday",
        4: "Friday",
        5: "Saturday",
        6: "Sunday"
    }
    days = []
    current_date = datetime(start_year, start_month, start_day)
    week = []

    while current_date <= datetime(end_year, end_month, end_day):
        if current_date.weekday() in weekdays:
            week.append(current_date)
        if current_date.weekday() == 6:  # End of the week
            if week and len(week) == 4:  # Add the non-empty week to the list if week list has 4 days that we need (Mon, Fri, Sat, Sun)
                days.append(week)
            week = []  # Reset for the next week
        current_date += timedelta(days=1)
        
    if week and len(week) == 4:  # Add the last week if it's not empty and has 4 days
        days.append(week)
    
    return days

def history_format_date_sitewatch(dates):
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
        

        return tuple([date.strftime("%Y-%m-%d") for date in dates])
        

def history_format_date_hamilton(dates):
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
        return tuple([date.strftime("%Y-%m-%d") for date in dates])

def history_format_date_washify(dates):
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
        return tuple([date.strftime("%m/%d/%Y") for date in dates])


if __name__=="__main__":
    start_year = 2022
    start_month = 7
    start_day = 3
    end_year = 2024
    end_month = 7
    end_day = 28
    result = generate_weekdays_historical(start_year, start_month, start_day, end_year, end_month, end_day)
    # for week in result:
    #     print(week)

    print(len(result))
    # print(type(result[0][0]))
    # print(result)
    
    for one_week in result:
        print(history_format_date_sitewatch(one_week))

