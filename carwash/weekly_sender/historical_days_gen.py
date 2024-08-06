from datetime import datetime, timedelta


def generate_weekdays_historical(year, start_month,start_day,end_year,end_month):
    weekdays = {
        0: "Monday",
        4: "Friday",
        5: "Saturday",
        6: "Sunday"
    }
    days = []
    current_date = datetime(year, start_month, start_day)
    week = []

    while (current_date.year in [2022,2023,2024]) and  ( not (current_date.year==end_year and current_date.month==end_month)) :
        if current_date.weekday() in weekdays:
            week.append(current_date) #.strftime("%Y-%m-%d %A")
        if current_date.weekday() == 6:  # End of the week
            if week and len(week)==4:  # Add the non-empty week to the list if week list ahs 4 days taht we need mon,fri,sat,sun
                days.append(week)
            week = []  # Reset for the next week
        current_date += timedelta(days=1)
        
    if week:  # Add the last week if it's not empty
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
    year=2022
    start_month =5
    start_day = 29
    end_year=2024
    end_month = 7
    result = generate_weekdays_historical(year, start_month,start_day,end_year,end_month)
    # for week in result:
    #     print(week)

    print(len(result))
    # print(type(result[0][0]))
    # print(result)
    
    for one_week in result[:2]:
        print(history_format_date_sitewatch(one_week))

