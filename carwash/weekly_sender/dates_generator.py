from datetime import datetime
import calendar

def get_dates_for_current_year(dates=None):
    """ Generates start and end date of current year"""

    if dates:
        return tuple(datetime.strptime(date, '%Y-%m-%d') for date in dates)

    else:
        today = datetime.today()
        start_date = today.replace(day=1)

        today_day = today.day
        if today_day in (7, 14, 21,) or today_day == calendar.monthrange(today.year, today.month)[1]:
            end_date = today
        else:
            end_date = None

        return (start_date, end_date)


def get_dates_for_last_year(dates=None):
    """ Generates start and end date of last year"""

    if dates:
        return tuple(datetime.strptime(date, '%Y-%m-%d') for date in dates)

    else:
        today = datetime.today()
        start_date = today.replace(day=1).replace(year=today.year - 1)

        today_day = today.day

        if today_day in (7, 14, 21,) or today_day == calendar.monthrange(start_date.year, start_date.month)[1]:
            end_date = today.replace(year=today.year -1)
        else: end_date = None

    return (start_date, end_date)


def format_dates_sitewatch(dates):

    if isinstance(dates,tuple):
        return tuple(date.strftime("%Y-%m-%d") for date in dates)


def format_dates_hamilton(dates):

    if isinstance(dates,tuple):
        return tuple(date.strftime("%Y-%m-%d") for date in dates)


def format_dates_washify(dates):

    if isinstance(dates,tuple):
        formated_dates = tuple(date.strftime('%m/%d/%Y') for date in dates)

        return formated_dates


if __name__=="__main__":

    start_date_c_year="2024-11-01"
    end_date_c_year="2024-11-07"

    start_date_l_year="2023-11-01"
    end_date_l_year="2023-11-07"

    start_date_c, end_date_c = get_dates_for_current_year()
    start_date_l, end_date_l = get_dates_for_last_year()

    if end_date_c is not None:
        print(start_date_c, end_date_c)
        print(start_date_l, end_date_l)
        print(f'for washify dates : {format_dates_washify((start_date_c, end_date_c, start_date_l, end_date_l))}')
        print(f'for sitewatch dates : {format_dates_sitewatch((start_date_c, end_date_c, start_date_l, end_date_l))}')
        print(f'for hamilton dates : {format_dates_hamilton((start_date_c, end_date_c, start_date_l, end_date_l))}')
    else:
        print('No date')
