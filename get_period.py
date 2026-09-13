import datetime
from datetime import timedelta


def get_week_period() -> list:
    my_dt = datetime.date.today()
    my_dt_trunc = datetime.date(my_dt.year,  # Truncate time component
                                my_dt.month,
                                my_dt.day)
    this_week = my_dt.isocalendar()[1]
    last_week = this_week - 1
    start_of_week = my_dt_trunc - timedelta(days=my_dt_trunc.weekday())  # timedelta & weekday

    end_of_week = start_of_week + timedelta(days=6)  # timedelta function

    last_start_week = start_of_week + timedelta(days=-7)
    last_start_week_string = last_start_week.strftime('%m/%d/%Y')
    last_end_week = end_of_week + + timedelta(days=-7)
    last_end_week_string = last_end_week.strftime('%m/%d/%Y')
    last_week_period = last_start_week_string + ' - ' + last_end_week_string
    return last_week, last_week_period


